'''''' 模块: VBA/错误查找  类型: 错误查找.bas  行数: 159
Attribute VB_Name = "错误查找"
Option Explicit

' ==========================================
' 宏名称：Workbook_Error_Scanner_V11.2 (单页/全表双模式版)
' 功能：遍历工作簿，定位错误值
' 更新：增加了【扫描当前工作表】与【扫描全部】的选择功能
' ==========================================

Sub Workbook_Error_Scanner_V11()
    Dim ws As Worksheet
    Dim rngErrors As Range, rngFormulas As Range, rngConstants As Range
    Dim cell As Range
    Dim sheetErrorCount As Long, totalErrorSheets As Long
    Dim userChoice As VbMsgBoxResult    ' 用于单个错误的跳转确认
    Dim scanScope As VbMsgBoxResult     ' 用于记录扫描范围的选择
    Dim cellChoice As VbMsgBoxResult
    Dim errorMsg As String
    Dim i As Long
    
    ' =======================================================
    ' 【核心修复】: 锁定用户当前的工作簿
    ' =======================================================
    Dim targetWB As Workbook
    Dim tempWS As Worksheet
    
    ' 尝试获取当前活动的窗口
    On Error Resume Next
    Set tempWS = Application.ActiveWindow.ActiveSheet
    On Error GoTo 0
    
    If tempWS Is Nothing Then Exit Sub
    
    ' 防止误操作插件本身
    If tempWS.Parent.Name = ThisWorkbook.Name Then
        MsgBox "请先点击一下你需要检查的数据表格，然后再点击按钮。", vbExclamation
        Exit Sub
    End If
    
    ' 锁定目标工作簿
    Set targetWB = tempWS.Parent
    ' =======================================================
    
    ' --- 新增：询问用户扫描范围 ---
    scanScope = MsgBox("请选择扫描范围：" & vbCrLf & vbCrLf & _
                       "● 点击【是 (Yes)】：扫描整个工作簿 (所有可见表)。" & vbCrLf & _
                       "● 点击【否 (No)】：仅扫描当前活动工作表 (【" & tempWS.Name & "】)。" & vbCrLf & _
                       "● 点击【取消】：退出程序。", _
                       vbYesNoCancel + vbQuestion, "选择扫描模式")
                       
    If scanScope = vbCancel Then Exit Sub
    
    totalErrorSheets = 0
    
    ' --- 1. 开始遍历工作表 ---
    For Each ws In targetWB.Worksheets
        
        ' 【逻辑判断】：
        ' 条件1：工作表必须可见
        ' 条件2：如果用户选了【否(只查单表)】，则当前ws必须等于tempWS，否则跳过
        If ws.Visible = xlSheetVisible Then
            
            ' 如果选了单表模式(vbNo)，且当前循环的表不是目标表，则直接跳过
            If scanScope = vbNo And ws.Name <> tempWS.Name Then GoTo NextSheet
            
            ' 清空对象，防止残留
            Set rngErrors = Nothing
            Set rngFormulas = Nothing
            Set rngConstants = Nothing
            
            ' --- 2. 高速查找当前表的错误 ---
            On Error Resume Next
            ' 在当前ws中查找
            Set rngFormulas = ws.Cells.SpecialCells(xlCellTypeFormulas, xlErrors)
            Set rngConstants = ws.Cells.SpecialCells(xlCellTypeConstants, xlErrors)
            On Error GoTo 0
            
            ' 合并错误范围
            If Not rngFormulas Is Nothing And Not rngConstants Is Nothing Then
                Set rngErrors = Union(rngFormulas, rngConstants)
            ElseIf Not rngFormulas Is Nothing Then
                Set rngErrors = rngFormulas
            ElseIf Not rngConstants Is Nothing Then
                Set rngErrors = rngConstants
            End If
            
            ' --- 3. 如果当前表发现错误 ---
            If Not rngErrors Is Nothing Then
                totalErrorSheets = totalErrorSheets + 1
                sheetErrorCount = rngErrors.Cells.count
                
                ' 激活工作表，让用户看到
                ws.Activate
                ' 选中所有错误，给个视觉概览
                rngErrors.Select
                
                ' --- 询问用户：是否要逐个查看？ ---
                userChoice = MsgBox("工作表 【" & ws.Name & "】 发现了 " & sheetErrorCount & " 个错误。" & vbCrLf & vbCrLf & _
                                    "● 点击【是 (Yes)】：逐个跳转定位修复。" & vbCrLf & _
                                    "● 点击【否 (No)】：跳过此表，继续检查。" & vbCrLf & _
                                    "● 点击【取消】：停止所有检查。", _
                                    vbYesNoCancel + vbExclamation, "发现错误")
                
                If userChoice = vbCancel Then
                    Exit Sub ' 彻底退出
                ElseIf userChoice = vbYes Then
                    ' --- 4. 进入单元格逐个跳转模式 ---
                    i = 0
                    For Each cell In rngErrors
                        i = i + 1
                        ' 选中并滚动到该单元格
                        cell.Activate
                        Application.GoTo Reference:=cell, Scroll:=True
                        
                        errorMsg = "正在检查: " & ws.Name & vbCrLf & _
                                   "位置: " & cell.Address(False, False) & vbCrLf & _
                                   "错误类型: " & GetErrorText_V11(cell.value) & vbCrLf & vbCrLf & _
                                   "进度: (" & i & " / " & sheetErrorCount & ")" & vbCrLf & vbCrLf & _
                                   "点击【是】继续下一个。" & vbCrLf & _
                                   "点击【否】不再查看此表的剩余错误。"
                                   
                        cellChoice = MsgBox(errorMsg, vbYesNo + vbQuestion, "错误定位中")
                        
                        If cellChoice = vbNo Then Exit For ' 退出当前表的循环
                    Next cell
                End If
            End If
        End If
        
NextSheet:
    Next ws
    
    ' --- 5. 结束提示 ---
    If totalErrorSheets = 0 Then
        If scanScope = vbNo Then
            MsgBox "完美！当前工作表 【" & tempWS.Name & "】 未发现任何错误。", vbInformation, "单表巡检完成"
        Else
            MsgBox "完美！在所有可见工作表中未发现任何错误。", vbInformation, "全表巡检完成"
        End If
    Else
        MsgBox "巡检结束！检查已完成。", vbInformation, "巡检完成"
    End If

End Sub

' 辅助函数：翻译错误代码 (保持不变)
Function GetErrorText_V11(errVal As Variant) As String
    Select Case errVal
        Case CVErr(xlErrDiv0): GetErrorText_V11 = "#DIV/0! (除以零)"
        Case CVErr(xlErrNA): GetErrorText_V11 = "#N/A (值不可用)"
        Case CVErr(xlErrName): GetErrorText_V11 = "#NAME? (名称无效)"
        Case CVErr(xlErrNull): GetErrorText_V11 = "#NULL! (空值)"
        Case CVErr(xlErrNum): GetErrorText_V11 = "#NUM! (数字无效)"
        Case CVErr(xlErrRef): GetErrorText_V11 = "#REF! (引用无效)"
        Case CVErr(xlErrValue): GetErrorText_V11 = "#VALUE! (类型错误)"
        Case Else: GetErrorText_V11 = "未知错误"
    End Select
End Function

