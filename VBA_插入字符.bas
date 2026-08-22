'''''' 模块: VBA/插入字符  类型: 插入字符.bas  行数: 237
Attribute VB_Name = "插入字符"
Option Explicit

' ==========================================
' 宏名称：BatchInsert_Universal_v5
' 功能：批量插入字符 (支持关键词定位 + 固定位置索引)
' 升级：输入 0=开头, -1=结尾, N=第N位, 文本=关键词查找
' ==========================================

Sub BatchInsert_Universal_v5()
    Dim rngSource As Range, rngDestStart As Range
    Dim cell As Range, targetCell As Range
    Dim insertText As String, locInput As String
    Dim firstCellVal As String, previewVal As String
    Dim foundSample As Boolean
    Dim count As Long
    Dim rowOffset As Long, colOffset As Long
    
    ' 模式开关：True=关键词模式, False=固定位置模式
    Dim isKeywordMode As Boolean
    Dim fixPos As Long
    
    ' 记忆变量
    Dim lastLocInput As String, lastInsertText As String
    Dim insertKeywordPos As VbMsgBoxResult ' 仅关键词模式用
    
    ' =======================================================
    ' 【插件安全锁】
    ' =======================================================
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Application.ActiveWindow.ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If ws.Parent.Name = ThisWorkbook.Name Then
        MsgBox "请先点击一下你要操作的数据表格！", vbExclamation
        Exit Sub
    End If
    ' =======================================================

    On Error Resume Next
    
    ' --- 1. 选区 ---
    Dim defaultAddr As String
    If TypeName(Application.Selection) = "Range" Then defaultAddr = Application.Selection.Address
    
    Set rngSource = Application.InputBox("第1步：请框选【原始数据】区域", "源数据", defaultAddr, Type:=8)
    If rngSource Is Nothing Then Exit Sub
    
    ' --- 2. 输出位置 ---
    Set rngDestStart = Application.InputBox("第2步：请点选【结果存放】起始格", "输出位置", Type:=8)
    If rngDestStart Is Nothing Then Exit Sub
    Set rngDestStart = rngDestStart.Cells(1)
    
    ' 防撞车
    If Not Intersect(rngSource, rngDestStart) Is Nothing Then
        If MsgBox("警告：输出位置覆盖原数据！是否继续？", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If
    On Error GoTo 0
    
    ' =======================================================
    ' --- 3. 核心交互：智能判断模式 ---
    ' =======================================================
    
    ' 读取记忆
    lastLocInput = GetSetting("MyExcelMacros", "BatchInsert", "LocInput", "")
    lastInsertText = GetSetting("MyExcelMacros", "BatchInsert", "InsertText", "")
    
    ' >> 输入插入内容 (提前到这里，逻辑更顺)
    insertText = InputBox("第3步：请输入【要插入的内容】：", "输入内容", lastInsertText)
    If insertText = "" Then Exit Sub
    SaveSetting "MyExcelMacros", "BatchInsert", "InsertText", insertText
    
    ' >> 输入位置规则
    Dim promptMsg As String
    promptMsg = "第4步：请定义插入位置（智能识别）：" & vbCrLf & vbCrLf & _
                "【固定位置模式】请输入数字：" & vbCrLf & _
                "   0  -> 插入在 [开头]" & vbCrLf & _
                "  -1  -> 插入在 [结尾]" & vbCrLf & _
                "   2  -> 插入在 [第2个字符之后] (以此类推)" & vbCrLf & vbCrLf & _
                "【关键词模式】请输入文字：" & vbCrLf & _
                "   例如输入 ""省"" -> 寻找该字进行插入"
                
    locInput = InputBox(promptMsg, "定义规则", lastLocInput)
    If locInput = "" Then Exit Sub
    SaveSetting "MyExcelMacros", "BatchInsert", "LocInput", locInput
    
    ' >> 判断是数字还是文字
    If IsNumeric(locInput) Then
        isKeywordMode = False
        fixPos = CLng(locInput) ' 转换为数字
    Else
        isKeywordMode = True
        ' 关键词模式下，仍需询问前后
        insertKeywordPos = MsgBox("您输入了关键词 [" & locInput & "]" & vbCrLf & _
                                  "请选择插入方向：" & vbCrLf & vbCrLf & _
                                  "【是 (Yes)】 -> 插在关键词 [前面]" & vbCrLf & _
                                  "【否 (No)】  -> 插在关键词 [后面]", vbYesNoCancel + vbQuestion, "关键词方向")
        If insertKeywordPos = vbCancel Then Exit Sub
    End If
    
    ' =======================================================
    ' --- 4. 智能预览 ---
    ' =======================================================
    foundSample = False
    
    For Each cell In rngSource
        If cell.EntireRow.Hidden = False And Not IsError(cell.value) And Len(cell.value) > 0 Then
            
            firstCellVal = cell.value
            Dim tempRes As String
            tempRes = ""
            
            If isKeywordMode Then
                ' >> 关键词模式预览
                If InStr(1, firstCellVal, locInput, vbTextCompare) > 0 Then
                    If insertKeywordPos = vbYes Then
                        tempRes = Replace(firstCellVal, locInput, insertText & locInput, , , vbTextCompare)
                    Else
                        tempRes = Replace(firstCellVal, locInput, locInput & insertText, , , vbTextCompare)
                    End If
                End If
            Else
                ' >> 固定位置模式预览
                tempRes = GetFixedPosResult(firstCellVal, insertText, fixPos)
            End If
            
            ' 如果计算出了结果（且结果不为空），则展示预览
            If tempRes <> "" Then
                previewVal = tempRes
                ' 计算目标位置用于高亮(可选)
                Set targetCell = rngDestStart.Offset(cell.Row - rngSource.Cells(1).Row, cell.Column - rngSource.Cells(1).Column)
                foundSample = True
                Exit For
            End If
        End If
    Next cell
    
    If foundSample Then
        If MsgBox("【效果预览】" & vbCrLf & vbCrLf & _
                  "模式： " & IIf(isKeywordMode, "关键词匹配", "固定位置插入") & vbCrLf & _
                  "原文： " & firstCellVal & vbCrLf & _
                  "结果： " & previewVal & vbCrLf & vbCrLf & _
                  "是否执行？", vbYesNo + vbQuestion, "确认") = vbNo Then Exit Sub
    Else
        If isKeywordMode Then
            MsgBox "未找到包含关键词 """ & locInput & """ 的数据。", vbExclamation
            Exit Sub
        Else
            ' 固定位置模式通常不会找不到，除非全为空
             If MsgBox("预览未找到有效数据（可能是数据为空）。是否强行继续？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        End If
    End If
    
    ' =======================================================
    ' --- 5. 批量执行 ---
    ' =======================================================
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    count = 0
    
    For Each cell In rngSource
        rowOffset = cell.Row - rngSource.Cells(1).Row
        colOffset = cell.Column - rngSource.Cells(1).Column
        Set targetCell = rngDestStart.Offset(rowOffset, colOffset)
        
        If cell.EntireRow.Hidden = False Then
            If Not IsError(cell.value) And Len(cell.value) > 0 Then
                
                Dim finalVal As String
                finalVal = cell.value ' 默认保留原值
                
                If isKeywordMode Then
                    ' 关键词逻辑
                    If InStr(1, cell.value, locInput, vbTextCompare) > 0 Then
                        If insertKeywordPos = vbYes Then
                            finalVal = Replace(cell.value, locInput, insertText & locInput, , , vbTextCompare)
                        Else
                            finalVal = Replace(cell.value, locInput, locInput & insertText, , , vbTextCompare)
                        End If
                        count = count + 1
                    End If
                Else
                    ' 固定位置逻辑
                    finalVal = GetFixedPosResult(cell.value, insertText, fixPos)
                    count = count + 1
                End If
                
                targetCell.value = finalVal
                
            Else
                targetCell.value = ""
            End If
        End If
    Next cell
    
    ' 自动调整列宽
    rngDestStart.Resize(rngSource.Rows.count, rngSource.Columns.count).EntireColumn.AutoFit
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    
    MsgBox "完成！已处理 " & count & " 条数据。", vbInformation

End Sub

' --- 辅助函数：处理固定位置逻辑 ---
Function GetFixedPosResult(strSource As String, strInsert As String, pos As Long) As String
    Dim L As Long
    L = Len(strSource)
    
    If pos = 0 Then
        ' 开头
        GetFixedPosResult = strInsert & strSource
    ElseIf pos = -1 Then
        ' 结尾
        GetFixedPosResult = strSource & strInsert
    Else
        ' 中间某位置 (第N个字之后)
        If pos >= L Then
            ' 如果指定位置超过长度，默认追加到结尾(或者你可以选择不处理)
            GetFixedPosResult = strSource & strInsert
        ElseIf pos > 0 Then
            ' 正常截断插入: 左边pos个字符 + 插入词 + 剩余字符
            GetFixedPosResult = Left(strSource, pos) & strInsert & Mid(strSource, pos + 1)
        Else
            ' 负数情况(除-1外)，暂时按原值返回，也可扩充为倒数第N位
            GetFixedPosResult = strSource
        End If
    End If
End Function

