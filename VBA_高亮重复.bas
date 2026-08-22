'''''' 模块: VBA/高亮重复  类型: 高亮重复.bas  行数: 212
Attribute VB_Name = "高亮重复"
Option Explicit

' =================================================================================
' 宏名称：Super_Duplicate_Highlighter
' 功能：二合一查重工具 (清洗模式/找重模式) + 只向下查找 + 智能跳过空行
' =================================================================================

Sub Super_Duplicate_Highlighter()
    Dim ws As Worksheet
    Dim rngSelect As Range
    Dim colCell As Range
    Dim lastRow As Long, i As Long, startRow As Long
    Dim dict As Object
    Dim strKey As String
    Dim isRowEmpty As Boolean
    Dim colIndices As Collection
    Dim vCol As Variant
    Dim modeChoice As VbMsgBoxResult
    Dim backupAns As VbMsgBoxResult
    
    ' 状态优化变量
    Dim screenUpdateState As Boolean
    Dim calcState As XlCalculation
    Dim eventsState As Boolean
    Dim statusBarState As Boolean

    ' =======================================================
    ' 1. 环境检查与范围选择
    ' =======================================================
    On Error Resume Next
    Set ws = Application.ActiveWindow.ActiveSheet
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    If ws.Parent.Name = ThisWorkbook.Name Then
        MsgBox "请先切换到数据表格！", vbExclamation
        Exit Sub
    End If

    ' 让用户选择列和起始行
    On Error Resume Next
    Set rngSelect = Application.InputBox( _
        prompt:="请选中【起始行】的列标题或单元格" & vbCrLf & _
                "（程序将从这一行开始，向下查找）", _
        Title:="超级查重工具", _
        Type:=8)
    On Error GoTo 0
    If rngSelect Is Nothing Then Exit Sub
    Set ws = rngSelect.Parent
    startRow = rngSelect.Row ' 锁定起始行

    ' =======================================================
    ' 2. 核心：模式选择 (用户决策)
    ' =======================================================
    modeChoice = MsgBox("请选择查重模式：" & vbCrLf & vbCrLf & _
                        "【是 (Yes)】：清洗模式 (保留第1个)" & vbCrLf & _
                        "   -> 第1次出现的保持白色，第2、3次出现的标黄。" & vbCrLf & _
                        "   -> 用途：方便删除多余项。" & vbCrLf & vbCrLf & _
                        "【否 (No)】：找重模式 (高亮所有)" & vbCrLf & _
                        "   -> 只要是重复的，全部标黄 (包括第1个)。" & vbCrLf & _
                        "   -> 用途：查看哪些数据是一组的。", _
                        vbYesNoCancel + vbQuestion, "请选择模式")
    
    If modeChoice = vbCancel Then Exit Sub

    ' =======================================================
    ' 3. 安全备份
    ' =======================================================
    backupAns = MsgBox("是否在执行前创建备份副本？" & vbCrLf & "(强烈建议选是)", vbYesNo + vbExclamation)
    If backupAns = vbYes Then
        Application.ScreenUpdating = False
        ws.Copy After:=ws
        On Error Resume Next
        ActiveSheet.Name = Left(ws.Name, 10) & "_备份" & Format(Now, "hhmm")
        On Error GoTo ErrorHandler
        ws.Activate
    End If

    On Error GoTo ErrorHandler

    ' =======================================================
    ' 4. 初始化与性能优化
    ' =======================================================
    screenUpdateState = Application.ScreenUpdating
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    statusBarState = Application.DisplayStatusBar
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = True

    ' 确定最后一行
    lastRow = ws.Cells(ws.Rows.count, rngSelect.Cells(1).Column).End(xlUp).Row
    If lastRow < startRow Then GoTo CleanUp ' 数据不足

    ' 记录要查的列
    Set colIndices = New Collection
    For Each colCell In rngSelect
        On Error Resume Next
        colIndices.Add colCell.Column, CStr(colCell.Column)
        On Error GoTo 0
    Next colCell

    Set dict = CreateObject("Scripting.Dictionary")

    ' 清除旧颜色 (只清除 startRow 以下的区域)
    For Each vCol In colIndices
        ws.Range(ws.Cells(startRow, vCol), ws.Cells(lastRow, vCol)).Interior.ColorIndex = xlNone
    Next vCol

    ' =======================================================
    ' 5. 分支逻辑处理
    ' =======================================================
    
    If modeChoice = vbYes Then
        ' >>>>>>>>>>>>>>> 模式 A: 清洗模式 (保留首个) <<<<<<<<<<<<<<<
        For i = startRow To lastRow
            If i Mod 500 = 0 Then Application.StatusBar = "清洗模式分析中... " & Format((i - startRow) / (lastRow - startRow), "0%")
            
            strKey = GetRowKey(ws, i, colIndices)
            If strKey = "EMPTY" Then GoTo NextRowA
            
            If dict.Exists(strKey) Then
                ' 字典里有 -> 说明是重复的 -> 标黄
                ColorRow ws, i, colIndices
            Else
                ' 字典里没有 -> 第一次出现 -> 记账
                dict.Add strKey, 1
            End If
NextRowA:
        Next i
        
    Else
        ' >>>>>>>>>>>>>>> 模式 B: 找重模式 (高亮所有) <<<<<<<<<<<<<<<
        ' 需要跑两遍循环
        
        ' 第一遍：只计数
        For i = startRow To lastRow
            If i Mod 1000 = 0 Then Application.StatusBar = "正在统计数据... " & Format((i - startRow) / (lastRow - startRow), "0%")
            
            strKey = GetRowKey(ws, i, colIndices)
            If strKey <> "EMPTY" Then
                If dict.Exists(strKey) Then
                    dict(strKey) = dict(strKey) + 1
                Else
                    dict(strKey) = 1
                End If
            End If
        Next i
        
        ' 第二遍：根据计数上色
        For i = startRow To lastRow
            If i Mod 500 = 0 Then Application.StatusBar = "正在高亮数据... " & Format((i - startRow) / (lastRow - startRow), "0%")
            
            strKey = GetRowKey(ws, i, colIndices)
            If strKey <> "EMPTY" Then
                If dict(strKey) > 1 Then
                    ColorRow ws, i, colIndices
                End If
            End If
        Next i
    End If

    ' =======================================================
    ' 6. 结束
    ' =======================================================
    GoTo CleanUp

ErrorHandler:
    MsgBox "运行出错：" & Err.Description, vbCritical

CleanUp:
    Application.Calculation = calcState
    Application.ScreenUpdating = screenUpdateState
    Application.EnableEvents = eventsState
    Application.StatusBar = False
    Application.DisplayStatusBar = statusBarState
    MsgBox "处理完成！", vbInformation

End Sub

' -----------------------------------------------------------
' 辅助函数：获取行的唯一指纹 (Key)
' -----------------------------------------------------------
Function GetRowKey(ws As Worksheet, r As Long, cols As Collection) As String
    Dim k As String, v As Variant, isEmptyRow As Boolean
    isEmptyRow = True
    k = ""
    For Each v In cols
        k = k & ws.Cells(r, v).value & "|"
        If Len(Trim(ws.Cells(r, v).value)) > 0 Then isEmptyRow = False
    Next v
    If isEmptyRow Then
        GetRowKey = "EMPTY"
    Else
        GetRowKey = k
    End If
End Function

' -----------------------------------------------------------
' 辅助函数：给指定行的指定列上色
' -----------------------------------------------------------
Sub ColorRow(ws As Worksheet, r As Long, cols As Collection)
    Dim v As Variant
    For Each v In cols
        ws.Cells(r, v).Interior.Color = vbYellow
    Next v
End Sub

