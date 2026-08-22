'''''' 模块: VBA/插入精炼行  类型: 插入精炼行.bas  行数: 225
Attribute VB_Name = "插入精炼行"
Option Explicit

' ==========================================
' 宏名称：InsertGroupHeaders_SafeMode_Plugin
' 功能：分组插入标题行 (H-T列留空) - 【悬浮进度+精准复制版】
' ==========================================

Sub InsertGroupHeaders_SafeMode_Plugin()
    Dim ws As Worksheet
    Dim lastRow As Long, actualLastCol As Long
    Dim colInput As String, colParts() As String
    Dim checkCols() As Long
    Dim k As Long, part As String, v As Long
    Dim isDifferent As Boolean
    Dim defaultVal As String, confirmAns As VbMsgBoxResult
    
    ' 状态变量
    Dim screenUpdateState As Boolean, calcState As XlCalculation
    Dim eventsState As Boolean, origView As XlWindowView
    
    ' 优化与内存变量
    Dim arrCols() As Variant
    Dim visRng As Range, area As Range
    Dim visRowNums() As Long, visCount As Long
    Dim currRow As Long, prevRow As Long
    Dim targetRows() As Long, targetCount As Long
    Dim firstRow As Long, rowsCount As Long, rIdx As Long
    
    ' 进度条专用变量
    Dim progBox As Shape
    Dim winTop As Double, winLeft As Double

    ' =======================================================
    ' 1. 安全锁与备份
    ' =======================================================
    On Error Resume Next
    Set ws = Application.ActiveWindow.ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If ws.Parent.Name = ThisWorkbook.Name Then
        MsgBox "请先点击一下你要操作的数据表格！", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    
    confirmAns = MsgBox("?? 是否在执行前，自动创建一个当前工作表的【备份副本】？", vbYesNoCancel + vbExclamation, "安全保护")
    If confirmAns = vbCancel Then Exit Sub
    If confirmAns = vbYes Then
        Application.ScreenUpdating = False
        ws.Copy After:=ws
        On Error Resume Next
        ActiveSheet.Name = Left(ws.Name, 20) & "_备份" & Format(Now, "hhmm")
        On Error GoTo ErrorHandler
        ws.Activate
        Application.ScreenUpdating = True
    End If

    ' --- 获取依据列 ---
    defaultVal = GetSetting("MyExcelMacros", "GroupHeader", "LastCols", "1")
    colInput = InputBox("请输入判断依据列号：" & vbCrLf & "例如：'1,2' (对应A、B列)", "分组插入工具", defaultVal)
    If colInput = "" Then Exit Sub
    colInput = Replace(colInput, "，", ",")
    colParts = Split(colInput, ",")
    ReDim checkCols(0 To UBound(colParts))
    On Error Resume Next
    For k = 0 To UBound(colParts)
        part = Trim(colParts(k))
        If IsNumeric(part) Then checkCols(k) = CLng(part) Else checkCols(k) = ws.Range(part & "1").Column
    Next k
    On Error GoTo ErrorHandler
    SaveSetting "MyExcelMacros", "GroupHeader", "LastCols", colInput
    
    lastRow = ws.Cells(ws.Rows.count, checkCols(0)).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' ?? 性能点：获取真实使用的最大列宽，拒绝复制整行16384列
    actualLastCol = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If actualLastCol < 20 Then actualLastCol = 20 ' 确保至少覆盖到T列

    ' ==========================================
    ' 2. 环境冻结与UI接管
    ' ==========================================
    origView = ActiveWindow.View
    ActiveWindow.View = xlNormalView
    
    screenUpdateState = Application.ScreenUpdating
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ws.DisplayPageBreaks = False
    
    ' ?? 性能点：切断条件格式在插入时的灾难级重算 (支持新版Excel)
    On Error Resume Next
    ws.EnableFormatConditionsCalculation = False
    On Error GoTo ErrorHandler

    ' ?? 创建显式的【悬浮进度面板】，解决底部状态栏假死看不见的问题
    winTop = ActiveWindow.VisibleRange.Top + ActiveWindow.VisibleRange.Height / 2 - 40
    winLeft = ActiveWindow.VisibleRange.Left + ActiveWindow.VisibleRange.Width / 2 - 150
    Set progBox = ws.Shapes.AddShape(msoShapeRectangle, winLeft, winTop, 300, 80)
    With progBox
        .Name = "MyVBA_ProgressBox"
        .Fill.ForeColor.RGB = RGB(40, 40, 40)
        .Line.Visible = msoFalse
        With .TextFrame.Characters.Font
            .Color = RGB(255, 255, 255)
            .Size = 14
            .Bold = True
        End With
        .TextFrame.Characters.text = "?? 正在进行内存读取与排查..." & vbCrLf & "请勿点击鼠标..."
    End With
    
    ' 强制刷新一次屏幕以显示黑框
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False

    ' ==========================================
    ' 3. 第一阶段：纯内存闪电排查
    ' ==========================================
    ReDim arrCols(0 To UBound(checkCols))
    For k = 0 To UBound(checkCols)
        arrCols(k) = ws.Range(ws.Cells(1, checkCols(k)), ws.Cells(lastRow, checkCols(k))).Value2
    Next k

    On Error Resume Next
    Set visRng = ws.Range("A1:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo ErrorHandler
    If visRng Is Nothing Then GoTo CleanUp
    
    ReDim visRowNums(1 To lastRow)
    visCount = 0
    For Each area In visRng.Areas
        firstRow = area.Row
        rowsCount = area.Rows.count
        For rIdx = 0 To rowsCount - 1
            visCount = visCount + 1
            visRowNums(visCount) = firstRow + rIdx
        Next rIdx
    Next area

    ReDim targetRows(1 To visCount)
    targetCount = 0
    
    For v = visCount To 2 Step -1
        currRow = visRowNums(v)
        prevRow = visRowNums(v - 1)
        isDifferent = False
        For k = 0 To UBound(checkCols)
            If arrCols(k)(currRow, 1) <> arrCols(k)(prevRow, 1) Then
                isDifferent = True
                Exit For
            End If
        Next k
        
        If isDifferent Then
            targetCount = targetCount + 1
            targetRows(targetCount) = currRow
        End If
    Next v
    
    ' ==========================================
    ' 4. 第二阶段：物理定点写入 + 强刷UI面板
    ' ==========================================
    If targetCount > 0 Then
        For k = 1 To targetCount
            currRow = targetRows(k)
            
            ' ?? 核心：每隔 10 行强制更新悬浮黑框进度！(你绝对能看到进度了)
            If k Mod 10 = 0 Or k = 1 Then
                Application.ScreenUpdating = True ' 瞬间打开屏幕刷新
                progBox.TextFrame.Characters.text = "?? 正在写入第 " & k & " 个分组..." & vbCrLf & _
                                                    "总进度: " & Format(k / targetCount, "0.0%") & " (" & k & " / " & targetCount & ")"
                DoEvents ' 让 Windows 重绘这个黑框
                Application.ScreenUpdating = False ' 迅速关掉屏幕刷新继续极速跑
            End If
            
            ' 插入操作
            ws.Rows(currRow).Insert Shift:=xlDown
            
            ' ?? 终极性能点：精准复制！只复制到实际存在的最后一列，彻底杜绝16384列的灾难复制
            ws.Range(ws.Cells(currRow + 1, 1), ws.Cells(currRow + 1, actualLastCol)).Copy _
                destination:=ws.Range(ws.Cells(currRow, 1), ws.Cells(currRow, actualLastCol))
            
            ' 清理 H(8) 到 T(20) 列
            ws.Cells(currRow, 8).Resize(1, 13).ClearContents
        Next k
    End If
    
CleanUp:
    On Error Resume Next
    progBox.Delete ' 删除悬浮进度面板
    ws.EnableFormatConditionsCalculation = True ' 恢复条件格式
    On Error GoTo 0
    
    ActiveWindow.View = origView
    Application.CutCopyMode = False
    Application.Calculation = calcState
    Application.ScreenUpdating = screenUpdateState
    Application.EnableEvents = eventsState
    
    If targetCount > 0 Then
        MsgBox "? 分组插入极速完成！共处理了 " & targetCount & " 个分组。", vbInformation
    Else
        MsgBox "没有发现需要分组的数据。", vbInformation
    End If
    Exit Sub

ErrorHandler:
    On Error Resume Next
    progBox.Delete
    ws.EnableFormatConditionsCalculation = True
    ActiveWindow.View = xlNormalView
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "运行出错：" & Err.Description, vbCritical
End Sub

