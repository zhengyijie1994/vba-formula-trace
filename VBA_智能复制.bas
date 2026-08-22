'''''' 模块: VBA/智能复制  类型: 智能复制.bas  行数: 362
Attribute VB_Name = "智能复制"
Option Explicit

' ==============================================================================
' 宏名称：SmartCopy_History_Manager (智能列提取工具)
' 版本号：v18.5 (修复版 - 解决 1004 可见单元格报错问题)
' ==============================================================================
Public Const APP_VERSION As String = "v18.5"

Sub SmartCopy_History_Manager()
    ' ==========================================
    ' 1. 变量定义区域
    ' ==========================================
    Dim rawInput As String, finalRule As String, lastRunRule As String
    Dim rngPick As Range, rngCheck As Range
    Dim startRow As Long, endRow As Long, pickCol As Long
    
    Dim historyPrompt As String
    Dim i As Integer, j As Integer
    Dim slotVal As String, slotName As String
    Dim cmdPrefix As String, cmdID As String, eqPos As Integer
    
    Dim processList As Collection
    Dim currentItem As String, nextItem As String
    Dim colIndex As Long
    
    Dim sourceSheet As Worksheet, newSheet As Worksheet
    Dim destColIndex As Integer, blockStartCol As Integer, sumCount As Long
    Dim targetName As String
    Dim finalDestRowCount As Long
    
    Dim rngSourceBlock As Range
    Dim scrollPos As Long
    Dim useFastMode As Boolean
    Dim batchCount As Integer
    
    ' --- 变量 (用于鼠标选列功能) ---
    Dim rngColumnPick As Range
    Dim areaPart As Range
    Dim c_idx As Long
    Dim colStr As String
    
    ' =======================================================
    ' 【环境检查】: 强制锁定当前活动窗口的表格
    ' =======================================================
    On Error Resume Next
    Set sourceSheet = Application.ActiveWindow.ActiveSheet
    On Error GoTo 0
    
    ' 双重检查：如果没有抓到表，或者抓到了插件本身，强制退出
    If sourceSheet Is Nothing Then Exit Sub
    If sourceSheet.Parent.Name = ThisWorkbook.Name Then
        MsgBox "请先点击一下你需要操作的数据表格（激活它），然后再点击按钮。", vbExclamation
        Exit Sub
    End If
    
    ' =======================================================
    ' 步骤 1: 规则输入 & 管理
    ' =======================================================
    ' 读取上次运行记录
    lastRunRule = ReadSmartProp("LastRun")
    If lastRunRule = "" Then lastRunRule = "1, 2, 3"
    
    ' 构建历史记录提示文本
    historyPrompt = "【存储模式】： 文件内嵌 (随文件移动)" & vbCrLf & _
                    "【上次运行】： " & lastRunRule & vbCrLf & _
                    "--------------------------------" & vbCrLf & "【本表已存规则】" & vbCrLf
    
    For i = 1 To 8
        slotVal = ReadSmartProp("Slot" & i)
        slotName = ReadSmartProp("Slot" & i & "_Name")
        If slotVal <> "" Then
            historyPrompt = historyPrompt & "[h" & i & "]" & IIf(slotName <> "", " " & slotName, "") & ": " & slotVal & vbCrLf
        End If
    Next i
    
    historyPrompt = historyPrompt & "--------------------------------" & vbCrLf & _
                    "指令：输入 m 鼠标选列 / h1 调取 / s1=... 保存"
    
    ' 弹出输入框
    rawInput = Application.InputBox(historyPrompt, "步骤 1/3: 规则输入 (" & APP_VERSION & ")", Default:=lastRunRule, Type:=2)
    If rawInput = "False" Or Trim(rawInput) = "" Then Exit Sub
    
    rawInput = Trim(rawInput)
    cmdPrefix = LCase(Left(rawInput, 1))
    eqPos = InStr(rawInput, "=")
    
    ' --- 管理指令处理 (Save/Name/Delete/History/Mouse) ---
    If cmdPrefix = "s" And eqPos > 2 Then
        cmdID = Mid(rawInput, 2, eqPos - 2)
        SaveSmartProp "Slot" & cmdID, Mid(rawInput, eqPos + 1)
        MsgBox "√ 规则已写入文件内部 [h" & cmdID & "]", vbInformation
        Exit Sub
    ElseIf cmdPrefix = "n" And eqPos > 2 Then
        cmdID = Mid(rawInput, 2, eqPos - 2)
        SaveSmartProp "Slot" & cmdID & "_Name", Mid(rawInput, eqPos + 1)
        MsgBox "√ 备注已保存", vbInformation
        Exit Sub
    ElseIf cmdPrefix = "d" And IsNumeric(Mid(rawInput, 2)) Then
        cmdID = Mid(rawInput, 2)
        DeleteSmartProp "Slot" & cmdID
        DeleteSmartProp "Slot" & cmdID & "_Name"
        MsgBox "√ 已清除", vbInformation
        Exit Sub
    ElseIf cmdPrefix = "h" And IsNumeric(Mid(rawInput, 2)) Then
        finalRule = ReadSmartProp("Slot" & Mid(rawInput, 2))
        If finalRule = "" Then MsgBox "无此存档", vbExclamation: Exit Sub
        
    ' --- 鼠标拾取模式 (m) ---
    ElseIf cmdPrefix = "m" And Len(rawInput) = 1 Then
        On Error Resume Next
        sourceSheet.Activate
        Set rngColumnPick = Application.InputBox("【鼠标拾取模式】" & vbCrLf & vbCrLf & _
                                                 "请用鼠标点击列头 (例如点击 D列)。" & vbCrLf & _
                                                 "1. 按住 Ctrl 可不连续多选 (顺序：点D, 点B -> 对应 4,2)" & vbCrLf & _
                                                 "2. 拖拽可选择连续区域 (A:C -> 对应 1,2,3)", _
                                                 "步骤 1.5: 鼠标拾取列", Type:=8)
        On Error GoTo 0
        
        If rngColumnPick Is Nothing Then Exit Sub
        
        colStr = ""
        For Each areaPart In rngColumnPick.Areas
            For c_idx = 1 To areaPart.Columns.count
                colStr = colStr & areaPart.Columns(c_idx).Column & ", "
            Next c_idx
        Next areaPart
        
        If Len(colStr) > 2 Then
            finalRule = Left(colStr, Len(colStr) - 2)
        Else
            Exit Sub
        End If
        
    Else
        finalRule = rawInput
    End If
    
    ' 保存本次运行记录并解析
    SaveSmartProp "LastRun", finalRule
    Set processList = ParseColumnRule_v13(finalRule)
    
    ' =======================================================
    ' 步骤 2: 鼠标定位 (行范围选择)
    ' =======================================================
    On Error Resume Next
    sourceSheet.Activate
    Set rngPick = Application.InputBox("规则已确认：" & finalRule & vbCrLf & vbCrLf & _
                                     "请用鼠标定位数据源：" & vbCrLf & _
                                     "【模式 A】点选单格 = 自动计算至最后一行" & vbCrLf & _
                                     "【模式 B】框选范围 = 只复制选中的行数", _
                                     "步骤 2/3: 智能定位数据", Type:=8)
    On Error GoTo 0
    If rngPick Is Nothing Then Exit Sub
    
    If rngPick.Parent.Name <> sourceSheet.Name Then
        Set sourceSheet = rngPick.Worksheet
    End If
    
    startRow = rngPick.Row
    pickCol = rngPick.Column
    
    If rngPick.Rows.count > 1 Then
        endRow = startRow + rngPick.Rows.count - 1
    Else
        endRow = sourceSheet.Cells(sourceSheet.Rows.count, pickCol).End(xlUp).Row
        If endRow < startRow Then endRow = startRow
    End If

    sourceSheet.Activate
    sourceSheet.Cells(endRow, pickCol).Select
    scrollPos = endRow - 15: If scrollPos < 1 Then scrollPos = 1
    ActiveWindow.ScrollRow = scrollPos
    If MsgBox("【预检查】" & vbCrLf & "提取规则: " & finalRule & vbCrLf & "数据末行: " & endRow & vbCrLf & vbCrLf & "点击【是】开始处理。", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    ' =======================================================
    ' 步骤 3: 执行引擎
    ' =======================================================
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .PrintCommunication = False
    End With
    
    targetName = Left(sourceSheet.Name, 20) & "_处理结果"
    
    Dim targetWB As Workbook
    Set targetWB = sourceSheet.Parent
    
    On Error Resume Next
    Set newSheet = targetWB.Sheets(targetName)
    If Not newSheet Is Nothing Then
        newSheet.Cells.Clear
        newSheet.Move After:=targetWB.Sheets(targetWB.Sheets.count)
    Else
        Set newSheet = targetWB.Sheets.Add(After:=targetWB.Sheets(targetWB.Sheets.count))
        newSheet.Name = targetName
    End If
    On Error GoTo 0
    
    ' === 智能检测隐藏行 ===
    Set rngCheck = sourceSheet.Range(sourceSheet.Cells(startRow, pickCol), sourceSheet.Cells(endRow, pickCol))
    If rngCheck.SpecialCells(xlCellTypeVisible).count < rngCheck.count Then
        useFastMode = False ' 发现隐藏行 -> 安全模式
    Else
        useFastMode = True  ' 无任何隐藏 -> 极速模式
    End If
    
    destColIndex = 1
    blockStartCol = 0
    finalDestRowCount = 0
    
    i = 1
    Do While i <= processList.count
        currentItem = CStr(processList(i))
        
        ' --- 处理小计和空列 ---
        If Left(currentItem, 1) = "+" Or currentItem = "0" Then
            If Left(currentItem, 1) = "+" Then
                If Len(currentItem) > 1 And IsNumeric(Mid(currentItem, 2)) Then sumCount = CLng(Mid(currentItem, 2)) Else sumCount = IIf(blockStartCol > 0, destColIndex - blockStartCol, 0)
                If sumCount > 0 And destColIndex > sumCount And finalDestRowCount > 0 Then
                    newSheet.Cells(1, destColIndex).value = "小计"
                    newSheet.Range(newSheet.Cells(1, destColIndex), newSheet.Cells(finalDestRowCount, destColIndex)).FormulaR1C1 = "=SUM(RC[" & (-1 * sumCount) & "]:RC[-1])"
                    blockStartCol = 0: destColIndex = destColIndex + 1
                End If
            Else: blockStartCol = 0: destColIndex = destColIndex + 1: End If
            i = i + 1
            
        ' --- 处理数据列 ---
        Else
            colIndex = CLng(currentItem)
            
            ' 连续列打包检测
            batchCount = 1
            If i < processList.count Then
                For j = i + 1 To processList.count
                    nextItem = CStr(processList(j))
                    If IsNumeric(nextItem) And Not (Left(nextItem, 1) = "+" Or nextItem = "0") Then
                        If CLng(nextItem) = colIndex + batchCount Then batchCount = batchCount + 1 Else Exit For
                    Else: Exit For: End If
                Next j
            End If
            
            If colIndex > 0 Then
                If blockStartCol = 0 Then blockStartCol = destColIndex
                
                Set rngSourceBlock = sourceSheet.Range(sourceSheet.Cells(startRow, colIndex), sourceSheet.Cells(endRow, colIndex + batchCount - 1))
                
                ' =======================================================
                ' 【核心修复区域 v18.5】: 防御 1004 错误
                ' =======================================================
                If useFastMode Then
                    ' 极速模式：整块复制 (适用于无隐藏行)
                    rngSourceBlock.Copy
                    newSheet.Cells(1, destColIndex).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    If finalDestRowCount = 0 Then finalDestRowCount = endRow - startRow + 1
                Else
                    ' 安全模式：带错误防御的可见单元格复制
                    On Error Resume Next ' <--- 开启防爆盾
                    Dim rngVisible As Range
                    Set rngVisible = Nothing
                    
                    ' 尝试获取可见区域
                    Set rngVisible = rngSourceBlock.SpecialCells(xlCellTypeVisible)
                    
                    ' 只有当确实找到了可见单元格时，才执行复制
                    If Err.Number = 0 And Not rngVisible Is Nothing Then
                        rngVisible.Copy
                        newSheet.Cells(1, destColIndex).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                        
                        ' 更新行数 (仅第一次有效复制时计算)
                        If finalDestRowCount = 0 Then
                            finalDestRowCount = newSheet.Cells(newSheet.Rows.count, destColIndex).End(xlUp).Row
                            If finalDestRowCount < 1 Then finalDestRowCount = 1
                        End If
                    End If
                    ' 如果这列全是隐藏行，代码会自动跳过不报错
                    On Error GoTo 0 ' <--- 关闭防爆盾
                End If
                ' =======================================================
                
                destColIndex = destColIndex + batchCount
                i = i + batchCount
            Else
                i = i + 1
            End If
        End If
    Loop
    
    newSheet.Columns.AutoFit
    Application.CutCopyMode = False
    
    With Application
        .PrintCommunication = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    
    newSheet.Activate
    If finalDestRowCount > 0 And destColIndex > 1 Then
        newSheet.Range(newSheet.Cells(1, 1), newSheet.Cells(finalDestRowCount, destColIndex - 1)).Select
    End If
    
End Sub

' ======================================================
' 存储核心函数 (保持原样)
' ======================================================
Function ReadSmartProp(propName As String) As String
    On Error Resume Next
    ReadSmartProp = ActiveWorkbook.CustomDocumentProperties("SmartCopy_" & propName).value
    On Error GoTo 0
End Function

Sub SaveSmartProp(propName As String, propValue As String)
    Dim fullKey As String: fullKey = "SmartCopy_" & propName
    On Error Resume Next: ActiveWorkbook.CustomDocumentProperties(fullKey).Delete: On Error GoTo 0
    On Error Resume Next: ActiveWorkbook.CustomDocumentProperties.Add Name:=fullKey, LinkToContent:=False, Type:=4, value:=propValue: On Error GoTo 0
End Sub

Sub DeleteSmartProp(propName As String)
    On Error Resume Next: ActiveWorkbook.CustomDocumentProperties("SmartCopy_" & propName).Delete: On Error GoTo 0
End Sub

' ======================================================
' 规则解析器 (保持原样)
' ======================================================
Function ParseColumnRule_v13(strInput As String) As Collection
    Dim result As New Collection
    Dim parts() As String, subPart As Variant, rangeParts() As String
    Dim startNum As Long, endNum As Long, k As Long, tempStr As String
    strInput = Replace(strInput, "，", ","): strInput = Replace(strInput, " ", ""): parts = Split(strInput, ",")
    For Each subPart In parts
        tempStr = CStr(subPart)
        If Left(tempStr, 1) = "+" Then
            result.Add tempStr
        ElseIf InStr(tempStr, "-") > 0 Then
            rangeParts = Split(tempStr, "-")
            If UBound(rangeParts) = 1 Then
                If IsNumeric(rangeParts(0)) And IsNumeric(rangeParts(1)) Then
                    startNum = CLng(rangeParts(0)): endNum = CLng(rangeParts(1))
                    If startNum <= endNum Then
                        For k = startNum To endNum: result.Add k: Next k
                    End If
                End If
            End If
        Else
            If IsNumeric(tempStr) Then
                result.Add CLng(tempStr)
            Else
                On Error Resume Next
                Dim colNum As Long: colNum = Range(tempStr & "1").Column
                If Err.Number = 0 Then result.Add colNum
                On Error GoTo 0
            End If
        End If
    Next subPart
    Set ParseColumnRule_v13 = result
End Function

