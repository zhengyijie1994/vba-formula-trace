'''''' 模块: VBA/错误逆透视  类型: 错误逆透视.bas  行数: 306
Attribute VB_Name = "错误逆透视"
Option Explicit

' ==========================================
' 宏名称：ErrorTransform_Manager (V10.2 - 防漏数据增强版)
' 功能：错误值提取工具 - 支持自定义列排版
' 说明：修复了Transpose限制导致的长文本丢失问题，修复了静默跳过错误的问题
' ==========================================
Public Const APP_REG_KEY As String = "ErrorFix_HighSpeed_V10"

Sub Run_ErrorTransform_Manager()
    ' === 变量声明 ===
    Dim rawInput As String
    Dim finalRule As String
    Dim lastRunRule As String
    Dim historyPrompt As String
    Dim i As Integer, slotVal As String, slotName As String
    Dim cmdPrefix As String, cmdID As String, eqPos As Integer
    
    ' =======================================================
    ' 【核心修复 1】: 启动前检查环境
    ' =======================================================
    Dim tempWS As Worksheet
    On Error Resume Next
    Set tempWS = Application.ActiveWindow.ActiveSheet
    On Error GoTo 0
    
    If tempWS Is Nothing Then Exit Sub
    ' 防止误操作插件本身
    If tempWS.Parent.Name = ThisWorkbook.Name Then
        MsgBox "请先点击一下你需要操作的数据表格（激活它），然后再点击按钮。", vbExclamation
        Exit Sub
    End If
    ' =======================================================
    
    ' =======================================================
    ' 1. 读取配置与构建菜单
    ' =======================================================
    lastRunRule = GetSetting(APP_REG_KEY, "History", "LastRun", "5,6 | 3,4,5,0 | 66,67,68,69,71,72,73")
    
    historyPrompt = "【配置格式说明】 请严格使用竖线 | 分隔三段内容" & vbCrLf & _
                    "--------------------------------------------------------" & vbCrLf & _
                    "格式：表头行,起始行 | 输出列排版(支持0) | 检查列号" & vbCrLf & _
                    "--------------------------------------------------------" & vbCrLf & _
                    "● 第2段 [输出排版] 填写指南：" & vbCrLf & _
                    "   1. 输入数字代表复制源数据列 (例: 3=C列, 5=E列)" & vbCrLf & _
                    "   2. 输入 0 代表插入一列【空白列】" & vbCrLf & _
                    "   3. 系统会自动在最后添加 [属性] 和 [工艺] 两列" & vbCrLf & _
                    "   (示例: 3,4,5,0  ->  C列, D列, E列, 空白列)" & vbCrLf & _
                    "--------------------------------------------------------" & vbCrLf & _
                    "【已保存的常用配置】" & vbCrLf
    
    For i = 1 To 5
        slotVal = GetSetting(APP_REG_KEY, "History", "Slot" & i, "")
        slotName = GetSetting(APP_REG_KEY, "History", "Slot" & i & "_Name", "")
        If slotVal <> "" Then
            historyPrompt = historyPrompt & "[h" & i & "]" & IIf(slotName <> "", " " & slotName, "") & ": " & Left(slotVal, 25) & "..." & vbCrLf
        End If
    Next i
    
    historyPrompt = historyPrompt & "--------------------------------" & vbCrLf & _
                    "指令：h1(运行), s1=...(保存), n1=...(命名), d1(删除)"
    
    ' =======================================================
    ' 2. 接收输入
    ' =======================================================
    rawInput = Application.InputBox(historyPrompt, "自定义排版工具 V10.2 (安全版)", Default:=lastRunRule, Type:=2)
    If rawInput = "False" Or Trim(rawInput) = "" Then Exit Sub
    
    rawInput = Trim(rawInput)
    cmdPrefix = LCase(Left(rawInput, 1))
    eqPos = InStr(rawInput, "=")
    
    ' === 指令处理 ===
    If cmdPrefix = "s" And eqPos > 2 Then ' Save
        cmdID = Mid(rawInput, 2, eqPos - 2)
        If IsNumeric(cmdID) Then
            SaveSetting APP_REG_KEY, "History", "Slot" & cmdID, Mid(rawInput, eqPos + 1)
            MsgBox "√ 配置已保存至 [h" & cmdID & "]", vbInformation
        End If
        Exit Sub
    ElseIf cmdPrefix = "n" And eqPos > 2 Then ' Name
        cmdID = Mid(rawInput, 2, eqPos - 2)
        SaveSetting APP_REG_KEY, "History", "Slot" & cmdID & "_Name", Mid(rawInput, eqPos + 1)
        MsgBox "√ 名称已更新", vbInformation
        Exit Sub
    ElseIf cmdPrefix = "d" And IsNumeric(Mid(rawInput, 2)) Then ' Delete
        cmdID = Mid(rawInput, 2)
        DeleteSetting APP_REG_KEY, "History", "Slot" & cmdID
        DeleteSetting APP_REG_KEY, "History", "Slot" & cmdID & "_Name"
        Exit Sub
    ElseIf cmdPrefix = "h" And IsNumeric(Mid(rawInput, 2)) Then ' History
        cmdID = Mid(rawInput, 2)
        finalRule = GetSetting(APP_REG_KEY, "History", "Slot" & cmdID, "")
        If finalRule = "" Then MsgBox "X 空存档", vbExclamation: Exit Sub
    Else
        finalRule = rawInput ' 直接运行
    End If
    
    ' 记录本次运行
    SaveSetting APP_REG_KEY, "History", "LastRun", finalRule
    
    ' =======================================================
    ' 3. 解析并运行核心逻辑
    ' =======================================================
    Call ParseAndRun_V10(finalRule)
End Sub

' ======================================================
' 解析器 V10
' ======================================================
Sub ParseAndRun_V10(ruleString As String)
    Dim segments() As String
    Dim basicSettings() As String
    Dim pHeader As Long, pStart As Long
    Dim pInfoCols As String, pCheckCols As String
    
    ruleString = Replace(ruleString, "，", ",")
    ruleString = Replace(ruleString, "|", "!")
    segments = Split(ruleString, "!")
    
    ' 1. 基础设置
    If UBound(segments) < 0 Then Exit Sub
    basicSettings = Split(Trim(segments(0)), ",")
    If UBound(basicSettings) < 1 Then MsgBox "格式错误: 缺表头行或起始行", vbCritical: Exit Sub
    pHeader = CLng(basicSettings(0))
    pStart = CLng(basicSettings(1))
    
    ' 2. 信息列 (用户自定义排版)
    pInfoCols = "3,4,5,0" ' 默认值
    If UBound(segments) >= 1 Then pInfoCols = Trim(segments(1))
    If pInfoCols = "" Then pInfoCols = "3,4,5,0"
    
    ' 3. 检查列
    pCheckCols = ""
    If UBound(segments) >= 2 Then pCheckCols = Trim(segments(2))
    If pCheckCols = "" Then MsgBox "错误: 缺检查列", vbCritical: Exit Sub
    
    Call Core_Transform_V10(pHeader, pStart, pInfoCols, pCheckCols)
End Sub

' ======================================================
' 核心逻辑 V10.2 (修复漏数据风险与长文本写入)
' ======================================================
Sub Core_Transform_V10(pHeaderRow As Long, pStartRow As Long, pInfoColsStr As String, pCheckColsStr As String)
    Dim wsSrc As Worksheet, wsTgt As Worksheet
    Dim srcData As Variant, tgtData As Variant, finalData As Variant
    Dim r As Long, tgtRowCounter As Long, lastRow As Long
    Dim checkCols As Variant, infoCols As Variant
    Dim c_col As Variant
    Dim tempArr As Variant, i As Long, k As Long
    Dim propertyName As String, processName As String
    Dim maxColRead As Long: maxColRead = 0
    
    ' --- 1. 数组预处理 ---
    tempArr = Split(pCheckColsStr, ",")
    ReDim checkCols(LBound(tempArr) To UBound(tempArr))
    For i = LBound(tempArr) To UBound(tempArr)
        checkCols(i) = CLng(tempArr(i))
        If checkCols(i) > maxColRead Then maxColRead = checkCols(i)
    Next i
    
    tempArr = Split(pInfoColsStr, ",")
    ReDim infoCols(LBound(tempArr) To UBound(tempArr))
    For i = LBound(tempArr) To UBound(tempArr)
        infoCols(i) = CLng(tempArr(i))
        If infoCols(i) > 0 Then
            If infoCols(i) > maxColRead Then maxColRead = infoCols(i)
        End If
    Next i
    
    Dim numInfoCols As Long
    numInfoCols = UBound(infoCols) - LBound(infoCols) + 1
    
    ' --- 2. 准备数据 ---
    Set wsSrc = Application.ActiveWindow.ActiveSheet
    If wsSrc.Name = "错误报告" Then MsgBox "请激活源数据表，不要在错误报告表里操作", vbCritical: Exit Sub
    
    Dim lastCell As Range
    Set lastCell = wsSrc.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then Exit Sub
    lastRow = lastCell.Row
    If lastRow < pStartRow Then Exit Sub
    
    ' 【安全修正】防止用户填写的列号超过Excel实际列数导致报错
    If maxColRead > wsSrc.Columns.count Then maxColRead = wsSrc.Columns.count
    ' 【安全修正】防止读取范围小于实际数据范围导致下标越界
    If maxColRead < wsSrc.UsedRange.Columns.count Then maxColRead = wsSrc.UsedRange.Columns.count
    
    srcData = wsSrc.Range("A1", wsSrc.Cells(lastRow, maxColRead)).Value2
    
    Dim totalTgtCols As Long
    totalTgtCols = numInfoCols + 2
    
    Dim maxOut As Long: maxOut = 10000
    ' 保持 (列, 行) 结构以便 Preserve 扩容，写入时再手动转置
    ReDim tgtData(1 To totalTgtCols, 1 To maxOut)
    tgtRowCounter = 0
    
    ' --- 3. 循环处理 ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 【核心修复】移除 On Error Resume Next，使用逻辑判断避免漏数据
    On Error GoTo 0
    
    For r = pStartRow To lastRow
        For Each c_col In checkCols
            ' 边界检查：确保检查列在源数据范围内
            If c_col <= UBound(srcData, 2) Then
                If IsError(srcData(r, c_col)) Then
                    tgtRowCounter = tgtRowCounter + 1
                    
                    ' 动态扩容
                    If tgtRowCounter > maxOut Then
                        maxOut = maxOut * 2
                        ReDim Preserve tgtData(1 To totalTgtCols, 1 To maxOut)
                    End If
                    
                    ' A. 填充用户定义列
                    For k = 0 To numInfoCols - 1
                        Dim targetIdx As Long
                        targetIdx = infoCols(k)
                        
                        If targetIdx = 0 Then
                            tgtData(k + 1, tgtRowCounter) = ""
                        Else
                            ' 边界检查：确保信息列在源数据范围内
                            If targetIdx <= UBound(srcData, 2) Then
                                tgtData(k + 1, tgtRowCounter) = srcData(r, targetIdx)
                            Else
                                tgtData(k + 1, tgtRowCounter) = "列" & targetIdx & "越界"
                            End If
                        End If
                    Next k
                    
                    ' B. 处理属性名
                    propertyName = srcData(pHeaderRow, c_col)
                    Select Case propertyName
                        Case "低温成定": processName = "成定"
                        Case "低温预定": processName = "预定"
                        Case "高温中定", "烘干加起毛剂", "定型加起毛剂": processName = "烘干"
                        Case Else: processName = propertyName
                    End Select
                    
                    ' C. 填充固定列
                    tgtData(numInfoCols + 1, tgtRowCounter) = propertyName
                    tgtData(numInfoCols + 2, tgtRowCounter) = processName
                End If
            End If
        Next c_col
    Next r
    
    ' --- 4. 输出 ---
    Dim targetWB As Workbook
    Set targetWB = wsSrc.Parent
    
    Application.DisplayAlerts = False
    On Error Resume Next
    targetWB.Sheets("错误报告").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsTgt = targetWB.Sheets.Add
    wsTgt.Name = "错误报告"
    
    ' 写入表头
    For i = 0 To numInfoCols - 1
        If infoCols(i) = 0 Then
            wsTgt.Cells(1, i + 1).value = ""
        Else
            ' 边界保护
            If infoCols(i) <= UBound(srcData, 2) Then
                wsTgt.Cells(1, i + 1).value = srcData(pHeaderRow, infoCols(i))
            End If
        End If
    Next i
    
    wsTgt.Cells(1, numInfoCols + 1).value = "属性"
    wsTgt.Cells(1, numInfoCols + 2).value = "工艺"
    
    ' 【核心修复】使用手动双循环写入，替代 Application.Transpose
    If tgtRowCounter > 0 Then
        ' 准备最终输出数组 (行, 列)
        ReDim finalData(1 To tgtRowCounter, 1 To totalTgtCols)
        
        Dim rOut As Long, cOut As Long
        ' 手动转置：把 tgtData(列, 行) 变成 finalData(行, 列)
        For rOut = 1 To tgtRowCounter
            For cOut = 1 To totalTgtCols
                finalData(rOut, cOut) = tgtData(cOut, rOut)
            Next cOut
        Next rOut
        
        ' 一次性安全写入
        wsTgt.Range("A2").Resize(tgtRowCounter, totalTgtCols).Value2 = finalData
        wsTgt.Columns.AutoFit
        MsgBox "完成！已提取 " & tgtRowCounter & " 条错误。" & vbCrLf & "排版已按配置生成。", vbInformation
    Else
        MsgBox "未发现错误。", vbInformation
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

