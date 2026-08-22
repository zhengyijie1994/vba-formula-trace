'''''' 模块: VBA/自定义排序  类型: 自定义排序.bas  行数: 413
Attribute VB_Name = "自定义排序"
Option Explicit

' ==================================================================================
'  智能多列排序系统 (V4.1 预览增强版)
'  新增功能：支持输入 p + 序号 (如 p1) 来预览方案详情，防止忘记排序逻辑
'  核心架构：V4.0 金库版 (数据永不丢失)
' ==================================================================================

' 隐藏的金库表名称 (请勿修改)
Const SAFE_SHEET_NAME As String = "_Sys_Config_Safe_"
Const KEY_LAST_USED As String = "LastUsed"

Sub MultiColumnSmartSort_Manager()
    Dim ws As Worksheet
    Dim profileListStr As String
    Dim profiles() As String
    Dim menuMsg As String
    Dim userChoice As String
    Dim i As Integer
    Dim hasLast As Boolean
    
    ' 1. 环境检查
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ' 2. 读取列表
    profileListStr = GetProfileList()
    
    ' 3. 构建菜单
    menuMsg = "【智能排序控制台 - V4.1】" & vbCrLf & "--------------------------------" & vbCrLf & _
              "1. ? 新建排序 (New)" & vbCrLf
    
    ' 检查上一次记录
    If LoadDataFromSafe(KEY_LAST_USED) <> "" Then
        menuMsg = menuMsg & "2. ?? 重复上一次操作" & vbCrLf
        hasLast = True
    Else
        menuMsg = menuMsg & "2. (无最近记录)" & vbCrLf
        hasLast = False
    End If
    
    menuMsg = menuMsg & "--------------------------------" & vbCrLf & "已存方案：" & vbCrLf
    
    If profileListStr <> "" Then
        profiles = Split(profileListStr, "|")
        For i = LBound(profiles) To UBound(profiles)
            menuMsg = menuMsg & (i + 3) & ". ?? " & profiles(i) & vbCrLf
        Next i
    Else
        menuMsg = menuMsg & "(暂无保存的方案)" & vbCrLf
    End If
    
    menuMsg = menuMsg & vbCrLf & "--------------------------------" & vbCrLf & _
              "指令说明：" & vbCrLf & _
              "? 输入数字: 直接执行" & vbCrLf & _
              "? 输入 p+数字: 预览详情 (如 p3)" & vbCrLf & _
              "? 输入 d+数字: 删除方案 (如 d3)"
    
    ' 4. 获取输入
    userChoice = Application.InputBox(menuMsg, "选择模式", Type:=2)
    If userChoice = "False" Or userChoice = "" Then Exit Sub
    userChoice = LCase(Trim(userChoice)) ' 转小写并去空格
    
    ' --- 删除指令 (d) ---
    If Left(userChoice, 1) = "d" Then
        Call DeleteProfileHandler(Mid(userChoice, 2), profiles)
        Exit Sub
    End If
    
    ' --- 预览指令 (p) ---
    If Left(userChoice, 1) = "p" Then
        Call PreviewProfileHandler(ws, Mid(userChoice, 2), profiles, hasLast)
        Exit Sub
    End If
    
    ' --- 执行指令 ---
    Select Case userChoice
        Case "1"
            Call StartNewSortProcess(ws)
        Case "2"
            If hasLast Then
                Call ExecuteProfile(ws, KEY_LAST_USED, False)
            Else
                MsgBox "无记录。", vbExclamation
            End If
        Case Else
            Dim idx As Integer
            On Error Resume Next
            idx = CInt(userChoice) - 3
            If Err.Number <> 0 Or idx < LBound(profiles) Or idx > UBound(profiles) Then
                MsgBox "无效输入。", vbCritical
                Exit Sub
            End If
            On Error GoTo 0
            Call ExecuteProfile(ws, profiles(idx), True)
    End Select
End Sub

' ========================================================
'  新增：预览逻辑
' ========================================================
Sub PreviewProfileHandler(ws As Worksheet, idxStr As String, profiles() As String, hasLast As Boolean)
    Dim configStr As String
    Dim profileName As String
    Dim targetIdx As Integer
    
    On Error Resume Next
    targetIdx = CInt(idxStr)
    On Error GoTo 0
    
    ' 判断查看对象
    If targetIdx = 2 Then
        If Not hasLast Then MsgBox "无最近记录。", vbExclamation: Exit Sub
        profileName = KEY_LAST_USED
        configStr = LoadDataFromSafe(KEY_LAST_USED)
    Else
        targetIdx = targetIdx - 3
        If targetIdx < LBound(profiles) Or targetIdx > UBound(profiles) Then
            MsgBox "无效序号。", vbCritical: Exit Sub
        End If
        profileName = profiles(targetIdx)
        configStr = LoadDataFromSafe(profileName)
    End If
    
    ' 解析并生成可读报告
    Dim rules() As String, parts() As String
    Dim i As Integer
    Dim msg As String
    Dim rng As Range, headerVal As String, cellAddr As String
    Dim orderStr As String
    
    rules = Split(configStr, "|")
    msg = "【方案预览：" & IIf(profileName = KEY_LAST_USED, "上一次操作", profileName) & "】" & vbCrLf & _
          "-------------------------------------" & vbCrLf
    
    For i = LBound(rules) To UBound(rules)
        parts = Split(rules(i), "#")
        cellAddr = parts(0)
        
        ' 尝试获取当前表中对应列的标题
        headerVal = "(空标题)"
        On Error Resume Next
        Set rng = ws.Range(cellAddr)
        If Not rng Is Nothing Then headerVal = rng.value
        On Error GoTo 0
        
        ' 解析方向
        If parts(1) = "1" Then orderStr = "升序 (A-Z)" Else orderStr = "降序 (Z-A)"
        
        ' 拼接信息
        msg = msg & "第 " & (i + 1) & " 级：[" & headerVal & "] " & vbCrLf & _
              "       位置: " & cellAddr & "  ->  " & orderStr & vbCrLf
    Next i
    
    msg = msg & "-------------------------------------" & vbCrLf & "点击【确定】返回，需要执行请重新运行输入数字。"
    MsgBox msg, vbInformation, "方案详情"
End Sub

' ========================================================
'  核心流程 A：新建排序
' ========================================================
Sub StartNewSortProcess(ws As Worksheet)
    Dim fullTable As Range, sortKey As Range, inputRange As Range
    Dim isFirst As Boolean, configStr As String
    
    isFirst = True
    ws.Sort.SortFields.Clear
    configStr = ""
    
    Do
        Dim promptMsg As String
        If isFirst Then
            promptMsg = "【第 1 排序级别】" & vbCrLf & "点击列的任意单元格。"
        Else
            promptMsg = "【第 " & (ws.Sort.SortFields.count + 1) & " 排序级别】" & vbCrLf & "点击下一列。取消则开始排序。"
        End If
        
        Set inputRange = Nothing
        On Error Resume Next
        Set inputRange = Application.InputBox(promptMsg, "选择排序列", Type:=8)
        On Error GoTo 0
        
        If inputRange Is Nothing Then
            If isFirst Then Exit Sub
            Exit Do
        End If
        
        Set sortKey = inputRange.Cells(1, 1)
        If isFirst Then
            Set fullTable = sortKey.CurrentRegion
            If fullTable.Rows.count < 2 Then MsgBox "数据不足": Exit Sub
        End If
        
        ' 方向与自定义
        Dim ord As XlSortOrder, cust As String, useCust As Boolean
        Dim resp As VbMsgBoxResult
        cust = "N/A": useCust = False
        
        resp = MsgBox("已选中列：" & Split(sortKey.Address, "$")(1) & vbCrLf & _
                      "【是】升序 (A-Z) | 【否】降序 (Z-A) | 【取消】自定义", vbYesNoCancel + vbQuestion, "方向")
        
        If resp = vbYes Then
            ord = xlAscending
        ElseIf resp = vbNo Then
            ord = xlDescending
        Else
            Dim inputS As String
            inputS = InputBox("请输入自定义序列(逗号隔开):", "自定义序列")
            If Trim(inputS) <> "" Then
                cust = Replace(inputS, "，", ",")
                useCust = True
                ord = xlAscending
            Else
                ord = xlAscending
            End If
        End If
        
        If useCust Then
            ws.Sort.SortFields.Add key:=sortKey, Order:=ord, CustomOrder:=cust, SortOn:=xlSortOnValues
        Else
            ws.Sort.SortFields.Add key:=sortKey, Order:=ord, SortOn:=xlSortOnValues
        End If
        
        ' 记录规则
        Dim singleRule As String
        singleRule = sortKey.Address & "#" & ord & "#" & cust
        If configStr = "" Then configStr = singleRule Else configStr = configStr & "|" & singleRule
        isFirst = False
    Loop
    
    ' 执行排序
    With ws.Sort
        .SetRange fullTable
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' 自动保存 LastUsed 到金库
    Call SaveDataToSafe(KEY_LAST_USED, configStr)
    
    ' 保存为方案
    If MsgBox("是否保存为常用方案？", vbYesNo + vbQuestion, "保存") = vbYes Then
        Dim newName As String
        newName = InputBox("请输入方案名称：", "保存方案")
        If Trim(newName) <> "" Then
            Call SaveDataToSafe(newName, configStr)
            Call UpdateProfileList(newName, "ADD")
            MsgBox "方案 [" & newName & "] 已永久保存。", vbInformation
        End If
    End If
End Sub

' ========================================================
'  核心流程 B：执行方案
' ========================================================
Sub ExecuteProfile(ws As Worksheet, profileName As String, showMsg As Boolean)
    Dim configStr As String
    Dim rules() As String, parts() As String
    Dim i As Integer
    Dim fullTable As Range, sortKey As Range
    
    ' 从金库读取
    configStr = LoadDataFromSafe(profileName)
    
    If configStr = "" Then
        MsgBox "未找到方案数据。", vbCritical: Exit Sub
    End If
    
    ws.Sort.SortFields.Clear
    rules = Split(configStr, "|")
    
    On Error GoTo ErrHandler
    For i = LBound(rules) To UBound(rules)
        parts = Split(rules(i), "#")
        Set sortKey = ws.Range(parts(0))
        If i = LBound(rules) Then Set fullTable = sortKey.CurrentRegion
        
        Dim ord As XlSortOrder: ord = CLng(parts(1))
        Dim cust As String: cust = parts(2)
        
        If cust <> "N/A" And cust <> "" Then
            ws.Sort.SortFields.Add key:=sortKey, Order:=ord, CustomOrder:=cust, SortOn:=xlSortOnValues
        Else
            ws.Sort.SortFields.Add key:=sortKey, Order:=ord, SortOn:=xlSortOnValues
        End If
    Next i
    
    With ws.Sort
        .SetRange fullTable
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Call SaveDataToSafe(KEY_LAST_USED, configStr)
    If showMsg Then MsgBox "已应用方案。", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "列结构不匹配，无法排序。", vbCritical
End Sub

' ========================================================
'  辅助：删除
' ========================================================
Sub DeleteProfileHandler(idxStr As String, profiles() As String)
    On Error Resume Next
    Dim idx As Integer: idx = CInt(idxStr) - 3
    If idx >= LBound(profiles) And idx <= UBound(profiles) Then
        Dim targetName As String: targetName = profiles(idx)
        If MsgBox("确认删除 [" & targetName & "] 吗？", vbYesNo + vbExclamation) = vbYes Then
            Call UpdateProfileList(targetName, "REMOVE")
            MsgBox "已删除。", vbInformation
        End If
    End If
    On Error GoTo 0
End Sub

' ====================================================================
'  【底层引擎：金库存储技术 (V4.0)】
'   原理：自动创建一个 xlSheetVeryHidden 的隐藏表来存数据
' ====================================================================

' 1. 获取(或创建)金库表
Function GetSafeSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SAFE_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' 创建金库表
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SAFE_SHEET_NAME
        ' 关键：设置为深度隐藏，右键菜单里都看不到，防止手贱删除
        ws.Visible = xlSheetVeryHidden
        ' 初始化A列和B列作为 Key-Value 存储
        ws.Cells(1, 1).value = "KEY"
        ws.Cells(1, 2).value = "VALUE"
    End If
    Set GetSafeSheet = ws
End Function

' 2. 写入数据 (Key-Value 模式)
Sub SaveDataToSafe(key As String, value As String)
    Dim ws As Worksheet, rng As Range
    Set ws = GetSafeSheet()
    
    ' 查找 Key
    Set rng = ws.Columns(1).Find(key, LookIn:=xlValues, LookAt:=xlWhole)
    
    If rng Is Nothing Then
        ' 新增
        Dim newRow As Long
        newRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
        ws.Cells(newRow, 1).value = key
        ws.Cells(newRow, 2).value = value
    Else
        ' 更新
        rng.Offset(0, 1).value = value
    End If
End Sub

' 3. 读取数据
Function LoadDataFromSafe(key As String) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSafeSheet()
    
    Set rng = ws.Columns(1).Find(key, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        LoadDataFromSafe = rng.Offset(0, 1).value
    Else
        LoadDataFromSafe = ""
    End If
End Function

' 4. 管理列表
Function GetProfileList() As String
    GetProfileList = LoadDataFromSafe("_List_")
End Function

Sub UpdateProfileList(targetName As String, action As String)
    Dim currentList As String, newList As String
    Dim arr() As String, i As Integer
    
    currentList = GetProfileList()
    
    If action = "ADD" Then
        If InStr(1, "|" & currentList & "|", "|" & targetName & "|") = 0 Then
            If currentList = "" Then newList = targetName Else newList = currentList & "|" & targetName
            Call SaveDataToSafe("_List_", newList)
        End If
    ElseIf action = "REMOVE" Then
        If currentList <> "" Then
            arr = Split(currentList, "|")
            newList = ""
            For i = LBound(arr) To UBound(arr)
                If arr(i) <> targetName Then
                    If newList = "" Then newList = arr(i) Else newList = newList & "|" & arr(i)
                End If
            Next i
            Call SaveDataToSafe("_List_", newList)
        End If
    End If
End Sub

