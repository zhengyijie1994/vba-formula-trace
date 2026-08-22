'''''' 模块: VBA/保留小数自定义  类型: 保留小数自定义.bas  行数: 94
Attribute VB_Name = "保留小数自定义"
' 通用宏：保留小数_自定义
' 功能：对当前选中单元格区域的数值进行四舍五入或截断处理，保留用户指定的小数位数
' 适用：可保存到加载项 (.xlam) 中，在所有工作簿中调用

Sub 保留小数_自定义()
    Dim rng As Range
    Dim cell As Range
    Dim response As Integer
    Dim useRound As Boolean
    Dim decimalPlaces As Variant
    Dim numPlaces As Integer
    Dim numFormat As String
    
    ' 检查当前选中对象是否为单元格区域
    If Not TypeOf Selection Is Excel.Range Then
        MsgBox "请先选中单元格区域。", vbExclamation, "提示"
        Exit Sub
    End If
    
    Set rng = Selection
    
    ' 如果选中的区域不包含任何单元格（理论上不会发生），则退出
    If rng Is Nothing Then
        MsgBox "未选中有效单元格。", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 1. 让用户选择处理方式：四舍五入还是截断
    response = MsgBox("请选择处理方式：" & vbCrLf & _
                      "点击“是” - 四舍五入" & vbCrLf & _
                      "点击“否” - 直接截断", _
                      vbYesNo + vbQuestion, "选择操作")
    
    If response = vbYes Then
        useRound = True   ' 四舍五入
    ElseIf response = vbNo Then
        useRound = False  ' 截断
    Else
        Exit Sub          ' 用户取消
    End If
    
    ' 2. 让用户输入保留的小数位数
    decimalPlaces = InputBox("请输入要保留的小数位数（非负整数）：", "小数位数", "2")
    
    ' 检查用户是否取消
    If decimalPlaces = "" Then Exit Sub
    
    ' 验证输入是否为有效的非负整数
    If Not IsNumeric(decimalPlaces) Then
        MsgBox "输入无效，请输入数字。", vbExclamation, "错误"
        Exit Sub
    End If
    
    numPlaces = CInt(decimalPlaces)  ' 转换为整数
    If numPlaces < 0 Then
        MsgBox "小数位数不能为负数。", vbExclamation, "错误"
        Exit Sub
    End If
    
    ' 3. 根据小数位数生成显示格式
    If numPlaces = 0 Then
        numFormat = "0"           ' 不显示小数
    Else
        numFormat = "0." & String(numPlaces, "0")   ' 例如 0.00 或 0.000
    End If
    
    ' 关闭屏幕更新以提高运行速度
    Application.ScreenUpdating = False
    
    ' 遍历每个单元格
    On Error Resume Next   ' 忽略错误（如单元格包含错误值、被保护等）
    For Each cell In rng
        ' 跳过空单元格、非数值单元格和包含公式的单元格（可根据需要修改）
        If Not IsEmpty(cell) And IsNumeric(cell.value) And Not cell.HasFormula Then
            If useRound Then
                ' 使用普通四舍五入（WorksheetFunction.Round）
                cell.value = Application.WorksheetFunction.Round(cell.value, numPlaces)
            Else
                ' 使用截断（WorksheetFunction.RoundDown）
                cell.value = Application.WorksheetFunction.RoundDown(cell.value, numPlaces)
            End If
            ' 设置单元格格式为指定小数位数（可选，可根据需要注释掉）
            cell.NumberFormat = numFormat
        End If
    Next cell
    On Error GoTo 0
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    
    MsgBox "处理完成！", vbInformation, "完成"
End Sub

