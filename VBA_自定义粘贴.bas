'''''' 模块: VBA/自定义粘贴  类型: 自定义粘贴.bas  行数: 48
Attribute VB_Name = "自定义粘贴"
Option Explicit

Sub Smart_Paste_Stable()
    ' ==========================================
    ' 稳定版：只粘贴 + 跳过空单元格 (无撤销功能)
    ' 解决了剪贴板丢失报错的问题
    ' ==========================================
    Dim userChoice As VbMsgBoxResult
    
    ' 1. 检查环境
    If Application.CutCopyMode = 0 Then
        MsgBox "剪贴板为空！请先复制(Ctrl+C)数据。", vbExclamation
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "请选择要粘贴的目标单元格。", vbExclamation
        Exit Sub
    End If
    
    ' 2. 检查保护和合并单元格 (WPS敏感点)
    If ActiveSheet.ProtectContents Then MsgBox "工作表受保护，无法粘贴。", vbCritical: Exit Sub
    If IsNull(Selection.MergeCells) Or Selection.MergeCells Then
        MsgBox "目标区域包含合并单元格，无法使用'跳过空单元格'模式。", vbCritical
        Exit Sub
    End If
    
    ' 3. 询问模式
    userChoice = MsgBox("请选择粘贴模式：" & vbCrLf & vbCrLf & _
                        "【是 (Yes)】 ? 只粘贴数值 (保留原格式)" & vbCrLf & _
                        "【否 (No)】  ? 粘贴全部 (含格式)", _
                        vbYesNoCancel + vbQuestion + vbDefaultButton1, "智能粘贴")
    
    If userChoice = vbCancel Then Exit Sub
    
    ' 4. 执行粘贴
    On Error GoTo ErrHandler
    If userChoice = vbYes Then
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
    Else
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
    End If
    Exit Sub

ErrHandler:
    MsgBox "粘贴出错！原因可能是：复制源大小与目标区域形状严重冲突。", vbCritical
End Sub

