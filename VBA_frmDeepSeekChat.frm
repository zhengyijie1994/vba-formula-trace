'''''' 模块: VBA/frmDeepSeekChat  类型: frmDeepSeekChat.frm  行数: 514
Attribute VB_Name = "frmDeepSeekChat"
Attribute VB_Base = "0{CDEE5186-EC42-4E01-9B9F-E72E09639243}{62178628-3032-4C2D-841B-56619CCD43F2}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
            ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
            ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
            ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
            ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
        ByVal uFlags As Long) As Long
    Private Declare PtrSafe Function IsIconic Lib "user32" ( _
        ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
        ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
        ByVal uFlags As Long) As Long
    Private Declare Function IsIconic Lib "user32" ( _
        ByVal hwnd As Long) As Long
    Private Declare Function ShowWindow Lib "user32" ( _
        ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" ( _
        ByVal hwnd As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SW_RESTORE As Long = 9
Private Const MIN_FORM_WIDTH As Double = 400
Private Const MIN_FORM_HEIGHT As Double = 450
Private Const FORM_MARGIN As Double = 8

Private Const MAX_TRANSCRIPT_LENGTH As Long = 28000
Private m_HasConversation As Boolean
Private m_Sending As Boolean
Private m_FormulaAnalysisAvailable As Boolean
Private m_AdjustingLayout As Boolean
Private WithEvents cmdShowRequest As MSForms.CommandButton
Attribute cmdShowRequest.VB_VarHelpID = -1

#If VBA7 Then
    Private m_FormHwnd As LongPtr
#Else
    Private m_FormHwnd As Long
#End If

Public Property Get HasConversation() As Boolean
    HasConversation = m_HasConversation
End Property

Public Sub InitializeConversation(ByVal initialText As String, _
                                  ByVal contextTitle As String)
    m_HasConversation = True
    lblContext.Caption = contextTitle
    txtTranscript.value = LimitTranscript(initialText)
    txtQuestion.value = vbNullString
    lblStatus.Caption = "连续对话已就绪；修改会先预览，执行后会自动复查。"
    cmdUndo.enabled = True
    Call RefreshReferences
    Call LayoutControls
    Call ScrollTranscriptToEnd
End Sub

Public Sub ConfigureFormulaAnalysis(ByVal availableValue As Boolean, _
                                    ByVal captionText As String)
    m_FormulaAnalysisAvailable = availableValue
    cmdAnalyzeFormula.Visible = availableValue
    cmdAnalyzeFormula.enabled = availableValue And Not m_Sending
    cmdAnalyzeFormula.Caption = captionText
End Sub

Public Sub SetAnalysisBusy(ByVal busyValue As Boolean, ByVal statusText As String)
    m_Sending = busyValue
    Call SetBusy(busyValue, statusText)
    cmdStop.enabled = busyValue
    If Not busyValue Then cmdAnalyzeFormula.enabled = m_FormulaAnalysisAvailable
End Sub

Public Sub AppendSystemMessage(ByVal messageText As String)
    Call AppendTranscript("【系统】" & vbCrLf & messageText)
End Sub

Public Sub ResetPanel(ByVal statusText As String)
    m_HasConversation = False
    lblContext.Caption = "AI助手面板"
    txtTranscript.value = statusText
    txtQuestion.value = vbNullString
    lstReferences.Clear
    lblStatus.Caption = "面板保持开启；等待新的上下文。"
    cmdAnalyzeFormula.Visible = False
    cmdStop.enabled = False
    Call ScrollTranscriptToEnd
End Sub

Private Sub UserForm_Initialize()
    m_AdjustingLayout = True
    Me.Caption = "AI表格助手"
    Call RestorePanelPosition
    Set cmdShowRequest = Me.Controls.Add("Forms.CommandButton.1", "cmdShowRequest", True)
    cmdShowRequest.Caption = "查看上次请求"
    cmdShowRequest.Width = 88
    cmdShowRequest.Height = 28
    m_AdjustingLayout = False
    Call LayoutControls
End Sub

Private Sub UserForm_Activate()
    Call EnableResizableBorder
    Call LayoutControls
    On Error Resume Next
    txtQuestion.SetFocus
    On Error GoTo 0
End Sub

Private Sub UserForm_Resize()
    If m_AdjustingLayout Then Exit Sub
    If IsFormMinimized() Then Exit Sub
    On Error GoTo CleanUp
    m_AdjustingLayout = True
    If Me.Width < MIN_FORM_WIDTH Then Me.Width = MIN_FORM_WIDTH
    If Me.Height < MIN_FORM_HEIGHT Then Me.Height = MIN_FORM_HEIGHT
    Call LayoutControls
CleanUp:
    m_AdjustingLayout = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Call SavePanelPosition
        Me.Hide
    End If
End Sub

Private Sub EnableResizableBorder()
#If VBA7 Then
    Dim windowStyle As LongPtr
    Dim windowExStyle As LongPtr
#Else
    Dim windowStyle As Long
    Dim windowExStyle As Long
#End If

    On Error Resume Next
    m_FormHwnd = FindWindow("ThunderDFrame", Me.Caption)
    If m_FormHwnd = 0 Then m_FormHwnd = FindWindow("ThunderXFrame", Me.Caption)
    If m_FormHwnd = 0 Then m_FormHwnd = FindWindow(vbNullString, Me.Caption)
    If m_FormHwnd = 0 Then Exit Sub

    windowStyle = GetWindowLongPtr(m_FormHwnd, GWL_STYLE)
    windowStyle = windowStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or _
                  WS_MINIMIZEBOX Or WS_SYSMENU
    Call SetWindowLongPtr(m_FormHwnd, GWL_STYLE, windowStyle)

    windowExStyle = GetWindowLongPtr(m_FormHwnd, GWL_EXSTYLE)
    windowExStyle = (windowExStyle Or WS_EX_APPWINDOW) And Not WS_EX_TOOLWINDOW
    Call SetWindowLongPtr(m_FormHwnd, GWL_EXSTYLE, windowExStyle)

    Call SetWindowPos(m_FormHwnd, 0, 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
    On Error GoTo 0
End Sub

Public Sub RestoreWindow()
    On Error Resume Next
    Call EnableResizableBorder
    If m_FormHwnd = 0 Then Exit Sub
    If IsIconic(m_FormHwnd) <> 0 Then _
        Call ShowWindow(m_FormHwnd, SW_RESTORE)
    Call SetForegroundWindow(m_FormHwnd)
    On Error GoTo 0
End Sub

Private Function IsFormMinimized() As Boolean
    On Error Resume Next
    If m_FormHwnd = 0 Then Call EnableResizableBorder
    If m_FormHwnd <> 0 Then _
        IsFormMinimized = (IsIconic(m_FormHwnd) <> 0)
    On Error GoTo 0
End Function

Private Sub LayoutControls()
    Dim insideWidth As Double
    Dim insideHeight As Double
    Dim contentWidth As Double
    Dim bottomTop As Double
    Dim questionTop As Double
    Dim referenceTop As Double
    Dim referenceHeight As Double
    Dim transcriptHeight As Double
    Dim nextRight As Double

    On Error GoTo Failed
    insideWidth = Me.insideWidth
    insideHeight = Me.insideHeight
    contentWidth = insideWidth - FORM_MARGIN * 2
    If contentWidth < 200 Or insideHeight < 300 Then Exit Sub

    cmdAnalyzeFormula.Left = FORM_MARGIN
    cmdAnalyzeFormula.Top = 6
    cmdStop.Left = cmdAnalyzeFormula.Left + cmdAnalyzeFormula.Width + 6
    cmdStop.Top = 6
    lblContext.Left = cmdStop.Left + cmdStop.Width + 8
    lblContext.Top = 9
    lblContext.Width = insideWidth - lblContext.Left - FORM_MARGIN

    bottomTop = insideHeight - 36
    nextRight = insideWidth - FORM_MARGIN
    cmdClose.Left = nextRight - cmdClose.Width
    cmdClose.Top = bottomTop
    nextRight = cmdClose.Left - 6
    cmdSend.Left = nextRight - cmdSend.Width
    cmdSend.Top = bottomTop
    nextRight = cmdSend.Left - 6
    cmdUndo.Left = nextRight - cmdUndo.Width
    cmdUndo.Top = bottomTop
    nextRight = cmdUndo.Left - 6
    cmdNew.Left = nextRight - cmdNew.Width
    cmdNew.Top = bottomTop
    nextRight = cmdNew.Left - 6
    cmdShowRequest.Left = nextRight - cmdShowRequest.Width
    cmdShowRequest.Top = bottomTop
    lblStatus.Left = FORM_MARGIN
    lblStatus.Top = bottomTop - 2
    lblStatus.Width = cmdShowRequest.Left - FORM_MARGIN - 8
    lblStatus.Height = 34

    txtQuestion.Left = FORM_MARGIN
    txtQuestion.Width = contentWidth
    txtQuestion.Height = 68
    questionTop = bottomTop - txtQuestion.Height - 10
    txtQuestion.Top = questionTop
    lblQuestion.Left = FORM_MARGIN
    lblQuestion.Width = contentWidth
    lblQuestion.Top = questionTop - 19

    referenceHeight = insideHeight * 0.12
    If referenceHeight < 50 Then referenceHeight = 50
    If referenceHeight > 90 Then referenceHeight = 90
    referenceTop = lblQuestion.Top - referenceHeight - 23
    lblReferences.Left = FORM_MARGIN
    lblReferences.Width = contentWidth
    lblReferences.Top = referenceTop - 19
    lstReferences.Left = FORM_MARGIN
    lstReferences.Top = referenceTop
    lstReferences.Width = contentWidth
    lstReferences.Height = referenceHeight

    txtTranscript.Left = FORM_MARGIN
    txtTranscript.Top = 36
    txtTranscript.Width = contentWidth
    transcriptHeight = lblReferences.Top - txtTranscript.Top - 7
    If transcriptHeight < 100 Then transcriptHeight = 100
    txtTranscript.Height = transcriptHeight
    Exit Sub
Failed:
End Sub

Private Sub RestorePanelPosition()
    Dim savedWidth As Double
    Dim savedHeight As Double
    Dim savedLeft As Double
    Dim savedTop As Double
    Dim appLeft As Double
    Dim appTop As Double
    Dim appRight As Double
    Dim appBottom As Double

    On Error Resume Next
    savedWidth = val(GetSetting("FormulaTraceAI", "PanelLayout", "Width", "0"))
    savedHeight = val(GetSetting("FormulaTraceAI", "PanelLayout", "Height", "0"))
    savedLeft = val(GetSetting("FormulaTraceAI", "PanelLayout", "Left", "0"))
    savedTop = val(GetSetting("FormulaTraceAI", "PanelLayout", "Top", "0"))
    If savedWidth >= MIN_FORM_WIDTH Then Me.Width = savedWidth
    If savedHeight >= MIN_FORM_HEIGHT Then Me.Height = savedHeight
    Me.StartUpPosition = 0

    appLeft = Application.Left
    appTop = Application.Top
    appRight = appLeft + Application.Width
    appBottom = appTop + Application.Height
    ' 修复：保存的位置须与宿主窗口有足够重合（负坐标/被拔掉的副屏都算失效），
    ' 否则回退到默认位置，避免窗体出现在屏幕外不可见。
    If savedLeft + Me.Width - 300 < appLeft Or savedLeft > appRight - 300 Then savedLeft = 0
    If savedTop + 60 < appTop Or savedTop > appBottom - 60 Then savedTop = 0
    If savedLeft = 0 Then savedLeft = appRight - Me.Width - 8
    If savedTop = 0 Then savedTop = appTop + 28
    Me.Left = savedLeft
    Me.Top = savedTop
    On Error GoTo 0
End Sub

Private Sub SavePanelPosition()
    On Error Resume Next
    SaveSetting "FormulaTraceAI", "PanelLayout", "Width", CStr(Me.Width)
    SaveSetting "FormulaTraceAI", "PanelLayout", "Height", CStr(Me.Height)
    SaveSetting "FormulaTraceAI", "PanelLayout", "Left", CStr(Me.Left)
    SaveSetting "FormulaTraceAI", "PanelLayout", "Top", CStr(Me.Top)
    On Error GoTo 0
End Sub

Private Sub cmdAnalyzeFormula_Click()
    If m_Sending Then Exit Sub
    Call SetAnalysisBusy(True, "正在解析当前公式...")
    TraceAIPanelAnalyze True
End Sub

Private Sub cmdStop_Click()
    If Not m_Sending Then Exit Sub
    lblStatus.Caption = "正在取消请求..."
    cmdStop.enabled = False
    Call AIWorksheetAssistantCancel
End Sub

Private Sub cmdSend_Click()
    Dim questionText As String
    Dim replyText As String
    Dim errorText As String

    If m_Sending Then Exit Sub
    questionText = Trim$(txtQuestion.value)
    If Len(questionText) = 0 Then
        MsgBox "请输入问题或操作要求。", vbInformation
        Exit Sub
    End If
    If Len(questionText) > 6000 Then
        MsgBox "单条消息不能超过6000个字符。", vbInformation
        Exit Sub
    End If

    m_Sending = True
    SetBusy True, "AI正在思考，请稍候..."
    AppendTranscript "【你】" & vbCrLf & questionText
    txtQuestion.value = vbNullString
    DoEvents

    If AIWorksheetAssistantSend(questionText, replyText, errorText) Then
        AppendTranscript "【AI】" & vbCrLf & replyText
        SetBusy False, "已完成；双击下方引用可直接跳转。"
    Else
        AppendTranscript "【错误】" & vbCrLf & errorText
        SetBusy False, "本轮失败，未确认的操作不会执行。"
        MsgBox errorText, vbExclamation
    End If
    m_Sending = False
End Sub

Private Sub cmdShowRequest_Click()
    Dim reqText As String
    reqText = DeepSeekGetLastRequestBody()
    If Len(reqText) = 0 Then
        AppendTranscript "【请求数据】尚未发送过请求。"
    Else
        AppendTranscript "【请求数据】" & vbCrLf & reqText
    End If
End Sub

Private Sub cmdNew_Click()
    Dim initialText As String
    Dim contextTitle As String
    Dim errorText As String

    If m_Sending Then Exit Sub
    If MsgBox("确定清空当前问答并开始新会话吗？" & vbCrLf & _
              "不会重新调用首次公式分析，也不会影响撤销快照。", _
              vbQuestion + vbYesNo) <> vbYes Then Exit Sub

    If AIWorksheetAssistantNewSession(initialText, contextTitle, errorText) Then
        InitializeConversation initialText, contextTitle
    Else
        MsgBox errorText, vbExclamation
    End If
End Sub

Private Sub cmdUndo_Click()
    Dim resultText As String
    Dim errorText As String

    If m_Sending Then Exit Sub
    If MsgBox("确定撤销最近一次AI表格操作吗？", _
              vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    SetBusy True, "正在恢复撤销快照..."
    If AIWorksheetAssistantUndo(resultText, errorText) Then
        AppendTranscript "【系统】" & vbCrLf & resultText
        SetBusy False, "撤销完成。"
    Else
        SetBusy False, "没有完成撤销。"
        MsgBox errorText, vbInformation
    End If
End Sub

Private Sub cmdClose_Click()
    Call SavePanelPosition
    Me.Hide
End Sub

Private Sub lstReferences_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim referenceText As String
    Dim errorText As String

    If lstReferences.ListIndex < 0 Then Exit Sub
    referenceText = CStr(lstReferences.List(lstReferences.ListIndex))
    If AIWorksheetAssistantJumpToReference(referenceText, errorText) Then
        lblStatus.Caption = "已跳转：" & referenceText
    Else
        MsgBox errorText, vbExclamation
    End If
End Sub

Private Sub SetBusy(ByVal busyValue As Boolean, ByVal statusText As String)
    cmdSend.enabled = Not busyValue
    cmdNew.enabled = Not busyValue
    cmdUndo.enabled = Not busyValue
    cmdAnalyzeFormula.enabled = (Not busyValue) And m_FormulaAnalysisAvailable
    cmdStop.enabled = busyValue
    lstReferences.enabled = Not busyValue
    txtQuestion.enabled = Not busyValue
    lblStatus.Caption = statusText
    Me.Repaint
End Sub

Private Sub AppendTranscript(ByVal blockText As String)
    Dim newText As String

    If Len(txtTranscript.value) > 0 Then
        newText = txtTranscript.value & vbCrLf & vbCrLf & blockText
    Else
        newText = blockText
    End If
    txtTranscript.value = LimitTranscript(newText)
    Call RefreshReferences
    Call ScrollTranscriptToEnd
End Sub

Private Sub ScrollTranscriptToEnd()
    Dim questionStart As Long

    On Error Resume Next
    questionStart = txtQuestion.SelStart
    txtTranscript.SetFocus
    txtTranscript.SelStart = Len(txtTranscript.value)
    txtTranscript.SelLength = 0
    Me.Repaint
    DoEvents
    If txtQuestion.enabled Then
        txtQuestion.SetFocus
        If questionStart > Len(txtQuestion.value) Then _
            questionStart = Len(txtQuestion.value)
        txtQuestion.SelStart = questionStart
    End If
    On Error GoTo 0
End Sub

Private Sub RefreshReferences()
    On Error Resume Next
    Call AIWorksheetAssistantPopulateReferences(txtTranscript.value, lstReferences)
    If lstReferences.ListCount = 0 Then
        lblReferences.Caption = "可跳转引用：暂无"
    Else
        lblReferences.Caption = "可跳转引用（双击定位）：" & _
                                CStr(lstReferences.ListCount) & " 项"
    End If
    On Error GoTo 0
End Sub

Private Function LimitTranscript(ByVal textValue As String) As String
    If Len(textValue) <= MAX_TRANSCRIPT_LENGTH Then
        LimitTranscript = textValue
    Else
        LimitTranscript = "……较早对话已在显示区折叠……" & vbCrLf & vbCrLf & _
                          Right$(textValue, MAX_TRANSCRIPT_LENGTH - 24)
    End If
End Function



