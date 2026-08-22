'''''' 模块: VBA/frmFormulaTrace  类型: frmFormulaTrace.frm  行数: 337
Attribute VB_Name = "frmFormulaTrace"
Attribute VB_Base = "0{CE2BBCF3-F83B-4CF5-82CB-56BE4F14ABB2}{33441C1A-143F-4C05-BF66-CFD59D478C8F}"
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
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

Private Const MIN_FORM_WIDTH As Double = 360
Private Const MIN_FORM_HEIGHT As Double = 250
Private Const COMPACT_BREAKPOINT As Double = 620
Private Const FORM_MARGIN As Double = 8
Private Const BUTTON_GAP As Double = 4

#If VBA7 Then
    Private m_FormHwnd As LongPtr
#Else
    Private m_FormHwnd As Long
#End If

Private m_AdjustingLayout As Boolean

Private Sub UserForm_Initialize()
    cmdReturnFormula.Caption = ChrW$(36820) & ChrW$(22238) & ChrW$(20844) & ChrW$(24335)
    cmdReturnFormula.ControlTipText = "Return to the current formula cell"
    m_AdjustingLayout = True
    Me.Width = 420
    Me.Height = 280
    m_AdjustingLayout = False
    Call LayoutControls
End Sub

Private Sub UserForm_Activate()
    Call EnableResizableBorder
    Call LayoutControls
    TraceRefresh True
End Sub

Private Sub UserForm_Resize()
    If m_AdjustingLayout Then Exit Sub

    On Error GoTo CleanUp
    m_AdjustingLayout = True

    If Me.Width < MIN_FORM_WIDTH Then Me.Width = MIN_FORM_WIDTH
    If Me.Height < MIN_FORM_HEIGHT Then Me.Height = MIN_FORM_HEIGHT

    Call LayoutControls

CleanUp:
    m_AdjustingLayout = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    TraceFormClosed True
End Sub

Private Sub EnableResizableBorder()
#If VBA7 Then
    Dim windowStyle As LongPtr
#Else
    Dim windowStyle As Long
#End If

    On Error Resume Next

    m_FormHwnd = FindWindow("ThunderDFrame", Me.Caption)
    If m_FormHwnd = 0 Then m_FormHwnd = FindWindow("ThunderXFrame", Me.Caption)
    If m_FormHwnd = 0 Then m_FormHwnd = FindWindow(vbNullString, Me.Caption)
    If m_FormHwnd = 0 Then Exit Sub

    windowStyle = GetWindowLongPtr(m_FormHwnd, GWL_STYLE)
    windowStyle = windowStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX
    Call SetWindowLongPtr(m_FormHwnd, GWL_STYLE, windowStyle)

    Call SetWindowPos(m_FormHwnd, 0, 0, 0, 0, 0, _
                      SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)

    On Error GoTo 0
End Sub

Private Sub LayoutControls()
    Dim insideWidth As Double
    Dim insideHeight As Double
    Dim contentWidth As Double
    Dim formulaTop As Double
    Dim formulaHeight As Double
    Dim firstButtonTop As Double
    Dim secondButtonTop As Double

    On Error GoTo LayoutFailed

    insideWidth = Me.insideWidth
    insideHeight = Me.insideHeight
    contentWidth = insideWidth - FORM_MARGIN * 2
    If contentWidth < 100 Then Exit Sub

    ' Compact the information block so the formula keeps most of the space.
    lblStatus.Left = FORM_MARGIN
    lblStatus.Top = 5
    lblStatus.Width = contentWidth
    lblStatus.Height = 17

    lblPath.Left = FORM_MARGIN
    lblPath.Top = 23
    lblPath.Width = contentWidth
    lblPath.Height = 16

    lblCurrent.Left = FORM_MARGIN
    lblCurrent.Top = 40
    lblCurrent.Width = contentWidth
    lblCurrent.Height = 22

    lblHint.Left = FORM_MARGIN
    lblHint.Top = 63
    lblHint.Width = contentWidth
    lblHint.Height = 20

    cmdReturnFormula.Width = 62
    cmdReturnFormula.Height = 18
    cmdReturnFormula.Left = cmdValues.Left - BUTTON_GAP - cmdReturnFormula.Width
    cmdReturnFormula.Top = 81

    cmdValues.Width = 58
    cmdValues.Height = 18
    cmdValues.Left = FORM_MARGIN + contentWidth - cmdValues.Width
    cmdReturnFormula.Left = cmdValues.Left - BUTTON_GAP - cmdReturnFormula.Width
    cmdValues.Top = 81

    cmdRootCause.Width = 62
    cmdRootCause.Height = 18
    cmdRootCause.Left = cmdReturnFormula.Left - BUTTON_GAP - cmdRootCause.Width
    cmdRootCause.Top = 81

    cmdAI.Width = 58
    cmdAI.Height = 18
    cmdAI.Left = cmdRootCause.Left - BUTTON_GAP - cmdAI.Width
    cmdAI.Top = 81

    cmdAISettings.Width = 50
    cmdAISettings.Height = 18
    cmdAISettings.Left = cmdAI.Left - BUTTON_GAP - cmdAISettings.Width
    cmdAISettings.Top = 81

    lblFormulaCaption.Left = FORM_MARGIN
    lblFormulaCaption.Top = 84
    lblFormulaCaption.Width = contentWidth - cmdValues.Width - _
                              cmdRootCause.Width - cmdAI.Width - _
                              cmdAISettings.Width - cmdReturnFormula.Width - BUTTON_GAP * 5
    lblFormulaCaption.Height = 15

    formulaTop = 100
    txtFormula.Left = FORM_MARGIN
    txtFormula.Top = formulaTop
    txtFormula.Width = contentWidth

    If insideWidth < COMPACT_BREAKPOINT Then
        secondButtonTop = insideHeight - cmdNext.Height - 7
        firstButtonTop = secondButtonTop - cmdNext.Height - BUTTON_GAP
        formulaHeight = firstButtonTop - formulaTop - 7
        If formulaHeight < 45 Then formulaHeight = 45
        txtFormula.Height = formulaHeight
        Call LayoutTwoButtonRows(contentWidth, firstButtonTop, secondButtonTop)
    Else
        firstButtonTop = insideHeight - cmdNext.Height - 8
        formulaHeight = firstButtonTop - formulaTop - 8
        If formulaHeight < 55 Then formulaHeight = 55
        txtFormula.Height = formulaHeight
        Call LayoutOneButtonRow(contentWidth, firstButtonTop)
    End If

    Exit Sub

LayoutFailed:
    ' The first resize event can arrive before every control is ready.
End Sub

Private Sub LayoutTwoButtonRows(ByVal contentWidth As Double, _
                                ByVal firstTop As Double, ByVal secondTop As Double)
    Dim rowWidth As Double
    Dim nextLeft As Double

    rowWidth = cmdUseCurrent.Width + cmdPrevious.Width + cmdNext.Width + cmdDrill.Width + _
               BUTTON_GAP * 3
    nextLeft = FORM_MARGIN + (contentWidth - rowWidth) / 2
    PositionButton cmdUseCurrent, nextLeft, firstTop, BUTTON_GAP
    PositionButton cmdPrevious, nextLeft, firstTop, BUTTON_GAP
    PositionButton cmdNext, nextLeft, firstTop, BUTTON_GAP
    PositionButton cmdDrill, nextLeft, firstTop, 0

    rowWidth = cmdBack.Width + cmdMode.Width + cmdRoot.Width + cmdClose.Width + _
               BUTTON_GAP * 3
    nextLeft = FORM_MARGIN + (contentWidth - rowWidth) / 2
    PositionButton cmdBack, nextLeft, secondTop, BUTTON_GAP
    PositionButton cmdMode, nextLeft, secondTop, BUTTON_GAP
    PositionButton cmdRoot, nextLeft, secondTop, BUTTON_GAP
    PositionButton cmdClose, nextLeft, secondTop, 0
End Sub

Private Sub LayoutOneButtonRow(ByVal contentWidth As Double, ByVal buttonTop As Double)
    Dim totalButtonWidth As Double
    Dim gapWidth As Double
    Dim nextLeft As Double

    totalButtonWidth = cmdUseCurrent.Width + cmdPrevious.Width + cmdNext.Width + _
                       cmdDrill.Width + cmdBack.Width + cmdMode.Width + _
                       cmdRoot.Width + cmdClose.Width
    gapWidth = (contentWidth - totalButtonWidth) / 7
    If gapWidth < BUTTON_GAP Then gapWidth = BUTTON_GAP

    nextLeft = FORM_MARGIN
    PositionButton cmdUseCurrent, nextLeft, buttonTop, gapWidth
    PositionButton cmdPrevious, nextLeft, buttonTop, gapWidth
    PositionButton cmdNext, nextLeft, buttonTop, gapWidth
    PositionButton cmdDrill, nextLeft, buttonTop, gapWidth
    PositionButton cmdBack, nextLeft, buttonTop, gapWidth
    PositionButton cmdMode, nextLeft, buttonTop, gapWidth
    PositionButton cmdRoot, nextLeft, buttonTop, gapWidth
    PositionButton cmdClose, nextLeft, buttonTop, 0
End Sub

Private Sub PositionButton(ByVal buttonControl As Object, ByRef nextLeft As Double, _
                           ByVal buttonTop As Double, ByVal gapWidth As Double)
    buttonControl.Left = nextLeft
    buttonControl.Top = buttonTop
    nextLeft = nextLeft + buttonControl.Width + gapWidth
End Sub

Private Sub cmdUseCurrent_Click()
    TraceUseCurrent True
End Sub

Private Sub cmdPrevious_Click()
    TracePrevious True
End Sub

Private Sub cmdNext_Click()
    TraceNext True
End Sub

Private Sub cmdDrill_Click()
    TraceDrill True
End Sub

Private Sub cmdBack_Click()
    TraceBack True
End Sub

Private Sub cmdMode_Click()
    TraceToggleMode True
End Sub

Private Sub cmdRoot_Click()
    TraceRoot True
End Sub

Private Sub cmdAISettings_Click()
    TraceAISettings True
End Sub

Private Sub cmdAI_Click()
    TraceAIChat True
End Sub

Private Sub cmdAI_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                            ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then TraceAISettings True
End Sub

Private Sub cmdRootCause_Click()
    TraceFindRootCause True
End Sub

Private Sub cmdReturnFormula_Click()
    TraceReturnToFormula True
End Sub

Private Sub cmdValues_Click()
    TraceToggleValues True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub









