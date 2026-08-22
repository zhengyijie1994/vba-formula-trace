'''''' 模块: VBA/frmAIActionPreview  类型: frmAIActionPreview.frm  行数: 40
Attribute VB_Name = "frmAIActionPreview"
Attribute VB_Base = "0{2DABB077-4930-4E56-AC3E-80F35751D9E5}{4BF5ACC6-C7D5-485C-8381-2491CCA1EA8C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private m_WasApproved As Boolean

Public Property Get WasApproved() As Boolean
    WasApproved = m_WasApproved
End Property

Public Sub InitializePreview(ByVal summaryText As String, _
                             ByVal operationsText As String, _
                             ByVal impactText As String)
    m_WasApproved = False
    lblSummary.Caption = summaryText
    txtOperations.value = operationsText
    lblImpact.Caption = impactText & "；执行前将创建一次撤销快照。"
End Sub

Private Sub cmdExecute_Click()
    m_WasApproved = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    m_WasApproved = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    m_WasApproved = False
End Sub


