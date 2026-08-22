'''''' 模块: VBA/frmDeepSeekSettings  类型: frmDeepSeekSettings.frm  行数: 104
Attribute VB_Name = "frmDeepSeekSettings"
Attribute VB_Base = "0{C04AE3D3-614F-4836-B7EF-C48D3A159779}{7B56C994-3621-449C-A441-366F05765E5D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private m_WasSaved As Boolean
Private chkThinking As MSForms.CheckBox

Public Property Get WasSaved() As Boolean
    WasSaved = m_WasSaved
End Property

Private Sub UserForm_Initialize()
    m_WasSaved = False
    Me.Width = 420
    Me.Height = 290
    txtApiKey.value = DeepSeekGetApiKey()
    txtEndpoint.value = DeepSeekGetEndpoint()
    txtModel.value = DeepSeekGetModel()
    chkRemember.value = DeepSeekHasSavedKey()
    lblNotice.Top = 186
    cmdClear.Top = 238
    cmdTest.Top = 238
    cmdSave.Top = 238
    cmdCancel.Top = 238
    Set chkThinking = Me.Controls.Add("Forms.CheckBox.1", "chkThinking", True)
    chkThinking.Caption = "启用模型深度思考（更准确但更慢；关闭可加快响应）"
    chkThinking.Left = 88
    chkThinking.Top = 160
    chkThinking.Width = 307
    chkThinking.Height = 20
    chkThinking.value = DeepSeekGetThinking()
End Sub

Private Sub cmdTest_Click()
    Dim resultText As String
    Dim succeeded As Boolean

    cmdTest.enabled = False
    cmdSave.enabled = False
    cmdTest.Caption = "测试中"
    lblNotice.Caption = "正在连接接口；本次测试会产生少量Token..."
    Me.Repaint
    DoEvents

    succeeded = DeepSeekTestConnection( _
                    CStr(txtApiKey.value), CStr(txtEndpoint.value), _
                    CStr(txtModel.value), resultText)

    cmdTest.enabled = True
    cmdSave.enabled = True
    cmdTest.Caption = "测试连接"
    If succeeded Then
        lblNotice.Caption = "连接成功；详细HTTP和Token用量已显示。"
        MsgBox resultText, vbInformation
    Else
        lblNotice.Caption = "连接测试失败，请检查密钥、地址、模型或余额。"
        MsgBox resultText, vbExclamation
    End If
End Sub

Private Sub cmdSave_Click()
    Dim errorText As String

    If DeepSeekSaveSettings(CStr(txtApiKey.value), _
                            CStr(txtEndpoint.value), _
                            CStr(txtModel.value), _
                            CBool(chkRemember.value), errorText) Then
        Call DeepSeekSetThinking(CBool(chkThinking.value))
        m_WasSaved = True
        Me.Hide
    Else
        MsgBox "设置未保存：" & errorText, vbExclamation
    End If
End Sub

Private Sub cmdClear_Click()
    If MsgBox("确定清除已保存的DeepSeek API Key吗？", _
              vbQuestion + vbYesNo) = vbYes Then
        Call DeepSeekClearSavedKey
        txtApiKey.value = vbNullString
        chkRemember.value = False
    End If
End Sub

Private Sub cmdCancel_Click()
    m_WasSaved = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        m_WasSaved = False
        Me.Hide
    End If
End Sub


