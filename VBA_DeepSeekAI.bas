'''''' 模块: VBA/DeepSeekAI  类型: DeepSeekAI.bas  行数: 1151
Attribute VB_Name = "DeepSeekAI"
Option Explicit

Private Const SETTINGS_APP As String = "FormulaTraceDeepSeek"
Private Const SETTINGS_SECTION As String = "Settings"
Private Const KEY_API As String = "ApiKeyProtected"
Private Const KEY_ENDPOINT As String = "Endpoint"
Private Const KEY_MODEL As String = "Model"
Private Const KEY_THINKING As String = "Thinking"

Private Const DEFAULT_ENDPOINT As String = "https://api.deepseek.com/chat/completions"
Private Const DEFAULT_MODEL As String = "deepseek-v4-pro"
Private Const REQUEST_TIMEOUT_SECONDS As Long = 120
Private Const CRYPTPROTECT_UI_FORBIDDEN As Long = &H1
Private Const LMEM_FIXED As Long = &H0

Private Type DATA_BLOB
    cbData As Long
#If Win64 Then
    alignmentPadding As Long
#End If
#If VBA7 Then
    pbData As LongPtr
#Else
    pbData As Long
#End If
End Type

#If VBA7 Then
    Private Declare PtrSafe Function CryptProtectData Lib "crypt32.dll" ( _
        ByRef pDataIn As DATA_BLOB, ByVal szDataDescr As LongPtr, _
        ByVal pOptionalEntropy As LongPtr, ByVal pvReserved As LongPtr, _
        ByVal pPromptStruct As LongPtr, ByVal dwFlags As Long, _
        ByRef pDataOut As DATA_BLOB) As Long
    Private Declare PtrSafe Function CryptUnprotectData Lib "crypt32.dll" ( _
        ByRef pDataIn As DATA_BLOB, ByVal ppszDataDescr As LongPtr, _
        ByVal pOptionalEntropy As LongPtr, ByVal pvReserved As LongPtr, _
        ByVal pPromptStruct As LongPtr, ByVal dwFlags As Long, _
        ByRef pDataOut As DATA_BLOB) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef destination As Any, ByRef source As Any, ByVal byteCount As LongPtr)
    Private Declare PtrSafe Function LocalFree Lib "kernel32" ( _
        ByVal memoryHandle As LongPtr) As LongPtr
    Private Declare PtrSafe Function LocalAlloc Lib "kernel32" ( _
        ByVal allocationFlags As Long, ByVal byteCount As LongPtr) As LongPtr
#Else
    Private Declare Function CryptProtectData Lib "crypt32.dll" ( _
        ByRef pDataIn As DATA_BLOB, ByVal szDataDescr As Long, _
        ByVal pOptionalEntropy As Long, ByVal pvReserved As Long, _
        ByVal pPromptStruct As Long, ByVal dwFlags As Long, _
        ByRef pDataOut As DATA_BLOB) As Long
    Private Declare Function CryptUnprotectData Lib "crypt32.dll" ( _
        ByRef pDataIn As DATA_BLOB, ByVal ppszDataDescr As Long, _
        ByVal pOptionalEntropy As Long, ByVal pvReserved As Long, _
        ByVal pPromptStruct As Long, ByVal dwFlags As Long, _
        ByRef pDataOut As DATA_BLOB) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef destination As Any, ByRef source As Any, ByVal byteCount As Long)
    Private Declare Function LocalFree Lib "kernel32" ( _
        ByVal memoryHandle As Long) As Long
    Private Declare Function LocalAlloc Lib "kernel32" ( _
        ByVal allocationFlags As Long, ByVal byteCount As Long) As Long
#End If

Private m_ApiKey As String
Private m_Endpoint As String
Private m_Model As String

Private Const CHAT_MAX_MESSAGES As Long = 20
Private Const CHAT_MAX_SYSTEM As Long = 18000
Private Const CHAT_MAX_CONTEXT As Long = 18000
Private Const CHAT_MAX_MESSAGE As Long = 9000

Private m_ChatSystem As String
Private m_ChatContext As String
Private m_ChatInitialAssistant As String
Private m_ChatRoles(1 To CHAT_MAX_MESSAGES) As String
Private m_ChatContents(1 To CHAT_MAX_MESSAGES) As String
Private m_ChatCount As Long
Private m_ActiveRequest As Object
Private m_CancelRequested As Boolean
Private m_ThinkingStart As Single
Private m_LastRequestBody As String

Public Sub ConfigureDeepSeekAI()
    Call DeepSeekShowSettings
End Sub

Public Function DeepSeekShowSettings() As Boolean
    On Error GoTo Failed

    Load frmDeepSeekSettings
    frmDeepSeekSettings.Show vbModal
    DeepSeekShowSettings = frmDeepSeekSettings.WasSaved
    Unload frmDeepSeekSettings
    Exit Function

Failed:
    On Error Resume Next
    Unload frmDeepSeekSettings
    On Error GoTo 0
    MsgBox "无法打开DeepSeek设置：" & Err.Description, vbExclamation
End Function

Public Function DeepSeekEnsureConfigured() As Boolean
    If Len(DeepSeekGetApiKey()) > 0 Then
        DeepSeekEnsureConfigured = True
    Else
        DeepSeekEnsureConfigured = DeepSeekShowSettings()
        If DeepSeekEnsureConfigured Then
            DeepSeekEnsureConfigured = (Len(DeepSeekGetApiKey()) > 0)
        End If
    End If
End Function

Public Function DeepSeekGetApiKey() As String
    Dim protectedText As String

    If Len(m_ApiKey) > 0 Then
        DeepSeekGetApiKey = m_ApiKey
        Exit Function
    End If

    protectedText = GetSetting(SETTINGS_APP, SETTINGS_SECTION, KEY_API, vbNullString)
    If Len(protectedText) > 0 Then m_ApiKey = UnprotectSecret(protectedText)
    DeepSeekGetApiKey = m_ApiKey
End Function

Public Function DeepSeekGetEndpoint() As String
    If Len(m_Endpoint) = 0 Then
        m_Endpoint = Trim$(GetSetting(SETTINGS_APP, SETTINGS_SECTION, _
                                     KEY_ENDPOINT, DEFAULT_ENDPOINT))
        If Len(m_Endpoint) = 0 Then m_Endpoint = DEFAULT_ENDPOINT
        m_Endpoint = NormalizeDeepSeekEndpoint(m_Endpoint)
    End If
    DeepSeekGetEndpoint = m_Endpoint
End Function

Public Function DeepSeekGetModel() As String
    If Len(m_Model) = 0 Then
        m_Model = Trim$(GetSetting(SETTINGS_APP, SETTINGS_SECTION, _
                                  KEY_MODEL, DEFAULT_MODEL))
        If Len(m_Model) = 0 Then m_Model = DEFAULT_MODEL
    End If
    DeepSeekGetModel = m_Model
End Function

Public Function DeepSeekHasSavedKey() As Boolean
    DeepSeekHasSavedKey = _
        (Len(GetSetting(SETTINGS_APP, SETTINGS_SECTION, KEY_API, vbNullString)) > 0)
End Function

Public Function DeepSeekSaveSettings(ByVal apiKey As String, _
                                     ByVal endpoint As String, _
                                     ByVal modelName As String, _
                                     ByVal rememberKey As Boolean, _
                                     ByRef errorText As String) As Boolean
    Dim protectedText As String

    On Error GoTo Failed

    apiKey = Trim$(apiKey)
    endpoint = Trim$(endpoint)
    endpoint = NormalizeDeepSeekEndpoint(endpoint)
    modelName = Trim$(modelName)

    If Len(apiKey) = 0 Then
        errorText = "请输入DeepSeek API Key。"
        Exit Function
    End If
    If LCase$(Left$(endpoint, 7)) <> "http://" And _
       LCase$(Left$(endpoint, 8)) <> "https://" Then
        errorText = "接口地址必须以 http:// 或 https:// 开头。"
        Exit Function
    End If
    If Len(modelName) = 0 Then
        errorText = "请输入模型名称。"
        Exit Function
    End If

    m_ApiKey = apiKey
    m_Endpoint = endpoint
    m_Model = modelName

    SaveSetting SETTINGS_APP, SETTINGS_SECTION, KEY_ENDPOINT, endpoint
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, KEY_MODEL, modelName

    If rememberKey Then
        protectedText = ProtectSecret(apiKey)
        If Len(protectedText) = 0 Then
            errorText = "API Key加密失败，未保存。"
            Exit Function
        End If
        SaveSetting SETTINGS_APP, SETTINGS_SECTION, KEY_API, protectedText
    Else
        On Error Resume Next
        DeleteSetting SETTINGS_APP, SETTINGS_SECTION, KEY_API
        On Error GoTo Failed
    End If

    errorText = vbNullString
    DeepSeekSaveSettings = True
    Exit Function

Failed:
    errorText = Err.Description
End Function

Public Sub DeepSeekClearSavedKey()
    On Error Resume Next
    DeleteSetting SETTINGS_APP, SETTINGS_SECTION, KEY_API
    On Error GoTo 0
    m_ApiKey = vbNullString
End Sub

Public Sub DeepSeekClearAllSettings()
    On Error Resume Next
    DeleteSetting SETTINGS_APP, SETTINGS_SECTION
    On Error GoTo 0
    m_ApiKey = vbNullString
    m_Endpoint = vbNullString
    m_Model = vbNullString
End Sub

Public Function DeepSeekGetThinking() As Boolean
    DeepSeekGetThinking = (GetSetting(SETTINGS_APP, SETTINGS_SECTION, _
                                  KEY_THINKING, "1") = "1")
End Function

Public Sub DeepSeekSetThinking(ByVal enabled As Boolean)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, KEY_THINKING, CStr(Abs(enabled))
End Sub

Public Sub DeepSeekResetThinkingClock()
    m_ThinkingStart = 0
End Sub

Public Function DeepSeekGetLastRequestBody() As String
    DeepSeekGetLastRequestBody = m_LastRequestBody
End Function

Public Function DeepSeekAnalyzePrompt(ByVal promptText As String, _
                                      ByRef analysisText As String, _
                                      ByRef errorText As String, _
                                      ByRef auditText As String) As Boolean
    DeepSeekAnalyzePrompt = SendDeepSeekRequest( _
        promptText, DeepSeekGetApiKey(), DeepSeekGetEndpoint(), _
        DeepSeekGetModel(), False, analysisText, errorText, auditText)
End Function

Public Function DeepSeekTestConnection(ByVal apiKey As String, _
                                       ByVal endpoint As String, _
                                       ByVal modelName As String, _
                                       ByRef resultText As String) As Boolean
    Dim responseContent As String
    Dim errorText As String
    Dim auditText As String

    apiKey = Trim$(apiKey)
    endpoint = NormalizeDeepSeekEndpoint(endpoint)
    modelName = Trim$(modelName)
    If Len(apiKey) = 0 Then resultText = "请输入API Key后再测试。": Exit Function
    If LCase$(Left$(endpoint, 7)) <> "http://" And _
       LCase$(Left$(endpoint, 8)) <> "https://" Then
        resultText = "接口地址必须以 http:// 或 https:// 开头。"
        Exit Function
    End If
    If Len(modelName) = 0 Then resultText = "请输入模型名称后再测试。": Exit Function

    DeepSeekTestConnection = SendDeepSeekRequest( _
        "这是一次API连接测试。请按要求返回成功状态。", _
        apiKey, endpoint, modelName, True, responseContent, errorText, auditText)
    If DeepSeekTestConnection Then
        resultText = "连接成功。测试请求会产生少量Token。" & vbCrLf & vbCrLf & auditText
    Else
        resultText = "连接失败：" & errorText
        If Len(auditText) > 0 Then resultText = resultText & vbCrLf & vbCrLf & auditText
    End If
End Function

Private Function SendDeepSeekRequest(ByVal promptText As String, _
                                     ByVal apiKey As String, _
                                     ByVal endpoint As String, _
                                     ByVal modelName As String, _
                                     ByVal connectionTest As Boolean, _
                                     ByRef analysisText As String, _
                                     ByRef errorText As String, _
                                     ByRef auditText As String) As Boolean
    Dim request As Object
    Dim requestBody As String
    Dim responseText As String
    Dim completed As Boolean
    Dim statusCode As Long
    Dim searchStart As Long
    Dim startTimer As Single
    Dim elapsedSeconds As Double

    On Error GoTo Failed
    analysisText = vbNullString: errorText = vbNullString: auditText = vbNullString
    If Len(apiKey) = 0 Then errorText = "尚未配置DeepSeek API Key。": Exit Function

    requestBody = BuildRequestBody(modelName, promptText, connectionTest)
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set m_ActiveRequest = request
    m_CancelRequested = False
    request.SetTimeouts 10000, 10000, 30000, REQUEST_TIMEOUT_SECONDS * 1000
    request.Open "POST", endpoint, True
    request.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    request.SetRequestHeader "Authorization", "Bearer " & apiKey
    request.SetRequestHeader "User-Agent", "Excel-Formula-Trace-AI/2.0"
    startTimer = Timer
    request.Send StringToUtf8Bytes(requestBody)

    Do
        If m_CancelRequested Then
            request.Abort
            errorText = "用户已取消请求。"
            Set m_ActiveRequest = Nothing
            Exit Function
        End If
        completed = CBool(request.WaitForResponse(1))
        DoEvents
        If completed Then Exit Do
        If GetElapsedSeconds(startTimer) >= REQUEST_TIMEOUT_SECONDS Then
            request.Abort
            errorText = "请求超时（" & CStr(REQUEST_TIMEOUT_SECONDS) & "秒）。"
            auditText = BuildAuditText(vbNullString, 0, endpoint, GetElapsedSeconds(startTimer))
            Exit Function
        End If
    Loop

    elapsedSeconds = GetElapsedSeconds(startTimer)
    statusCode = CLng(request.Status)
    responseText = Utf8BytesToString(request.ResponseBody)
    auditText = BuildAuditText(responseText, statusCode, endpoint, elapsedSeconds)
    If statusCode < 200 Or statusCode >= 300 Then
        errorText = JsonReadString(responseText, "message", 1)
        If Len(errorText) = 0 Then errorText = LimitText(responseText, 800)
        errorText = "HTTP " & CStr(statusCode) & "：" & errorText
        Exit Function
    End If

    searchStart = InStr(1, responseText, """choices""", vbBinaryCompare)
    If searchStart = 0 Then searchStart = 1
    analysisText = JsonReadString(responseText, "content", searchStart)
    If Len(Trim$(analysisText)) = 0 Then
        errorText = "DeepSeek返回成功，但没有可显示的分析内容。"
        Exit Function
    End If
    SendDeepSeekRequest = True
    Exit Function
Failed:
    errorText = Err.Description
    auditText = BuildAuditText(vbNullString, 0, endpoint, GetElapsedSeconds(startTimer))
End Function

Public Function DeepSeekFormatStructuredAnalysis(ByVal rawAnalysis As String, _
                                                 ByRef formattedText As String, _
                                                 ByRef errorText As String) As Boolean
    Dim jsonText As String, conclusionText As String, expressionText As String
    Dim evidenceText As String, suggestionsText As String
    Dim confidenceText As String, missingText As String

    jsonText = ExtractJsonObject(rawAnalysis)
    If Len(jsonText) = 0 Then errorText = "AI没有返回JSON对象。": Exit Function
    conclusionText = JsonReadString(jsonText, "conclusion", 1)
    expressionText = JsonReadString(jsonText, "error_expression", 1)
    evidenceText = JsonReadStringArray(jsonText, "evidence")
    suggestionsText = JsonReadStringArray(jsonText, "suggestions")
    confidenceText = JsonReadString(jsonText, "confidence", 1)
    missingText = JsonReadStringArray(jsonText, "missing_information")
    If Len(conclusionText) = 0 Then
        errorText = "AI返回的JSON缺少conclusion字段。"
        Exit Function
    End If

    formattedText = "【结论】" & vbCrLf & conclusionText
    If Len(expressionText) > 0 Then formattedText = formattedText & vbCrLf & vbCrLf & "【出错子表达式】" & vbCrLf & expressionText
    If Len(evidenceText) > 0 Then formattedText = formattedText & vbCrLf & vbCrLf & "【证据】" & vbCrLf & evidenceText
    If Len(suggestionsText) > 0 Then formattedText = formattedText & vbCrLf & vbCrLf & "【修复建议】" & vbCrLf & suggestionsText
    If Len(confidenceText) > 0 Then formattedText = formattedText & vbCrLf & vbCrLf & "【置信度】" & vbCrLf & confidenceText
    If Len(missingText) > 0 Then formattedText = formattedText & vbCrLf & vbCrLf & "【仍缺少的信息】" & vbCrLf & missingText
    errorText = vbNullString
    DeepSeekFormatStructuredAnalysis = True
End Function

Public Sub DeepSeekCancelActiveRequest()
    m_CancelRequested = True
    On Error Resume Next
    If Not m_ActiveRequest Is Nothing Then m_ActiveRequest.Abort
    On Error GoTo 0
End Sub

Public Function DeepSeekRequestInProgress() As Boolean
    DeepSeekRequestInProgress = Not (m_ActiveRequest Is Nothing)
End Function

Public Sub DeepSeekConversationBegin(ByVal systemText As String, _
                                     ByVal contextText As String, _
                                     ByVal initialAssistant As String)
    Call DeepSeekConversationReset
    m_ChatSystem = LimitText(systemText, CHAT_MAX_SYSTEM)
    m_ChatContext = LimitText(contextText, CHAT_MAX_CONTEXT)
    m_ChatInitialAssistant = LimitText(initialAssistant, CHAT_MAX_MESSAGE)
End Sub

Public Sub DeepSeekConversationReset()
    Dim i As Long

    m_ChatSystem = vbNullString
    m_ChatContext = vbNullString
    m_ChatInitialAssistant = vbNullString
    For i = 1 To CHAT_MAX_MESSAGES
        m_ChatRoles(i) = vbNullString
        m_ChatContents(i) = vbNullString
    Next i
    m_ChatCount = 0
End Sub

Public Function DeepSeekConversationReady() As Boolean
    DeepSeekConversationReady = (Len(m_ChatSystem) > 0)
End Function

Public Function DeepSeekConversationSend(ByVal userText As String, _
                                         ByRef assistantText As String, _
                                         ByRef errorText As String, _
                                         ByRef auditText As String) As Boolean
    Dim requestBody As String

    If Not DeepSeekConversationReady() Then
        errorText = "AI会话尚未初始化。"
        Exit Function
    End If
    If Len(Trim$(userText)) = 0 Then
        errorText = "消息不能为空。"
        Exit Function
    End If

    requestBody = BuildConversationRequestBody( _
                      DeepSeekGetModel(), LimitText(userText, CHAT_MAX_MESSAGE))
    m_LastRequestBody = requestBody
    If Not SendDeepSeekConversationRequest( _
               requestBody, DeepSeekGetApiKey(), DeepSeekGetEndpoint(), _
               assistantText, errorText, auditText) Then
        Call ClearThinkingStatus
        Exit Function
    End If
    Call ClearThinkingStatus


    Call AppendConversationMessage("user", userText)
    Call AppendConversationMessage("assistant", assistantText)
    DeepSeekConversationSend = True
End Function

Private Sub AppendConversationMessage(ByVal roleName As String, _
                                      ByVal contentText As String)
    Dim i As Long

    If m_ChatCount >= CHAT_MAX_MESSAGES Then
        For i = 3 To CHAT_MAX_MESSAGES
            m_ChatRoles(i - 2) = m_ChatRoles(i)
            m_ChatContents(i - 2) = m_ChatContents(i)
        Next i
        m_ChatCount = CHAT_MAX_MESSAGES - 2
    End If

    m_ChatCount = m_ChatCount + 1
    m_ChatRoles(m_ChatCount) = roleName
    m_ChatContents(m_ChatCount) = LimitText(contentText, CHAT_MAX_MESSAGE)
End Sub

Private Function BuildConversationRequestBody(ByVal modelName As String, _
                                              ByVal currentUserText As String) As String
    Dim messagesText As String
    Dim optionText As String
    Dim i As Long

    messagesText = "{""role"":""system"",""content"":""" & _
                   JsonEscape(m_ChatSystem) & """}"

    If Len(m_ChatContext) > 0 Then
        messagesText = messagesText & ",{""role"":""user"",""content"":""" & _
                       JsonEscape(m_ChatContext) & """}"
        If Len(m_ChatInitialAssistant) > 0 Then
            messagesText = messagesText & _
                           ",{""role"":""assistant"",""content"":""" & _
                           JsonEscape(m_ChatInitialAssistant) & """}"
        End If
    End If

    For i = 1 To m_ChatCount
        messagesText = messagesText & ",{""role"":""" & _
                       JsonEscape(m_ChatRoles(i)) & _
                       """,""content"":""" & _
                       JsonEscape(m_ChatContents(i)) & """}"
    Next i

    messagesText = messagesText & ",{""role"":""user"",""content"":""" & _
                   JsonEscape(currentUserText) & """}"

    If DeepSeekGetThinking() Then
        optionText = """thinking"":{""type"":""enabled""},""reasoning_effort"":""high"","
    Else
        optionText = vbNullString
    End If
    BuildConversationRequestBody = _
        "{""model"":""" & JsonEscape(modelName) & _
        """,""messages"":[" & messagesText & _
        "],""response_format"":{""type"":""json_object""}," & _
        optionText & """stream"":false,""max_tokens"":10000}"
End Function

Private Function SendDeepSeekConversationRequest(ByVal requestBody As String, _
                                                 ByVal apiKey As String, _
                                                 ByVal endpoint As String, _
                                                 ByRef assistantText As String, _
                                                 ByRef errorText As String, _
                                                 ByRef auditText As String) As Boolean
    Dim request As Object
    Dim responseText As String
    Dim completed As Boolean
    Dim statusCode As Long
    Dim searchStart As Long
    Dim startTimer As Single
    Dim elapsedSeconds As Double
    Dim lastThinkingSecond As Long

    On Error GoTo Failed
    assistantText = vbNullString
    errorText = vbNullString
    auditText = vbNullString
    If Len(apiKey) = 0 Then
        errorText = "尚未配置DeepSeek API Key。"
        Exit Function
    End If

    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set m_ActiveRequest = request
    m_CancelRequested = False
    request.SetTimeouts 10000, 10000, 30000, REQUEST_TIMEOUT_SECONDS * 1000
    request.Open "POST", endpoint, True
    request.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    request.SetRequestHeader "Authorization", "Bearer " & apiKey
    request.SetRequestHeader "User-Agent", "Excel-Formula-Trace-AI/3.0"
    startTimer = Timer
    If m_ThinkingStart = 0 Then m_ThinkingStart = startTimer
    request.Send StringToUtf8Bytes(requestBody)

    Do
        If m_CancelRequested Then
            request.Abort
            errorText = "用户已取消请求。"
            Set m_ActiveRequest = Nothing
            Exit Function
        End If
        completed = CBool(request.WaitForResponse(1))
        DoEvents
        If completed Then Exit Do
        elapsedSeconds = GetElapsedSeconds(m_ThinkingStart)
        If CLng(elapsedSeconds) > lastThinkingSecond Then
            lastThinkingSecond = CLng(elapsedSeconds)
            Call SetThinkingStatus(CLng(elapsedSeconds))
        End If

        If GetElapsedSeconds(startTimer) >= REQUEST_TIMEOUT_SECONDS Then
            request.Abort
            errorText = "请求超时（" & CStr(REQUEST_TIMEOUT_SECONDS) & "秒）。"
            auditText = BuildAuditText(vbNullString, 0, endpoint, _
                                        GetElapsedSeconds(startTimer))
            Exit Function
        End If
    Loop

    elapsedSeconds = GetElapsedSeconds(startTimer)
    statusCode = CLng(request.Status)
    responseText = Utf8BytesToString(request.ResponseBody)
    auditText = BuildAuditText(responseText, statusCode, endpoint, elapsedSeconds)
    If statusCode < 200 Or statusCode >= 300 Then
        errorText = JsonReadString(responseText, "message", 1)
        If Len(errorText) = 0 Then errorText = LimitText(responseText, 800)
        errorText = "HTTP " & CStr(statusCode) & "：" & errorText
        Exit Function
    End If

    searchStart = InStr(1, responseText, """choices""", vbBinaryCompare)
    If searchStart = 0 Then searchStart = 1
    assistantText = JsonReadString(responseText, "content", searchStart)
    If Len(Trim$(assistantText)) = 0 Then
        assistantText = JsonReadString(responseText, "reasoning_content", searchStart)
        If Len(Trim$(assistantText)) > 0 Then
            assistantText = "【模型思考过长，正式回答未生成；以下是其思考过程】" & _
                            vbCrLf & vbCrLf & assistantText
        Else
            errorText = "DeepSeek返回成功，但没有可显示的内容。"
            Exit Function
        End If
    End If


    SendDeepSeekConversationRequest = True
    Exit Function

Failed:
    errorText = Err.Description
    auditText = BuildAuditText(vbNullString, 0, endpoint, _
                               GetElapsedSeconds(startTimer))
End Function

Private Sub SetThinkingStatus(ByVal seconds As Long)
    Dim loadedForm As Object
    On Error Resume Next
    Application.StatusBar = "AI正在思考... 已等待 " & CStr(seconds) & " 秒"
    For Each loadedForm In VBA.UserForms
        If StrComp(loadedForm.Name, "frmDeepSeekChat", vbTextCompare) = 0 Then
            loadedForm.lblStatus.Caption = "AI正在思考... 已等待 " & CStr(seconds) & " 秒"
            Exit For
        End If
    Next loadedForm
    On Error GoTo 0
End Sub

Private Sub ClearThinkingStatus()
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

Private Function NormalizeDeepSeekEndpoint(ByVal endpoint As String) As String
    Dim cleanEndpoint As String

    cleanEndpoint = Trim$(endpoint)
    Do While Right$(cleanEndpoint, 1) = "/"
        cleanEndpoint = Left$(cleanEndpoint, Len(cleanEndpoint) - 1)
    Loop

    Select Case LCase$(cleanEndpoint)
        Case "https://api.deepseek.com", "https://api.deepseek.com/v1"
            cleanEndpoint = cleanEndpoint & "/chat/completions"
    End Select

    NormalizeDeepSeekEndpoint = cleanEndpoint
End Function

Private Function BuildRequestBody(ByVal modelName As String, _
                                  ByVal promptText As String, _
                                  ByVal connectionTest As Boolean) As String
    Dim systemText As String
    Dim optionText As String

    If connectionTest Then
        systemText = _
            "这是API连接测试。只返回JSON对象：{""status"":""ok""}，不要添加其他内容。"
        optionText = _
            """thinking"":{""type"":""disabled""}," & _
            """stream"":false,""max_tokens"":80"
    Else
        systemText = _
            "你是严谨的Excel公式错误分析专家。只能依据用户提供的公式、错误和值证据判断，" & _
            "不得虚构工作簿内容。必须只返回一个JSON对象，不要Markdown代码块。固定结构：" & _
            "{""conclusion"":""明确结论"",""error_expression"":""具体出错子表达式""," & _
            """evidence"":[""证据1"",""证据2""],""suggestions"":[""建议1""]," & _
            """confidence"":""高/中/低"",""missing_information"":[]}。" & _
            "证据不足时降低置信度并填写missing_information。"
        optionText = _
            """thinking"":{""type"":""enabled""}," & _
            """reasoning_effort"":""high"",""stream"":false," & _
            """max_tokens"":1800"
    End If

    BuildRequestBody = _
        "{""model"":""" & JsonEscape(modelName) & _
        """,""messages"":[{""role"":""system"",""content"":""" & _
        JsonEscape(systemText) & _
        """},{""role"":""user"",""content"":""" & _
        JsonEscape(promptText) & _
        """}],""response_format"":{""type"":""json_object""}," & _
        optionText & "}"
End Function

Private Function BuildAuditText(ByVal responseText As String, _
                                ByVal statusCode As Long, _
                                ByVal endpoint As String, _
                                ByVal elapsedSeconds As Double) As String
    Dim responseId As String, actualModel As String, finishReason As String
    Dim usageStart As Long, detailsStart As Long
    Dim promptTokens As Long, completionTokens As Long, totalTokens As Long
    Dim cacheHitTokens As Long, cacheMissTokens As Long, reasoningTokens As Long
    Dim endpointName As String

    responseId = JsonReadString(responseText, "id", 1)
    actualModel = JsonReadString(responseText, "model", 1)
    finishReason = JsonReadString(responseText, "finish_reason", 1)
    usageStart = InStr(1, responseText, """usage""", vbBinaryCompare)
    If usageStart = 0 Then usageStart = 1
    detailsStart = InStr(usageStart, responseText, """completion_tokens_details""", vbBinaryCompare)
    If detailsStart = 0 Then detailsStart = usageStart

    promptTokens = JsonReadLong(responseText, "prompt_tokens", usageStart)
    completionTokens = JsonReadLong(responseText, "completion_tokens", usageStart)
    totalTokens = JsonReadLong(responseText, "total_tokens", usageStart)
    cacheHitTokens = JsonReadLong(responseText, "prompt_cache_hit_tokens", usageStart)
    cacheMissTokens = JsonReadLong(responseText, "prompt_cache_miss_tokens", usageStart)
    reasoningTokens = JsonReadLong(responseText, "reasoning_tokens", detailsStart)

    If LCase$(Left$(endpoint, Len("https://api.deepseek.com"))) = "https://api.deepseek.com" Then
        endpointName = "DeepSeek官方接口"
    Else
        endpointName = "自定义接口"
    End If

    BuildAuditText = _
        "连接：" & endpointName & vbCrLf & _
        "HTTP：" & IIf(statusCode = 0, "未建立", CStr(statusCode)) & vbCrLf & _
        "响应ID：" & TextOrMissing(responseId) & vbCrLf & _
        "实际模型：" & TextOrMissing(actualModel) & vbCrLf & _
        "输入Token：" & TokenOrMissing(promptTokens) & vbCrLf & _
        "输出Token：" & TokenOrMissing(completionTokens) & vbCrLf & _
        "思考Token：" & TokenOrMissing(reasoningTokens) & vbCrLf & _
        "总Token：" & TokenOrMissing(totalTokens) & vbCrLf & _
        "缓存命中：" & TokenOrMissing(cacheHitTokens) & vbCrLf & _
        "缓存未命中：" & TokenOrMissing(cacheMissTokens) & vbCrLf & _
        "结束原因：" & TextOrMissing(finishReason) & vbCrLf & _
        "耗时：" & Format$(elapsedSeconds, "0.0") & "秒"
End Function

Private Function JsonReadLong(ByVal jsonText As String, _
                              ByVal keyName As String, _
                              ByVal startAt As Long) As Long
    Dim keyToken As String, oneChar As String
    Dim keyPosition As Long, colonPosition As Long
    Dim numberStart As Long, numberEnd As Long

    JsonReadLong = -1
    keyToken = """" & keyName & """"
    keyPosition = InStr(startAt, jsonText, keyToken, vbBinaryCompare)
    If keyPosition = 0 Then Exit Function
    colonPosition = InStr(keyPosition + Len(keyToken), jsonText, ":", vbBinaryCompare)
    If colonPosition = 0 Then Exit Function
    numberStart = colonPosition + 1
    Do While numberStart <= Len(jsonText) And _
             InStr(1, " " & vbTab & vbCr & vbLf, Mid$(jsonText, numberStart, 1), vbBinaryCompare) > 0
        numberStart = numberStart + 1
    Loop
    numberEnd = numberStart
    Do While numberEnd <= Len(jsonText)
        oneChar = Mid$(jsonText, numberEnd, 1)
        If (oneChar < "0" Or oneChar > "9") And oneChar <> "-" Then Exit Do
        numberEnd = numberEnd + 1
    Loop
    If numberEnd <= numberStart Then Exit Function
    On Error Resume Next
    JsonReadLong = CLng(Mid$(jsonText, numberStart, numberEnd - numberStart))
    If Err.Number <> 0 Then JsonReadLong = -1
    On Error GoTo 0
End Function

Private Function JsonReadStringArray(ByVal jsonText As String, _
                                     ByVal keyName As String) As String
    Dim keyToken As String, itemText As String, resultText As String
    Dim keyPosition As Long, colonPosition As Long
    Dim arrayStart As Long, arrayEnd As Long, quotePosition As Long

    keyToken = """" & keyName & """"
    keyPosition = InStr(1, jsonText, keyToken, vbBinaryCompare)
    If keyPosition = 0 Then Exit Function
    colonPosition = InStr(keyPosition + Len(keyToken), jsonText, ":", vbBinaryCompare)
    If colonPosition = 0 Then Exit Function
    arrayStart = InStr(colonPosition + 1, jsonText, "[", vbBinaryCompare)
    If arrayStart = 0 Then Exit Function
    arrayEnd = FindJsonArrayEnd(jsonText, arrayStart)
    If arrayEnd = 0 Then Exit Function

    quotePosition = InStr(arrayStart + 1, jsonText, """", vbBinaryCompare)
    Do While quotePosition > 0 And quotePosition < arrayEnd
        itemText = DecodeJsonString(jsonText, quotePosition)
        If Len(itemText) > 0 Then
            If Len(resultText) > 0 Then resultText = resultText & vbCrLf
            resultText = resultText & "- " & itemText
        End If
        quotePosition = NextJsonStringPosition(jsonText, quotePosition)
        If quotePosition = 0 Or quotePosition >= arrayEnd Then Exit Do
    Loop
    JsonReadStringArray = resultText
End Function

Private Function FindJsonArrayEnd(ByVal jsonText As String, ByVal arrayStart As Long) As Long
    Dim i As Long, level As Long
    Dim insideString As Boolean, escaped As Boolean
    Dim oneChar As String
    For i = arrayStart To Len(jsonText)
        oneChar = Mid$(jsonText, i, 1)
        If insideString Then
            If escaped Then
                escaped = False
            ElseIf oneChar = "\" Then
                escaped = True
            ElseIf oneChar = """" Then
                insideString = False
            End If
        Else
            If oneChar = """" Then
                insideString = True
            ElseIf oneChar = "[" Then
                level = level + 1
            ElseIf oneChar = "]" Then
                level = level - 1
                If level = 0 Then FindJsonArrayEnd = i: Exit Function
            End If
        End If
    Next i
End Function

Private Function NextJsonStringPosition(ByVal jsonText As String, _
                                        ByVal openingQuote As Long) As Long
    Dim i As Long, escaped As Boolean
    Dim oneChar As String
    For i = openingQuote + 1 To Len(jsonText)
        oneChar = Mid$(jsonText, i, 1)
        If escaped Then
            escaped = False
        ElseIf oneChar = "\" Then
            escaped = True
        ElseIf oneChar = """" Then
            NextJsonStringPosition = InStr(i + 1, jsonText, """", vbBinaryCompare)
            Exit Function
        End If
    Next i
End Function

Private Function ExtractJsonObject(ByVal textValue As String) As String
    Dim firstBrace As Long, lastBrace As Long
    firstBrace = InStr(1, textValue, "{", vbBinaryCompare)
    lastBrace = InStrRev(textValue, "}", -1, vbBinaryCompare)
    If firstBrace > 0 And lastBrace >= firstBrace Then
        ExtractJsonObject = Mid$(textValue, firstBrace, lastBrace - firstBrace + 1)
    End If
End Function

Private Function GetElapsedSeconds(ByVal startTimer As Single) As Double
    Dim currentTimer As Single
    currentTimer = Timer
    If currentTimer < startTimer Then currentTimer = currentTimer + 86400
    GetElapsedSeconds = CDbl(currentTimer - startTimer)
End Function

Private Function TokenOrMissing(ByVal tokenCount As Long) As String
    If tokenCount < 0 Then
        TokenOrMissing = "未返回"
    Else
        TokenOrMissing = Format$(tokenCount, "#,##0")
    End If
End Function

Private Function TextOrMissing(ByVal textValue As String) As String
    If Len(textValue) = 0 Then TextOrMissing = "未返回" Else TextOrMissing = textValue
End Function

Private Function JsonEscape(ByVal textValue As String) As String
    Dim i As Long
    Dim charText As String
    Dim charCode As Long
    Dim resultText As String

    For i = 1 To Len(textValue)
        charText = Mid$(textValue, i, 1)
        charCode = AscW(charText)
        Select Case charCode
            Case 34
                resultText = resultText & "\"""
            Case 92
                resultText = resultText & "\\"
            Case 8
                resultText = resultText & "\b"
            Case 9
                resultText = resultText & "\t"
            Case 10
                resultText = resultText & "\n"
            Case 12
                resultText = resultText & "\f"
            Case 13
                resultText = resultText & "\r"
            Case 0 To 31
                resultText = resultText & "\u" & Right$("0000" & Hex$(charCode), 4)
            Case Else
                resultText = resultText & charText
        End Select
    Next i

    JsonEscape = resultText
End Function

Private Function JsonReadString(ByVal jsonText As String, _
                                ByVal keyName As String, _
                                ByVal startAt As Long) As String
    Dim keyToken As String
    Dim keyPosition As Long
    Dim colonPosition As Long
    Dim quotePosition As Long

    keyToken = """" & keyName & """"
    keyPosition = InStr(startAt, jsonText, keyToken, vbBinaryCompare)
    If keyPosition = 0 Then Exit Function

    colonPosition = InStr(keyPosition + Len(keyToken), jsonText, ":", vbBinaryCompare)
    If colonPosition = 0 Then Exit Function

    quotePosition = colonPosition + 1
    Do While quotePosition <= Len(jsonText) And _
             InStr(1, " " & vbTab & vbCr & vbLf, _
                   Mid$(jsonText, quotePosition, 1), vbBinaryCompare) > 0
        quotePosition = quotePosition + 1
    Loop

    If Mid$(jsonText, quotePosition, 1) <> """" Then Exit Function
    JsonReadString = DecodeJsonString(jsonText, quotePosition)
End Function

Private Function DecodeJsonString(ByVal jsonText As String, _
                                  ByVal openingQuote As Long) As String
    Dim i As Long
    Dim charText As String
    Dim escapeText As String
    Dim hexText As String
    Dim unicodeValue As Long
    Dim resultText As String

    i = openingQuote + 1
    Do While i <= Len(jsonText)
        charText = Mid$(jsonText, i, 1)
        If charText = """" Then Exit Do

        If charText = "\" Then
            i = i + 1
            If i > Len(jsonText) Then Exit Do
            escapeText = Mid$(jsonText, i, 1)
            Select Case escapeText
                Case """", "\", "/"
                    resultText = resultText & escapeText
                Case "b"
                    resultText = resultText & Chr$(8)
                Case "f"
                    resultText = resultText & Chr$(12)
                Case "n"
                    resultText = resultText & vbLf
                Case "r"
                    resultText = resultText & vbCr
                Case "t"
                    resultText = resultText & vbTab
                Case "u"
                    If i + 4 <= Len(jsonText) Then
                        hexText = Mid$(jsonText, i + 1, 4)
                        On Error Resume Next
                        unicodeValue = CLng("&H" & hexText)
                        If Err.Number = 0 Then
                            resultText = resultText & UnicodeCharacter(unicodeValue)
                        End If
                        Err.Clear
                        On Error GoTo 0
                        i = i + 4
                    End If
                Case Else
                    resultText = resultText & escapeText
            End Select
        Else
            resultText = resultText & charText
        End If
        i = i + 1
    Loop

    DecodeJsonString = resultText
End Function

Private Function UnicodeCharacter(ByVal unicodeValue As Long) As String
    If unicodeValue > 32767 Then unicodeValue = unicodeValue - 65536
    UnicodeCharacter = ChrW(unicodeValue)
End Function

Private Function Utf8BytesToString(ByVal byteData As Variant) As String
    Dim stream As Object

    On Error GoTo Failed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write byteData
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    Utf8BytesToString = stream.ReadText
    stream.Close
    Exit Function

Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Utf8BytesToString = vbNullString
End Function

Private Function StringToUtf8Bytes(ByVal textValue As String) As Variant
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textValue
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3
    StringToUtf8Bytes = stream.Read
    stream.Close
End Function

Private Function ProtectSecret(ByVal plainText As String) As String
    Dim inputBlob As DATA_BLOB
    Dim outputBlob As DATA_BLOB
    Dim protectedBytes() As Byte

    On Error GoTo CleanUp
    If Len(plainText) = 0 Then Exit Function

    inputBlob.cbData = LenB(plainText)
    inputBlob.pbData = StrPtr(plainText)

    If CryptProtectData(inputBlob, 0, 0, 0, 0, _
                        CRYPTPROTECT_UI_FORBIDDEN, outputBlob) = 0 Then
        Exit Function
    End If

    If outputBlob.cbData <= 0 Then Exit Function
    ReDim protectedBytes(0 To outputBlob.cbData - 1)
    CopyMemory protectedBytes(0), ByVal outputBlob.pbData, outputBlob.cbData
    ProtectSecret = Base64Encode(protectedBytes)

CleanUp:
    If outputBlob.pbData <> 0 Then Call LocalFree(outputBlob.pbData)
End Function

Private Function UnprotectSecret(ByVal protectedText As String) As String
    Dim shellObject As Object
    Dim processObject As Object
    Dim commandText As String
    Dim encodedCommand As String
    Dim startTime As Date

    On Error GoTo Failed
    If Len(protectedText) = 0 Then Exit Function

    commandText = _
        "Add-Type -AssemblyName System.Security;" & _
        "$b=[Convert]::FromBase64String('" & protectedText & "');" & _
        "$p=[Security.Cryptography.ProtectedData]::Unprotect(" & _
        "$b,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser);" & _
        "[Console]::Write([Text.Encoding]::Unicode.GetString($p))"
    encodedCommand = PowerShellEncodedCommand(commandText)
    If Len(encodedCommand) = 0 Then Exit Function

    Set shellObject = CreateObject("WScript.Shell")
    Set processObject = shellObject.Exec( _
        "powershell.exe -NoLogo -NoProfile -NonInteractive " & _
        "-WindowStyle Hidden -EncodedCommand " & encodedCommand)

    startTime = Now
    Do While processObject.Status = 0
        DoEvents
        If DateDiff("s", startTime, Now) >= 15 Then
            processObject.Terminate
            Exit Function
        End If
    Loop

    If processObject.ExitCode = 0 Then
        UnprotectSecret = CStr(processObject.StdOut.ReadAll)
    End If
    Exit Function

Failed:
    UnprotectSecret = vbNullString
End Function

Private Function PowerShellEncodedCommand(ByVal commandText As String) As String
    Dim stream As Object
    Dim byteData As Variant

    On Error GoTo Failed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "unicode"
    stream.Open
    stream.WriteText commandText
    stream.Position = 0
    stream.Type = 1
    stream.Position = 2
    byteData = stream.Read
    stream.Close
    PowerShellEncodedCommand = Base64EncodeVariant(byteData)
    Exit Function

Failed:
    PowerShellEncodedCommand = vbNullString
End Function

Private Function Base64EncodeVariant(ByVal byteData As Variant) As String
    Dim xmlDocument As Object
    Dim xmlNode As Object

    Set xmlDocument = CreateObject("MSXML2.DOMDocument.6.0")
    Set xmlNode = xmlDocument.createElement("base64")
    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = byteData
    Base64EncodeVariant = Replace(Replace(xmlNode.text, vbCr, vbNullString), _
                                  vbLf, vbNullString)
End Function

Private Function Base64Encode(ByRef byteData() As Byte) As String
    Dim xmlDocument As Object
    Dim xmlNode As Object

    Set xmlDocument = CreateObject("MSXML2.DOMDocument.6.0")
    Set xmlNode = xmlDocument.createElement("base64")
    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = byteData
    Base64Encode = Replace(Replace(xmlNode.text, vbCr, vbNullString), _
                           vbLf, vbNullString)
End Function

Private Function Base64Decode(ByVal base64Text As String) As Variant
    Dim xmlDocument As Object
    Dim xmlNode As Object

    Set xmlDocument = CreateObject("MSXML2.DOMDocument.6.0")
    Set xmlNode = xmlDocument.createElement("base64")
    xmlNode.DataType = "bin.base64"
    xmlNode.text = base64Text
    Base64Decode = xmlNode.nodeTypedValue
End Function

Private Function LimitText(ByVal textValue As String, _
                           ByVal maximumLength As Long) As String
    If Len(textValue) <= maximumLength Then
        LimitText = textValue
    Else
        LimitText = Left$(textValue, maximumLength - 3) & "..."
    End If
End Function




