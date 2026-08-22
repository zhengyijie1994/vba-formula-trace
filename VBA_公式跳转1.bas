'''''' 模块: VBA/公式跳转1  类型: 公式跳转1.bas  行数: 2733
Attribute VB_Name = "公式跳转1"
Option Explicit

Private Const MAX_REFS As Long = 5000
Private Const MAX_DEPTH As Long = 30
Private Const MAX_PROMPT As Long = 240
Private Const MAX_VALUE_ITEMS As Long = 80
Private Const MAX_VALUE_VIEW_LENGTH As Long = 30000
Private Const VALUE_FULL_SCAN_LIMIT As Double = 20000
Private Const AI_MAX_CONTEXT As Long = 14000
Private Const AI_MAX_RESULT As Long = 30000

Private Const REF_CELLS As Long = 1
Private Const REF_COLUMNS As Long = 2
Private Const REF_ROWS As Long = 3

Private Const CAND_ERROR As Long = 1
Private Const CAND_ZERO_BLANK As Long = 2
Private Const CAND_TEXT As Long = 3
Private Const CAND_ALL As Long = 4
Private Const CAND_BROKEN As Long = 5
Private Const CAND_INVALID As Long = 6

Private Type FormulaRef
    sheetName As String
    addressText As String
    firstRow As Long
    lastRow As Long
    firstColumn As Long
    lastColumn As Long
    refKind As Long
    HasTarget As Boolean
    tokenStart As Long
    tokenLength As Long
    displayText As String
End Type

Private Type TraceCandidate
    refIndex As Long
    candidateKind As Long
    Note As String
    firstCellAddress As String
End Type

Private m_RootCell As Range
Private m_SourceCell As Range
Private m_CurrentDrillCell As Range
Private m_PathCells As Collection
Private m_ResumeIndexes As Collection

Private m_FormulaText As String
Private m_Refs() As FormulaRef
Private m_RefCount As Long
Private m_Candidates() As TraceCandidate
Private m_CandidateCount As Long
Private m_CurrentIndex As Long

Private m_ErrorOnly As Boolean
Private m_FallbackAll As Boolean
Private m_ShowingSource As Boolean
Private m_ShowValues As Boolean
Private m_RootCauseSummary As String
Private m_ShowAI As Boolean
Private m_AIAnalysisText As String
Private m_Diagnosis As String
Private m_Initialized As Boolean

Public Sub TraceFormulaError()
    Call StartTraceFromActive
End Sub

Public Sub TraceUseCurrent(ByVal fromForm As Boolean)
    Call StartTraceFromActive
End Sub

Public Sub TracePrevious(ByVal fromForm As Boolean)
    Dim refIndex As Long
    Dim newIndex As Long

    If Not m_Initialized Then
        Call StartTraceFromActive
        Exit Sub
    End If

    m_RootCauseSummary = vbNullString
    Call ResetAIState

    If m_CandidateCount = 0 Then
        MsgBox "当前公式没有可继续追踪的单元格引用。" & vbCrLf & m_Diagnosis, _
               vbInformation
        Exit Sub
    End If

    newIndex = m_CurrentIndex - 1
    If newIndex < 1 Then newIndex = m_CandidateCount
    refIndex = m_Candidates(newIndex).refIndex

    If SelectCandidate(m_SourceCell, m_Refs(refIndex), m_Candidates(newIndex)) Then
        m_CurrentIndex = newIndex
        m_ShowingSource = False
        Call UpdateTraceWindow
    End If
End Sub

Public Sub TraceNext(ByVal fromForm As Boolean)
    Dim refIndex As Long
    Dim newIndex As Long

    If Not m_Initialized Then
        Call StartTraceFromActive
        Exit Sub
    End If

    m_RootCauseSummary = vbNullString
    Call ResetAIState

    If m_CandidateCount = 0 Then
        MsgBox "当前公式没有可继续追踪的单元格引用。" & vbCrLf & m_Diagnosis, _
               vbInformation
        Exit Sub
    End If

    newIndex = (m_CurrentIndex Mod m_CandidateCount) + 1
    refIndex = m_Candidates(newIndex).refIndex

    If SelectCandidate(m_SourceCell, m_Refs(refIndex), m_Candidates(newIndex)) Then
        m_CurrentIndex = newIndex
        m_ShowingSource = False
        Call UpdateTraceWindow
    End If
End Sub

Public Sub TraceDrill(ByVal fromForm As Boolean)
    Dim savedIndex As Long
    Dim newSource As Range
    Dim loadError As String

    If Not m_Initialized Then Exit Sub
    m_RootCauseSummary = vbNullString
    Call ResetAIState
    Call RefreshDrillCell

    If m_CurrentDrillCell Is Nothing Then
        MsgBox "当前候选中没有可继续深入的错误公式。", vbInformation
        Exit Sub
    End If

    If m_PathCells.count >= MAX_DEPTH Then
        MsgBox "已达到最大追踪深度 " & CStr(MAX_DEPTH) & "。", vbExclamation
        Exit Sub
    End If

    If IsRangeInPath(m_CurrentDrillCell, m_PathCells) Then
        MsgBox "检测到循环引用路径：" & GetCellLabel(m_CurrentDrillCell), vbExclamation
        Exit Sub
    End If

    savedIndex = m_CurrentIndex
    Set newSource = m_CurrentDrillCell
    m_ResumeIndexes.Add savedIndex
    m_PathCells.Add newSource
    Set m_SourceCell = newSource

    If Not LoadCurrentLevel(loadError) Then
        MsgBox loadError, vbExclamation
        m_PathCells.Remove m_PathCells.count
        m_ResumeIndexes.Remove m_ResumeIndexes.count
        Set m_SourceCell = m_PathCells.item(m_PathCells.count)
        Call LoadCurrentLevel(loadError)
        m_CurrentIndex = savedIndex
        m_ShowingSource = (m_CurrentIndex = 0)
    Else
        m_CurrentIndex = 0
        m_ShowingSource = True
        Application.GoTo m_SourceCell, True
    End If

    Call UpdateTraceWindow
End Sub

Public Sub TraceBack(ByVal fromForm As Boolean)
    Dim savedIndex As Long
    Dim refIndex As Long
    Dim loadError As String

    If Not m_Initialized Then Exit Sub
    m_RootCauseSummary = vbNullString
    Call ResetAIState

    If m_PathCells.count <= 1 Then
        MsgBox "当前已经是最初的公式。", vbInformation
        Exit Sub
    End If

    savedIndex = CLng(m_ResumeIndexes.item(m_ResumeIndexes.count))
    m_ResumeIndexes.Remove m_ResumeIndexes.count
    m_PathCells.Remove m_PathCells.count
    Set m_SourceCell = m_PathCells.item(m_PathCells.count)

    If Not LoadCurrentLevel(loadError) Then
        MsgBox loadError, vbExclamation
        Exit Sub
    End If

    If savedIndex > m_CandidateCount Then savedIndex = 0
    m_CurrentIndex = savedIndex
    m_ShowingSource = (m_CurrentIndex = 0)

    If m_CurrentIndex = 0 Then
        Application.GoTo m_SourceCell, True
    Else
        refIndex = m_Candidates(m_CurrentIndex).refIndex
        Call SelectCandidate(m_SourceCell, m_Refs(refIndex), m_Candidates(m_CurrentIndex))
    End If

    Call UpdateTraceWindow
End Sub

Public Sub TraceToggleMode(ByVal fromForm As Boolean)
    Dim loadError As String
    Dim oldErrorOnly As Boolean

    If Not m_Initialized Then Exit Sub
    m_RootCauseSummary = vbNullString
    Call ResetAIState
    oldErrorOnly = m_ErrorOnly
    m_ErrorOnly = Not m_ErrorOnly

    If Not LoadCurrentLevel(loadError) Then
        m_ErrorOnly = oldErrorOnly
        MsgBox loadError, vbExclamation
        Exit Sub
    End If

    m_CurrentIndex = 0
    m_ShowingSource = True
    Application.GoTo m_SourceCell, True
    Call UpdateTraceWindow
End Sub

Public Sub TraceToggleValues(ByVal fromForm As Boolean)
    If Not m_Initialized Then Exit Sub
    If m_ShowAI Then
        m_ShowAI = False
        m_ShowValues = False
    Else
        m_ShowValues = Not m_ShowValues
    End If
    Call UpdateTraceWindow
End Sub

Public Sub TraceAISettings(ByVal fromForm As Boolean)
    If DeepSeekShowSettings() Then
        If m_Initialized Then
            frmFormulaTrace.lblHint.Caption = "DeepSeek设置已保存；点击AI分析或继续对话。"
            frmFormulaTrace.Repaint
        End If
    End If
End Sub

Public Sub TraceAIAnalyze(ByVal fromForm As Boolean)
    Dim promptText As String
    Dim resultText As String
    Dim errorText As String
    Dim auditText As String
    Dim structuredText As String
    Dim structureError As String
    Dim verificationText As String
    Dim verificationError As String
    Dim focusCell As Range

    On Error GoTo Failed
    If Not m_Initialized Then
        If Not StartTraceFromActive() Then Exit Sub
    End If
    If Not DeepSeekEnsureConfigured() Then
        Call AIWorksheetAssistantSetAnalysisBusy(False, "尚未完成AI设置。")
        Exit Sub
    End If

    promptText = BuildDeepSeekPrompt()
    If Len(promptText) = 0 Then
        MsgBox "没有可发送的公式分析信息。", vbInformation
        Exit Sub
    End If

    Call AIWorksheetAssistantSetAnalysisBusy(True, "正在解析当前公式...")
    With frmFormulaTrace
        .cmdAI.enabled = False
        .cmdAI.Caption = "分析中"
        .lblHint.Caption = "正在请求DeepSeek分析当前公式..."
        .Repaint
    End With
    DoEvents

    If Not DeepSeekAnalyzePrompt(promptText, resultText, errorText, auditText) Then
        With frmFormulaTrace
            .cmdAI.enabled = True
            .cmdAI.Caption = "AI分析"
            .lblHint.Caption = "AI分析失败"
            .Repaint
        End With
        MsgBox "DeepSeek分析失败：" & errorText & _
               IIf(Len(auditText) > 0, vbCrLf & vbCrLf & auditText, vbNullString), _
               vbExclamation
        Call AIWorksheetAssistantSetAnalysisBusy(False, "解析失败：" & errorText)
        Exit Sub
    End If

    If Not DeepSeekFormatStructuredAnalysis( _
               resultText, structuredText, structureError) Then
        structuredText = "【结构解析提示】" & vbCrLf & structureError & _
                         vbCrLf & vbCrLf & "【AI原始结果】" & vbCrLf & resultText
    End If

    m_AIAnalysisText = "【DeepSeek AI分析】" & vbCrLf & vbCrLf & _
                       structuredText & vbCrLf & vbCrLf & _
                       "【调用核验】" & vbCrLf & auditText & vbCrLf & vbCrLf & _
                       "提示：AI结论仅供参考，请结合原始数据核对。"
    If Len(m_AIAnalysisText) > AI_MAX_RESULT Then
        m_AIAnalysisText = Left$(m_AIAnalysisText, AI_MAX_RESULT - 16) & _
                           vbCrLf & "……AI结果过长，已截断。"
    End If

    Set focusCell = GetTraceFocusCell()
    m_ShowAI = True
    m_ShowValues = False
    Call UpdateTraceWindow
    Call AIWorksheetAssistantStartFormulaSession( _
             promptText, resultText, m_AIAnalysisText, GetCellLabel(focusCell))
    Call AIWorksheetAssistantSetAnalysisBusy(True, "正在自动检查相关单元格和公式...")
    If AIWorksheetAssistantVerifyFormula(verificationText, verificationError) Then
        If Len(Trim$(verificationText)) > 0 Then
            Call AIWorksheetAssistantAppendVerification(verificationText)
            m_AIAnalysisText = m_AIAnalysisText & vbCrLf & vbCrLf & _
                "【AI自动核验】" & vbCrLf & verificationText
            If Len(m_AIAnalysisText) > AI_MAX_RESULT Then
                m_AIAnalysisText = Left$(m_AIAnalysisText, AI_MAX_RESULT - 16) & _
                                   vbCrLf & "……核验结果过长，已截断。"
            End If
            Call UpdateTraceWindow
        End If
        Call AIWorksheetAssistantSetAnalysisBusy(False, "自动核验完成，可以继续追问。")
    Else
        Call AIWorksheetAssistantAppendVerification( _
            "自动核验未完成：" & verificationError & _
            vbCrLf & "初步分析仍保留，可点击重新解析。")
        Call AIWorksheetAssistantSetAnalysisBusy(False, "初步分析完成；自动核验未完成。")
    End If
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    frmFormulaTrace.cmdAI.enabled = True
    frmFormulaTrace.cmdAI.Caption = "AI分析"
    frmFormulaTrace.lblHint.Caption = "AI分析失败"
    frmFormulaTrace.Repaint
    Call AIWorksheetAssistantSetAnalysisBusy(False, "解析失败：" & errorText)
    On Error GoTo 0
    MsgBox "DeepSeek分析失败：" & errorText, vbExclamation
End Sub

Public Sub TraceAIChat(ByVal fromForm As Boolean)
    Dim promptText As String
    Dim focusCell As Range

    If Not m_Initialized Then
        If Not StartTraceFromActive() Then Exit Sub
    End If
    If DeepSeekConversationReady() Then
        Call OpenWorksheetAIAssistant
    Else
        promptText = BuildDeepSeekPrompt()
        If Len(promptText) = 0 Then
            MsgBox "没有可用于分析的公式上下文。", vbInformation
            Exit Sub
        End If
        Set focusCell = GetTraceFocusCell()
        Call AIWorksheetAssistantPrepareFormulaSession( _
                 promptText, GetCellLabel(focusCell))
    End If
End Sub

Public Sub TraceAIPanelAnalyze(ByVal fromForm As Boolean)
    Call TraceAIAnalyze(fromForm)
End Sub

Public Sub TraceFindRootCause(ByVal fromForm As Boolean)
    Dim loadError As String
    Dim errorText As String
    Dim refIndex As Long
    Dim candidateIndex As Long
    Dim drillCell As Range
    Dim suspectCell As Range
    Dim oldScreenUpdating As Boolean
    Dim oldStatusBar As Variant
    Dim searchStarted As Boolean

    On Error GoTo Failed

    If Not m_Initialized Then
        If Not StartTraceFromActive() Then Exit Sub
    End If

    m_RootCauseSummary = vbNullString
    Call ResetAIState
    m_ErrorOnly = True

    oldScreenUpdating = Application.ScreenUpdating
    oldStatusBar = Application.StatusBar
    searchStarted = True
    Application.ScreenUpdating = False
    Application.StatusBar = "正在自动追踪公式根因..."

    On Error Resume Next
    frmFormulaTrace.cmdRootCause.enabled = False
    frmFormulaTrace.lblHint.Caption = "正在自动追踪根因..."
    frmFormulaTrace.Repaint
    On Error GoTo Failed

    Do
        If Not LoadCurrentLevel(loadError) Then
            Err.Raise vbObjectError + 2201, "TraceFindRootCause", loadError
        End If

        m_CurrentIndex = 0
        m_ShowingSource = True
        Set suspectCell = Nothing
        Set drillCell = Nothing

        If m_CandidateCount = 0 Or m_FallbackAll Then
            m_RootCauseSummary = BuildAutoRootNoCandidateSummary(m_SourceCell)
            Set suspectCell = m_SourceCell
            GoTo RootCauseFound
        End If

        candidateIndex = ChooseAutoRootCandidate( _
                             m_SourceCell.Worksheet.Parent, _
                             m_Refs, m_Candidates, m_CandidateCount, _
                             drillCell)

        If candidateIndex = 0 Then
            m_RootCauseSummary = BuildAutoRootNoCandidateSummary(m_SourceCell)
            Set suspectCell = m_SourceCell
            GoTo RootCauseFound
        End If

        m_CurrentIndex = candidateIndex
        m_ShowingSource = False
        refIndex = m_Candidates(candidateIndex).refIndex

        If drillCell Is Nothing Then
            Set suspectCell = GetAutoRootSuspect( _
                                  m_SourceCell.Worksheet.Parent, _
                                  m_Refs(refIndex), _
                                  m_Candidates(candidateIndex))
            m_RootCauseSummary = BuildAutoRootSummary( _
                                     m_SourceCell, m_Refs(refIndex), _
                                     m_Candidates(candidateIndex), _
                                     suspectCell, vbNullString)
            GoTo RootCauseFound
        End If

        If m_PathCells.count >= MAX_DEPTH Then
            Set suspectCell = drillCell
            m_RootCauseSummary = BuildAutoRootSummary( _
                                     m_SourceCell, m_Refs(refIndex), _
                                     m_Candidates(candidateIndex), _
                                     suspectCell, _
                                     "已达到最大追踪深度 " & CStr(MAX_DEPTH) & _
                                     "；该错误公式是当前最深候选。")
            GoTo RootCauseFound
        End If

        If IsRangeInPath(drillCell, m_PathCells) Then
            Set suspectCell = drillCell
            m_RootCauseSummary = BuildAutoRootSummary( _
                                     m_SourceCell, m_Refs(refIndex), _
                                     m_Candidates(candidateIndex), _
                                     suspectCell, _
                                     "检测到循环引用路径，无法继续深入。")
            GoTo RootCauseFound
        End If

        m_ResumeIndexes.Add m_CurrentIndex
        m_PathCells.Add drillCell
        Set m_SourceCell = drillCell
    Loop

RootCauseFound:
    m_ShowValues = True
    m_Diagnosis = "自动根因候选已定位；详见值透视顶部"

    If Not suspectCell Is Nothing Then
        Application.GoTo suspectCell, True
    ElseIf m_CurrentIndex = 0 Then
        Application.GoTo m_SourceCell, True
    Else
        refIndex = m_Candidates(m_CurrentIndex).refIndex
        Call SelectCandidate(m_SourceCell, m_Refs(refIndex), _
                             m_Candidates(m_CurrentIndex))
    End If

    Call UpdateTraceWindow

CleanExit:
    On Error Resume Next
    If searchStarted Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.StatusBar = oldStatusBar
    End If
    On Error GoTo 0
    Exit Sub

Failed:
    errorText = Err.Description
    m_Diagnosis = "自动追踪失败"
    On Error Resume Next
    If m_Initialized Then Call UpdateTraceWindow
    If searchStarted Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.StatusBar = oldStatusBar
    End If
    On Error GoTo 0
    MsgBox "自动寻找根因失败：" & errorText, vbExclamation
End Sub

Public Sub TraceRoot(ByVal fromForm As Boolean)
    Dim loadError As String

    If Not m_Initialized Then Exit Sub
    m_RootCauseSummary = vbNullString
    Call ResetAIState

    Do While m_PathCells.count > 1
        m_PathCells.Remove m_PathCells.count
    Loop

    Do While m_ResumeIndexes.count > 0
        m_ResumeIndexes.Remove m_ResumeIndexes.count
    Loop

    Set m_SourceCell = m_RootCell
    If Not LoadCurrentLevel(loadError) Then
        MsgBox loadError, vbExclamation
        Exit Sub
    End If

    m_CurrentIndex = 0
    m_ShowingSource = True
    Application.GoTo m_SourceCell, True
    Call UpdateTraceWindow
End Sub

Public Sub TraceRefresh(ByVal fromForm As Boolean)
    If m_Initialized Then Call UpdateTraceWindow
End Sub

Public Sub TraceFormClosed(ByVal fromForm As Boolean)
    Call ResetTraceState
End Sub

Public Sub TraceReturnToFormula(ByVal fromForm As Boolean)
    If Not m_Initialized Then Exit Sub
    If m_SourceCell Is Nothing Then Exit Sub
    On Error GoTo Failed
    Application.GoTo m_SourceCell, True
    Exit Sub
Failed:
    MsgBox "无法返回当前公式单元格：" & Err.Description, vbExclamation
End Sub

Private Function StartTraceFromActive() As Boolean
    Dim activeCell As Range
    Dim loadError As String

    On Error GoTo Failed

    On Error Resume Next
    Set activeCell = Application.activeCell
    On Error GoTo Failed

    If activeCell Is Nothing Then
        MsgBox "请先选择一个公式单元格。", vbInformation
        Exit Function
    End If

    If Not HasFormula(activeCell) Then
        MsgBox "当前单元格没有公式。", vbInformation
        Exit Function
    End If

    Call ResetTraceState
    Set m_RootCell = activeCell
    Set m_SourceCell = activeCell
    Set m_PathCells = New Collection
    Set m_ResumeIndexes = New Collection
    m_PathCells.Add m_SourceCell
    m_ErrorOnly = True
    m_ShowingSource = True

    If Not LoadCurrentLevel(loadError) Then
        MsgBox loadError, vbExclamation
        Exit Function
    End If

    m_Initialized = True
    Application.GoTo m_SourceCell, True

    If Not frmFormulaTrace.Visible Then
        frmFormulaTrace.Show vbModeless
    End If

    Call UpdateTraceWindow
    StartTraceFromActive = True
    Exit Function

Failed:
    MsgBox "启动公式错误追踪失败：" & Err.Description, vbExclamation
End Function

Private Function LoadCurrentLevel(ByRef loadError As String) As Boolean
    LoadCurrentLevel = LoadLevel(m_SourceCell, m_ErrorOnly, m_FormulaText, _
                                 m_Refs, m_RefCount, m_Candidates, _
                                 m_CandidateCount, m_FallbackAll, _
                                 m_Diagnosis, loadError)
End Function

Private Sub RefreshDrillCell()
    Dim refIndex As Long

    Set m_CurrentDrillCell = Nothing
    If Not m_Initialized Then Exit Sub
    If m_ShowingSource Or m_CurrentIndex = 0 Then Exit Sub
    If m_CurrentIndex > m_CandidateCount Then Exit Sub

    refIndex = m_Candidates(m_CurrentIndex).refIndex
    Set m_CurrentDrillCell = FindDrillCell( _
                                m_SourceCell.Worksheet.Parent, _
                                m_Refs(refIndex), m_Candidates(m_CurrentIndex))
End Sub

Private Sub UpdateTraceWindow()
    Dim modeText As String
    Dim currentText As String
    Dim hintText As String
    Dim progressText As String
    Dim refIndex As Long
    Dim selectStart As Long
    Dim selectLength As Long
    Dim formulaDisplayText As String
    Dim traceDisplayText As String

    If Not m_Initialized Then Exit Sub
    Call RefreshDrillCell

    If m_ErrorOnly Then
        If m_FallbackAll Then
            modeText = "错误优先（未发现明显错误，已显示全部）"
        Else
            modeText = "错误优先"
        End If
    Else
        modeText = "全部引用"
    End If

    If m_ShowingSource Or m_CurrentIndex = 0 Then
        progressText = "0 / " & CStr(m_CandidateCount)
        currentText = GetCellLabel(m_SourceCell)
        hintText = m_Diagnosis
        selectStart = 0
        selectLength = 0
    Else
        progressText = CStr(m_CurrentIndex) & " / " & CStr(m_CandidateCount)
        refIndex = m_Candidates(m_CurrentIndex).refIndex
        If m_Refs(refIndex).HasTarget Then
            currentText = m_Refs(refIndex).sheetName & "!" & m_Refs(refIndex).addressText
        Else
            currentText = m_Refs(refIndex).displayText
        End If
        hintText = m_Candidates(m_CurrentIndex).Note
        selectStart = m_Refs(refIndex).tokenStart - 1
        selectLength = m_Refs(refIndex).tokenLength
    End If

    If m_ShowAI Then
        formulaDisplayText = m_AIAnalysisText
        selectStart = 0
        selectLength = 0
    ElseIf m_ShowValues Then
        formulaDisplayText = BuildValuePerspective( _
                                 m_SourceCell, m_Refs, m_RefCount, _
                                 m_Candidates, m_CandidateCount, _
                                 m_CurrentIndex, m_ShowingSource, _
                                 selectStart, selectLength)
    Else
        formulaDisplayText = m_FormulaText
    End If

    If Not m_ShowAI Then
        traceDisplayText = WrapTraceDisplayText(formulaDisplayText, selectStart, selectLength)
        formulaDisplayText = traceDisplayText
    End If

    With frmFormulaTrace
        .lblStatus.Caption = GetErrorText(m_SourceCell) & "    " & modeText & _
                             "    深度 " & CStr(m_PathCells.count)
        .lblPath.Caption = "路径：" & PathText(m_PathCells)
        .lblCurrent.Caption = "候选 " & progressText & "    " & currentText
        .lblHint.Caption = hintText
        .txtFormula.value = formulaDisplayText
        .cmdPrevious.enabled = (m_CandidateCount > 0)
        .cmdNext.enabled = (m_CandidateCount > 0)
        .cmdDrill.enabled = Not (m_CurrentDrillCell Is Nothing)
        .cmdBack.enabled = (m_PathCells.count > 1)
        .cmdRoot.enabled = (m_PathCells.count > 1 Or Not SameCell(m_SourceCell, m_RootCell))
        .cmdValues.enabled = True
        .cmdReturnFormula.enabled = Not (m_SourceCell Is Nothing)
        .cmdRootCause.enabled = True
        .cmdAI.enabled = True
        .cmdAISettings.enabled = True
        If m_ErrorOnly Then
            .cmdMode.Caption = "查看全部"
        Else
            .cmdMode.Caption = "错误优先"
        End If
        If m_ShowAI Then
            .lblFormulaCaption.Caption = "AI分析结果"
            .cmdValues.Caption = "原公式"
            .cmdAI.Caption = "继续对话"
        ElseIf m_ShowValues Then
            .lblFormulaCaption.Caption = "值透视"
            .cmdValues.Caption = "原公式"
            If DeepSeekConversationReady() Then
                .cmdAI.Caption = "AI对话"
            Else
                .cmdAI.Caption = "AI分析"
            End If
        Else
            .lblFormulaCaption.Caption = "原公式"
            .cmdValues.Caption = "值透视"
            If DeepSeekConversationReady() Then
                .cmdAI.Caption = "AI对话"
            Else
                .cmdAI.Caption = "AI分析"
            End If
        End If
        .Repaint
    End With

    On Error Resume Next
    frmFormulaTrace.txtFormula.SetFocus
    frmFormulaTrace.txtFormula.SelStart = selectStart
    frmFormulaTrace.txtFormula.SelLength = selectLength
    On Error GoTo 0
End Sub

Private Sub ResetTraceState()
    Set m_RootCell = Nothing
    Set m_SourceCell = Nothing
    Set m_CurrentDrillCell = Nothing
    Set m_PathCells = Nothing
    Set m_ResumeIndexes = Nothing
    Erase m_Refs
    Erase m_Candidates
    m_FormulaText = vbNullString
    m_RefCount = 0
    m_CandidateCount = 0
    m_CurrentIndex = 0
    m_ErrorOnly = True
    m_FallbackAll = False
    m_ShowingSource = True
    m_ShowValues = False
    m_RootCauseSummary = vbNullString
    Call ResetAIState
    m_Diagnosis = vbNullString
    m_Initialized = False
End Sub

Private Sub ResetAIState()
    m_ShowAI = False
    m_AIAnalysisText = vbNullString
    ' Formula navigation must not discard the separate AI conversation.
End Sub

Private Function LoadLevel(ByVal sourceCell As Range, ByVal errorOnly As Boolean, _
                           ByRef formulaText As String, ByRef refs() As FormulaRef, _
                           ByRef refCount As Long, ByRef candidates() As TraceCandidate, _
                           ByRef candidateCount As Long, ByRef fallbackAll As Boolean, _
                           ByRef diagnosis As String, ByRef loadError As String) As Boolean
    Dim parsedRefs() As FormulaRef
    Dim parsedCandidates() As TraceCandidate
    Dim parsedFormulaText As String
    Dim parsedDiagnosis As String
    Dim parseError As String
    Dim parsedRefCount As Long
    Dim parsedCandidateCount As Long
    Dim parsedFallbackAll As Boolean
    Dim i As Long

    On Error GoTo Failed

    loadError = vbNullString

    If Not HasFormula(sourceCell) Then
        loadError = "当前单元格不再包含公式。"
        Exit Function
    End If

    parsedFormulaText = SafeCellFormula(sourceCell)
    If Len(parsedFormulaText) = 0 Then
        loadError = "无法读取当前单元格公式。"
        Exit Function
    End If
    parsedDiagnosis = GetDiagnosis(sourceCell, parsedFormulaText)

    parsedRefCount = ParseFormulaRefs( _
                         parsedFormulaText, sourceCell.Worksheet, _
                         parsedRefs, parseError)
    If parsedRefCount < 0 Then
        loadError = parseError
        Exit Function
    End If

    parsedCandidateCount = BuildCandidates( _
                               sourceCell, parsedRefs, parsedRefCount, errorOnly, _
                               parsedCandidates, parsedFallbackAll)

    Erase refs
    If parsedRefCount > 0 Then
        ReDim refs(1 To parsedRefCount)
        For i = 1 To parsedRefCount
            With refs(i)
                .sheetName = parsedRefs(i).sheetName
                .addressText = parsedRefs(i).addressText
                .firstRow = parsedRefs(i).firstRow
                .lastRow = parsedRefs(i).lastRow
                .firstColumn = parsedRefs(i).firstColumn
                .lastColumn = parsedRefs(i).lastColumn
                .refKind = parsedRefs(i).refKind
                .HasTarget = parsedRefs(i).HasTarget
                .tokenStart = parsedRefs(i).tokenStart
                .tokenLength = parsedRefs(i).tokenLength
                .displayText = parsedRefs(i).displayText
            End With
        Next i
    End If

    Erase candidates
    If parsedCandidateCount > 0 Then
        ReDim candidates(1 To parsedCandidateCount)
        For i = 1 To parsedCandidateCount
            candidates(i).refIndex = parsedCandidates(i).refIndex
            candidates(i).candidateKind = parsedCandidates(i).candidateKind
            candidates(i).Note = parsedCandidates(i).Note
            candidates(i).firstCellAddress = parsedCandidates(i).firstCellAddress
        Next i
    End If

    formulaText = parsedFormulaText
    refCount = parsedRefCount
    candidateCount = parsedCandidateCount
    fallbackAll = parsedFallbackAll
    diagnosis = parsedDiagnosis
    If Len(parseError) > 0 Then diagnosis = diagnosis & "；" & parseError
    LoadLevel = True
    Exit Function

Failed:
    loadError = "加载追踪信息失败：" & Err.Description
End Function

Private Function BuildCandidates(ByVal sourceCell As Range, ByRef refs() As FormulaRef, _
                                 ByVal refCount As Long, ByVal errorOnly As Boolean, _
                                 ByRef result() As TraceCandidate, _
                                 ByRef fallbackAll As Boolean) As Long
    Dim added() As Boolean
    Dim wb As Workbook
    Dim target As Range
    Dim scanTarget As Range
    Dim suspect As Range
    Dim errorName As String
    Dim noteText As String
    Dim scanNote As String
    Dim i As Long
    Dim itemCount As Long
    Dim capacity As Long

    On Error GoTo Failed

    Erase result
    fallbackAll = False
    If refCount = 0 Then Exit Function

    Set wb = sourceCell.Worksheet.Parent
    ReDim added(1 To refCount)
    errorName = UCase$(GetErrorText(sourceCell))

    If errorOnly Then
        For i = 1 To refCount
            If Not refs(i).HasTarget Then
                AddCandidate result, itemCount, capacity, i, CAND_BROKEN, "公式中存在失效引用"
                added(i) = True
            Else
                Set target = Nothing
                Set suspect = Nothing
                If Not ResolveRef(wb, refs(i), target) Then
                    AddCandidate result, itemCount, capacity, i, CAND_INVALID, "工作表或范围不可用"
                    added(i) = True
                Else
                    Set scanTarget = Nothing
                    scanNote = vbNullString
                    Call PrepareCandidateScanRange(target, scanTarget, scanNote)
                    If Not scanTarget Is Nothing Then Set suspect = GetErrorCells(scanTarget)
                    If Not suspect Is Nothing Then
                        noteText = ErrorRangeNote(suspect)
                        If Len(scanNote) > 0 Then noteText = noteText & "（" & scanNote & "）"
                        AddCandidate result, itemCount, capacity, i, CAND_ERROR, noteText, _
                                     suspect.Areas(1).Cells(1, 1).Address(False, False, xlA1)
                        added(i) = True
                    End If
                End If
            End If
        Next i

        If InStr(1, errorName, "#DIV/0!", vbTextCompare) > 0 Then
            For i = 1 To refCount
                If Not added(i) And refs(i).HasTarget Then
                    Set target = Nothing
                    If ResolveRef(wb, refs(i), target) Then
                        Set scanTarget = Nothing
                        scanNote = vbNullString
                        Call PrepareCandidateScanRange(target, scanTarget, scanNote)
                        If Not scanTarget Is Nothing And ContainsZeroOrBlank(scanTarget) Then
                            noteText = "包含0值或空白，可能导致除零"
                            If Len(scanNote) > 0 Then noteText = noteText & "（" & scanNote & "）"
                            AddCandidate result, itemCount, capacity, i, CAND_ZERO_BLANK, _
                                         noteText
                            added(i) = True
                        End If
                    End If
                End If
            Next i
        ElseIf InStr(1, errorName, "#VALUE!", vbTextCompare) > 0 Then
            For i = 1 To refCount
                If Not added(i) And refs(i).HasTarget Then
                    Set target = Nothing
                    Set scanTarget = Nothing
                    Set suspect = Nothing
                    If ResolveRef(wb, refs(i), target) Then
                        scanNote = vbNullString
                        Call PrepareCandidateScanRange(target, scanTarget, scanNote)
                        If Not scanTarget Is Nothing Then Set suspect = GetTextCells(scanTarget)
                        If Not suspect Is Nothing Then
                            noteText = Format$(RangeCount(suspect), "#,##0") & "个文本值，可能导致#VALUE!"
                            If Len(scanNote) > 0 Then noteText = noteText & "（" & scanNote & "）"
                            AddCandidate result, itemCount, capacity, i, CAND_TEXT, noteText
                            added(i) = True
                        End If
                    End If
                End If
            Next i
        End If

        If itemCount = 0 Then
            fallbackAll = True
            For i = 1 To refCount
                If refs(i).HasTarget Then
                    Set target = Nothing
                    If ResolveRef(wb, refs(i), target) Then
                        AddCandidate result, itemCount, capacity, i, CAND_ALL, "未发现明显上游错误"
                    Else
                        AddCandidate result, itemCount, capacity, i, CAND_INVALID, "工作表或范围不可用"
                    End If
                Else
                    AddCandidate result, itemCount, capacity, i, CAND_BROKEN, "公式中存在失效引用"
                End If
            Next i
        End If
    Else
        For i = 1 To refCount
            If refs(i).HasTarget Then
                Set target = Nothing
                If ResolveRef(wb, refs(i), target) Then
                    AddCandidate result, itemCount, capacity, i, CAND_ALL, "直接引用"
                Else
                    AddCandidate result, itemCount, capacity, i, CAND_INVALID, "工作表或范围不可用"
                End If
            Else
                AddCandidate result, itemCount, capacity, i, CAND_BROKEN, "公式中存在失效引用"
            End If
        Next i
    End If

    If itemCount > 0 Then ReDim Preserve result(1 To itemCount)
    BuildCandidates = itemCount
    Exit Function

Failed:
    Erase result
    BuildCandidates = 0
End Function

Private Sub AddCandidate(ByRef result() As TraceCandidate, ByRef itemCount As Long, _
                         ByRef capacity As Long, ByVal refIndex As Long, _
                         ByVal candidateKind As Long, ByVal noteText As String, _
                         Optional ByVal firstCellAddress As String = vbNullString)
    Dim newCapacity As Long

    itemCount = itemCount + 1
    If itemCount > capacity Then
        If capacity = 0 Then
            capacity = 16
            ReDim result(1 To capacity)
        Else
            newCapacity = capacity * 2
            If newCapacity > MAX_REFS Then newCapacity = MAX_REFS
            capacity = newCapacity
            ReDim Preserve result(1 To capacity)
        End If
    End If

    result(itemCount).refIndex = refIndex
    result(itemCount).candidateKind = candidateKind
    result(itemCount).Note = noteText
    result(itemCount).firstCellAddress = firstCellAddress
End Sub

Private Function ParseFormulaRefs(ByVal formulaText As String, ByVal baseSheet As Worksheet, _
                                  ByRef result() As FormulaRef, _
                                  ByRef parseError As String) As Long
    Dim reg As Object
    Dim matches As Object
    Dim oneMatch As Object
    Dim masked As String
    Dim prefixText As String
    Dim sheetName As String
    Dim addressText As String
    Dim tokenStart As Long
    Dim tokenLength As Long
    Dim firstRow As Long, lastRow As Long
    Dim firstColumn As Long, lastColumn As Long
    Dim refKind As Long
    Dim itemCount As Long
    Dim capacity As Long
    Dim newCapacity As Long

    Const PATTERN As String = _
        "(^|[^A-Za-z0-9_.])" & _
        "((?:'((?:[^']|'')+)'|([^'!+\-*/^&,;()=<>\s:\[\]]+))!)?" & _
        "(" & _
            "#REF![^+\-*/^&,;()=<>\s]*" & _
            "|(?:\$?[A-Za-z]{1,3}\$?\d{1,7})(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?" & _
            "|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}" & _
            "|\$?\d{1,7}:\$?\d{1,7}" & _
        ")" & _
        "(?![A-Za-z0-9_]|\s*\()"

    On Error GoTo Failed

    Erase result
    parseError = vbNullString
    masked = MaskStringLiterals(formulaText)

    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
        .PATTERN = PATTERN
    End With

    If reg.Test(masked) Then
        Set matches = reg.Execute(masked)
        For Each oneMatch In matches
            prefixText = CStr(oneMatch.SubMatches(0))
            tokenStart = CLng(oneMatch.FirstIndex) + Len(prefixText) + 1
            tokenLength = CLng(oneMatch.Length) - Len(prefixText)
            addressText = CStr(oneMatch.SubMatches(4))

            If prefixText = "[" Or prefixText = "!" Then GoTo NextMatch
            If IsInsideBrackets(masked, tokenStart) Then GoTo NextMatch

            If tokenStart + tokenLength <= Len(masked) Then
                If Mid$(masked, tokenStart + tokenLength, 1) = "!" Then
                    Call AppendParseWarning(parseError, _
                             "已跳过暂不支持的3D工作表引用")
                    GoTo NextMatch
                End If
            End If

            If prefixText = "]" Then
                Call AppendParseWarning(parseError, _
                         "已跳过暂不支持的外部工作簿引用")
                GoTo NextMatch
            End If

            If prefixText = ":" Then
                Call AppendParseWarning(parseError, _
                         "已跳过暂不支持的3D工作表引用")
                GoTo NextMatch
            End If

            If CStr(oneMatch.SubMatches(2)) <> vbNullString Then
                sheetName = Replace(CStr(oneMatch.SubMatches(2)), "''", "'")
            ElseIf CStr(oneMatch.SubMatches(3)) <> vbNullString Then
                sheetName = CStr(oneMatch.SubMatches(3))
            Else
                sheetName = baseSheet.Name
            End If

            If InStr(sheetName, "[") > 0 And InStr(sheetName, "]") > 0 Then
                Call AppendParseWarning(parseError, _
                         "已跳过暂不支持的外部工作簿引用")
                GoTo NextMatch
            End If

            If InStr(1, sheetName, ":", vbBinaryCompare) > 0 Then
                Call AppendParseWarning(parseError, _
                         "已跳过暂不支持的3D工作表引用")
                GoTo NextMatch
            End If

            If itemCount >= MAX_REFS Then
                parseError = "公式引用项超过 " & CStr(MAX_REFS) & " 个。"
                ParseFormulaRefs = -1
                Erase result
                Exit Function
            End If

            itemCount = itemCount + 1
            If itemCount > capacity Then
                If capacity = 0 Then
                    capacity = 32
                    ReDim result(1 To capacity)
                Else
                    newCapacity = capacity * 2
                    If newCapacity > MAX_REFS Then newCapacity = MAX_REFS
                    capacity = newCapacity
                    ReDim Preserve result(1 To capacity)
                End If
            End If

            With result(itemCount)
                .sheetName = sheetName
                .addressText = addressText
                .tokenStart = tokenStart
                .tokenLength = tokenLength
                .displayText = Mid$(formulaText, tokenStart, tokenLength)

                If InStr(1, addressText, "#REF!", vbTextCompare) > 0 Then
                    .HasTarget = False
                ElseIf ParseBounds(addressText, firstRow, lastRow, firstColumn, _
                                   lastColumn, refKind) Then
                    .HasTarget = True
                    .firstRow = firstRow
                    .lastRow = lastRow
                    .firstColumn = firstColumn
                    .lastColumn = lastColumn
                    .refKind = refKind
                Else
                    itemCount = itemCount - 1
                End If
            End With
NextMatch:
        Next oneMatch
    End If

    If itemCount = 0 Then
        Erase result
    Else
        ReDim Preserve result(1 To itemCount)
    End If
    ParseFormulaRefs = itemCount
    Exit Function

Failed:
    parseError = "解析公式失败：" & Err.Description
    Erase result
    ParseFormulaRefs = -1
End Function

Private Function ParseStructuredRefs(ByVal formulaText As String, ByVal baseSheet As Worksheet, _
                                     ByRef result() As FormulaRef, _
                                     ByRef parseError As String) As Long
    Dim reg As Object
    Dim matches As Object
    Dim oneMatch As Object
    Dim masked As String
    Dim tableName As String
    Dim columnSpec As String
    Dim itemCount As Long
    Dim capacity As Long
    Dim newCapacity As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim foundWs As Worksheet
    Dim tbl As ListObject
    Dim colName As String
    Dim specItems As Object
    Dim specItem As String
    Dim splitted As Object
    Dim i As Long
    Dim j As Long
    Dim firstRow As Long, lastRow As Long
    Dim firstCol As Long, lastCol As Long
    Dim colIdx As Long
    Dim isCurrentRow As Boolean
    Dim isAllData As Boolean
    Dim isHeaders As Boolean
    Dim isTotals As Boolean
    Dim isThisRow As Boolean
    Dim tokenStart As Long
    Dim tokenLength As Long
    Dim prefixLen As Long

    ' Pattern: [表名] or 表名[列] or 表名[[#数据],[列]]
    ' Also matches bare [@列] and [[#数据],[列]]
    Const PATTERN As String = _
        "(^|[^A-Za-z0-9_.])" & _
        "(" & _
            "'((?:[^']|'')+)'\.?" & _
            "|([A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*)" & _
        ")?" & _
        "\[" & _
        "((?:[@#])?(?:[A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*" & _
            "|\[#[A-Za-z\u4e00-\u9fff]+,[\s]*[A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*\])" & _
        ")\]"

    On Error GoTo Failed

    Erase result
    Set wb = baseSheet.Parent

    masked = MaskStringLiterals(formulaText)

    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
        .PATTERN = PATTERN
    End With

    If reg.Test(masked) Then
        Set matches = reg.Execute(masked)
        For Each oneMatch In matches
            prefixLen = Len(CStr(oneMatch.SubMatches(0)))
            tableName = vbNullString
            If CStr(oneMatch.SubMatches(2)) <> vbNullString Then
                tableName = Replace(CStr(oneMatch.SubMatches(2)), "''", "'")
            ElseIf CStr(oneMatch.SubMatches(3)) <> vbNullString Then
                tableName = CStr(oneMatch.SubMatches(3))
            End If
            columnSpec = CStr(oneMatch.SubMatches(4))
            If Len(columnSpec) = 0 Then GoTo NextStructured

            tokenStart = CLng(oneMatch.FirstIndex) + prefixLen + 1
            tokenLength = CLng(oneMatch.Length) - prefixLen

            If IsInsideBrackets(masked, tokenStart) Then GoTo NextStructured

            ' Find the table
            Set tbl = Nothing
            Set foundWs = Nothing
            If Len(tableName) > 0 Then
                For Each ws In wb.Worksheets
                    For Each tbl In ws.ListObjects
                        If StrComp(tbl.Name, tableName, vbTextCompare) = 0 Then
                            Set foundWs = ws
                            GoTo FoundTable
                        End If
                    Next tbl
                Next ws
            Else
                ' Bare [@col] - use base sheet, check all tables
                For Each ws In wb.Worksheets
                    For Each tbl In ws.ListObjects
                        ' Check if the formula cell is inside this table
                        If Not baseSheet Is Nothing Then
                            On Error Resume Next
                            If Not Application.Intersect(baseSheet.Range(tbl.Name), baseSheet.Cells(1, 1)) Is Nothing Then
                                On Error GoTo 0
                                Set foundWs = ws
                                GoTo FoundTable
                            End If
                            On Error GoTo 0
                        End If
                    Next tbl
                Next ws
                ' Fallback: search all tables
                For Each ws In wb.Worksheets
                    For Each tbl In ws.ListObjects
                        Set foundWs = ws
                        GoTo FoundTable
                    Next tbl
                Next ws
            End If
FoundTable:
            If tbl Is Nothing Then GoTo NextStructured

            ' Parse column spec
            isCurrentRow = False
            isAllData = False
            isHeaders = False
            isTotals = False
            isThisRow = False

            If Left$(columnSpec, 1) = "@" Then
                isCurrentRow = True
                colName = Mid$(columnSpec, 2)
            ElseIf Left$(columnSpec, 1) = "#" Then
                colName = columnSpec
                isThisRow = True
            ElseIf Left$(columnSpec, 1) = "[" Then
                ' [[#数据],[列名]] format
                isAllData = True
                Dim innerSpec As String
                innerSpec = Mid$(columnSpec, 2, Len(columnSpec) - 2)
                ' Split by "],["
                Dim parts() As String
                parts = Split(innerSpec, "],[")
                For i = 0 To UBound(parts)
                    Dim part As String
                    part = parts(i)
                    part = Replace(part, "[", vbNullString)
                    part = Replace(part, "]", vbNullString)
                    If Left$(part, 1) = "#" Then
                        If UCase$(part) = "#ALL" Then
                            isAllData = True: isHeaders = True: isTotals = True
                        ElseIf UCase$(part) = "#DATA" Then
                            isAllData = True
                        ElseIf UCase$(part) = "#HEADERS" Then
                            isHeaders = True
                            isAllData = False
                        ElseIf UCase$(part) = "#TOTALS" Then
                            isTotals = True
                            isAllData = False
                        ElseIf UCase$(part) = "#THIS ROW" Then
                            isThisRow = True
                            isAllData = False
                        End If
                    Else
                        colName = part
                    End If
                Next i
            Else
                colName = columnSpec
                isAllData = True
            End If

            If Len(colName) = 0 Then GoTo NextStructured

            ' Resolve column index
            colIdx = 0
            For j = 1 To tbl.ListColumns.count
                If StrComp(tbl.ListColumns(j).Name, colName, vbTextCompare) = 0 Then
                    colIdx = j
                    Exit For
                End If
            Next j
            If colIdx = 0 Then GoTo NextStructured

            ' Calculate range
            If isAllData Then
                firstRow = tbl.DataBodyRange.Row
                lastRow = firstRow + tbl.DataBodyRange.Rows.count - 1
            ElseIf isHeaders Then
                firstRow = tbl.HeaderRowRange.Row
                lastRow = tbl.HeaderRowRange.Row
            ElseIf isTotals Then
                If tbl.TotalsRowRange Is Nothing Then GoTo NextStructured
                firstRow = tbl.TotalsRowRange.Row
                lastRow = tbl.TotalsRowRange.Row
            ElseIf isThisRow Then
                ' Current row - resolve from the formula cell
                If baseSheet Is Nothing Then GoTo NextStructured
                firstRow = baseSheet.Cells(baseSheet.Cells(1, 1).Row, 1).Row
                ' Actually need the row of the formula cell
                On Error Resume Next
                Dim formulaRow As Long
                formulaRow = 0
                ' We can't get the formula cell row from here easily
                ' Use a heuristic: if [@col] or #THIS ROW, use the active cell row
                Err.Clear
                formulaRow = m_CurrentDrillCell.Row
                If Err.Number <> 0 Or formulaRow = 0 Then
                    formulaRow = baseSheet.Cells(1, 1).Row
                End If
                On Error GoTo 0
                firstRow = formulaRow
                lastRow = formulaRow
            ElseIf isCurrentRow Then
                If baseSheet Is Nothing Then GoTo NextStructured
                On Error Resume Next
                formulaRow = 0
                formulaRow = m_CurrentDrillCell.Row
                If Err.Number <> 0 Or formulaRow = 0 Then
                    formulaRow = baseSheet.Cells(1, 1).Row
                End If
                On Error GoTo 0
                firstRow = formulaRow
                lastRow = formulaRow
            Else
                firstRow = tbl.DataBodyRange.Row
                lastRow = firstRow + tbl.DataBodyRange.Rows.count - 1
            End If

            firstCol = tbl.ListColumns(colIdx).Range.Column
            lastCol = firstCol

            ' Add to result
            If itemCount >= MAX_REFS Then
                parseError = "公式引用项超过 " & CStr(MAX_REFS) & " 个。"
                ParseStructuredRefs = -1
                Erase result
                Exit Function
            End If
            itemCount = itemCount + 1
            If itemCount > capacity Then
                If capacity = 0 Then
                    capacity = 32
                    ReDim result(1 To capacity)
                Else
                    newCapacity = capacity * 2
                    If newCapacity > MAX_REFS Then newCapacity = MAX_REFS
                    capacity = newCapacity
                    ReDim Preserve result(1 To capacity)
                End If
            End If
            With result(itemCount)
                .sheetName = foundWs.Name
                .addressText = ColLetter(firstCol) & CStr(firstRow) & ":" & ColLetter(lastCol) & CStr(lastRow)
                .firstRow = firstRow
                .lastRow = lastRow
                .firstColumn = firstCol
                .lastColumn = lastCol
                .refKind = REF_CELLS
                .HasTarget = True
                .tokenStart = tokenStart
                .tokenLength = tokenLength
                .displayText = Mid$(formulaText, tokenStart, tokenLength)
            End With
NextStructured:
        Next oneMatch
    End If

    If itemCount = 0 Then
        Erase result
    Else
        ReDim Preserve result(1 To itemCount)
    End If
    ParseStructuredRefs = itemCount
    Exit Function

Failed:
    parseError = "解析结构化引用失败：" & Err.Description
    Erase result
    ParseStructuredRefs = -1
End Function

Private Function ColLetter(ByVal columnNumber As Long) As String
    If columnNumber <= 26 Then
        ColLetter = Chr$(64 + columnNumber)
    Else
        ColLetter = Chr$(64 + (columnNumber - 1) \ 26) & Chr$(64 + ((columnNumber - 1) Mod 26) + 1)
    End If
End Function

Private Sub AppendParseWarning(ByRef warningText As String, _
                               ByVal newWarning As String)
    If InStr(1, warningText, newWarning, vbTextCompare) > 0 Then Exit Sub
    If Len(warningText) > 0 Then warningText = warningText & "；"
    warningText = warningText & newWarning
End Sub

Private Function ParseBounds(ByVal addressText As String, ByRef firstRow As Long, _
                             ByRef lastRow As Long, ByRef firstColumn As Long, _
                             ByRef lastColumn As Long, ByRef refKind As Long) As Boolean
    Dim cleanText As String
    Dim colonPos As Long
    Dim leftText As String
    Dim rightText As String
    Dim row1 As Long, row2 As Long
    Dim col1 As Long, col2 As Long

    cleanText = Replace(addressText, "$", vbNullString)
    colonPos = InStr(1, cleanText, ":", vbBinaryCompare)

    If colonPos = 0 Then
        If Not ParseCellAddress(cleanText, col1, row1) Then Exit Function
        firstRow = row1: lastRow = row1
        firstColumn = col1: lastColumn = col1
        refKind = REF_CELLS
        ParseBounds = True
        Exit Function
    End If

    leftText = Left$(cleanText, colonPos - 1)
    rightText = Mid$(cleanText, colonPos + 1)

    If ParseCellAddress(leftText, col1, row1) Then
        If ParseCellAddress(rightText, col2, row2) Then
            NormalizeBounds row1, row2, col1, col2, firstRow, lastRow, firstColumn, lastColumn
            refKind = REF_CELLS
            ParseBounds = True
            Exit Function
        End If
    End If

    If ParseColumn(leftText, col1) Then
        If ParseColumn(rightText, col2) Then
            firstRow = 1: lastRow = 1048576
            If col1 <= col2 Then
                firstColumn = col1: lastColumn = col2
            Else
                firstColumn = col2: lastColumn = col1
            End If
            refKind = REF_COLUMNS
            ParseBounds = True
            Exit Function
        End If
    End If

    If ParseRow(leftText, row1) Then
        If ParseRow(rightText, row2) Then
            firstColumn = 1: lastColumn = 16384
            If row1 <= row2 Then
                firstRow = row1: lastRow = row2
            Else
                firstRow = row2: lastRow = row1
            End If
            refKind = REF_ROWS
            ParseBounds = True
        End If
    End If
End Function

Private Sub NormalizeBounds(ByVal row1 As Long, ByVal row2 As Long, _
                            ByVal col1 As Long, ByVal col2 As Long, _
                            ByRef firstRow As Long, ByRef lastRow As Long, _
                            ByRef firstColumn As Long, ByRef lastColumn As Long)
    If row1 <= row2 Then
        firstRow = row1: lastRow = row2
    Else
        firstRow = row2: lastRow = row1
    End If
    If col1 <= col2 Then
        firstColumn = col1: lastColumn = col2
    Else
        firstColumn = col2: lastColumn = col1
    End If
End Sub

Private Function ParseCellAddress(ByVal textValue As String, ByRef columnNumber As Long, _
                                  ByRef rowNumber As Long) As Boolean
    Dim i As Long
    Dim charCode As Long
    Dim columnText As String

    On Error GoTo Failed
    If Len(textValue) < 2 Then Exit Function

    i = 1
    Do While i <= Len(textValue)
        charCode = AscW(UCase$(Mid$(textValue, i, 1)))
        If charCode < 65 Or charCode > 90 Then Exit Do
        i = i + 1
    Loop
    If i = 1 Or i > Len(textValue) Then Exit Function

    columnText = Left$(textValue, i - 1)
    If Len(columnText) > 3 Then Exit Function
    If Not IsDigits(Mid$(textValue, i)) Then Exit Function

    columnNumber = ColumnToNumber(columnText)
    rowNumber = CLng(Mid$(textValue, i))
    If columnNumber < 1 Or columnNumber > 16384 Then Exit Function
    If rowNumber < 1 Or rowNumber > 1048576 Then Exit Function
    ParseCellAddress = True
    Exit Function
Failed:
    ParseCellAddress = False
End Function

Private Function ParseColumn(ByVal textValue As String, ByRef columnNumber As Long) As Boolean
    Dim i As Long
    Dim charCode As Long
    If Len(textValue) < 1 Or Len(textValue) > 3 Then Exit Function
    For i = 1 To Len(textValue)
        charCode = AscW(UCase$(Mid$(textValue, i, 1)))
        If charCode < 65 Or charCode > 90 Then Exit Function
    Next i
    columnNumber = ColumnToNumber(textValue)
    ParseColumn = (columnNumber >= 1 And columnNumber <= 16384)
End Function

Private Function ParseRow(ByVal textValue As String, ByRef rowNumber As Long) As Boolean
    On Error GoTo Failed
    If Not IsDigits(textValue) Then Exit Function
    rowNumber = CLng(textValue)
    ParseRow = (rowNumber >= 1 And rowNumber <= 1048576)
    Exit Function
Failed:
    ParseRow = False
End Function

Private Function IsDigits(ByVal textValue As String) As Boolean
    Dim i As Long
    Dim charCode As Long
    If Len(textValue) = 0 Then Exit Function
    For i = 1 To Len(textValue)
        charCode = AscW(Mid$(textValue, i, 1))
        If charCode < 48 Or charCode > 57 Then Exit Function
    Next i
    IsDigits = True
End Function

Private Function ColumnToNumber(ByVal columnText As String) As Long
    Dim i As Long
    Dim resultValue As Long
    For i = 1 To Len(columnText)
        resultValue = resultValue * 26 + AscW(UCase$(Mid$(columnText, i, 1))) - 64
    Next i
    ColumnToNumber = resultValue
End Function

Private Function ResolveRef(ByVal wb As Workbook, ByRef item As FormulaRef, _
                            ByRef target As Range) As Boolean
    Dim ws As Worksheet
    On Error GoTo Failed
    Set target = Nothing
    If Not item.HasTarget Then Exit Function
    Set ws = wb.Worksheets(item.sheetName)
    Set target = ws.Range(ws.Cells(item.firstRow, item.firstColumn), _
                          ws.Cells(item.lastRow, item.lastColumn))
    ResolveRef = True
    Exit Function
Failed:
    Set target = Nothing
End Function

Private Sub PrepareCandidateScanRange(ByVal target As Range, _
                                      ByRef scanRange As Range, _
                                      ByRef scopeText As String)
    Dim usedPart As Range

    Set scanRange = Nothing
    scopeText = vbNullString

    If RangeCount(target) <= VALUE_FULL_SCAN_LIMIT Then
        Set scanRange = target
        Exit Sub
    End If

    On Error Resume Next
    Set usedPart = Application.Intersect(target, target.Worksheet.UsedRange)
    On Error GoTo 0

    If usedPart Is Nothing Then
        Set scanRange = target.Cells(1, 1)
        scopeText = "大范围位于已用区域外，按空白判断"
    ElseIf RangeCount(usedPart) > VALUE_FULL_SCAN_LIMIT Then
        Set scanRange = LimitRangeForScan(usedPart, VALUE_FULL_SCAN_LIMIT)
        If scanRange Is Nothing Then Set scanRange = usedPart.Cells(1, 1)
        scopeText = "仅扫描已用交集前 " & _
                    Format$(RangeCount(scanRange), "#,##0") & " 格"
    Else
        Set scanRange = usedPart
        scopeText = "按已用交集 " & _
                    Format$(RangeCount(usedPart), "#,##0") & " 格扫描"
    End If
End Sub

Private Function LimitRangeForScan(ByVal target As Range, _
                                   ByVal maximumCells As Double) As Range
    Dim columnCount As Long
    Dim rowCount As Long
    Dim maximumRows As Long

    On Error GoTo Failed

    columnCount = CLng(target.Columns.CountLarge)
    rowCount = CLng(target.Rows.CountLarge)
    If columnCount < 1 Or rowCount < 1 Then Exit Function

    maximumRows = CLng(Fix(maximumCells / CDbl(columnCount)))
    If maximumRows < 1 Then maximumRows = 1
    If maximumRows > rowCount Then maximumRows = rowCount

    Set LimitRangeForScan = target.Cells(1, 1).Resize(maximumRows, columnCount)
    Exit Function

Failed:
    Set LimitRangeForScan = Nothing
End Function

Private Function SelectCandidate(ByVal sourceCell As Range, ByRef item As FormulaRef, _
                                 ByRef candidate As TraceCandidate) As Boolean
    Dim target As Range
    Dim selectedRange As Range
    Dim errorText As String

    On Error GoTo Failed

    If candidate.candidateKind = CAND_BROKEN Or Not item.HasTarget Then
        Application.GoTo sourceCell, True
        SelectCandidate = True
        Exit Function
    End If

    If Not ResolveRef(sourceCell.Worksheet.Parent, item, target) Then
        Application.GoTo sourceCell, True
        MsgBox "无法定位引用：" & item.displayText, vbExclamation
        Exit Function
    End If

    Select Case candidate.candidateKind
        Case CAND_ERROR
            If Len(candidate.firstCellAddress) > 0 Then
                On Error Resume Next
                Set selectedRange = target.Worksheet.Range(candidate.firstCellAddress)
                On Error GoTo Failed
            End If
            If selectedRange Is Nothing Then Set selectedRange = GetErrorCells(target)
        Case CAND_ZERO_BLANK
            Set selectedRange = FindFirstZeroOrBlankCell(target)
        Case CAND_TEXT
            Set selectedRange = GetTextCells(target)
        Case Else
            Set selectedRange = target
    End Select

    If selectedRange Is Nothing Then Set selectedRange = target
    If candidate.candidateKind <> CAND_ALL Then
        Set selectedRange = selectedRange.Areas(1).Cells(1, 1)
    End If
    Application.GoTo selectedRange, True
    SelectCandidate = True
    Exit Function

Failed:
    errorText = Err.Description
    MsgBox "无法选择引用：" & item.displayText & vbCrLf & errorText, vbExclamation
End Function

Private Function FindDrillCell(ByVal wb As Workbook, ByRef item As FormulaRef, _
                                      ByRef candidate As TraceCandidate) As Range
    Dim target As Range
    Dim formulaErrors As Range
    If Not item.HasTarget Then Exit Function
    If Not ResolveRef(wb, item, target) Then Exit Function
    If Len(candidate.firstCellAddress) > 0 Then
        On Error Resume Next
        Set FindDrillCell = target.Worksheet.Range(candidate.firstCellAddress)
        On Error GoTo 0
        If Not FindDrillCell Is Nothing Then
            If HasFormula(FindDrillCell) Then Exit Function
            Set FindDrillCell = Nothing
        End If
    End If
    Set formulaErrors = GetFormulaErrorCells(target)
    If Not formulaErrors Is Nothing Then
        Set FindDrillCell = formulaErrors.Areas(1).Cells(1, 1)
    End If
End Function

Private Function GetErrorCells(ByVal target As Range) As Range
    Dim formulaErrors As Range
    Dim constantErrors As Range
    Dim resultRange As Range
    On Error Resume Next
    Set formulaErrors = target.SpecialCells(xlCellTypeFormulas, xlErrors)
    Set constantErrors = target.SpecialCells(xlCellTypeConstants, xlErrors)
    If Not formulaErrors Is Nothing Then
        Set formulaErrors = Application.Intersect(target, formulaErrors)
    End If
    If Not constantErrors Is Nothing Then
        Set constantErrors = Application.Intersect(target, constantErrors)
    End If
    On Error GoTo 0
    If Not formulaErrors Is Nothing Then Set resultRange = formulaErrors
    If Not constantErrors Is Nothing Then
        If resultRange Is Nothing Then
            Set resultRange = constantErrors
        Else
            Set resultRange = Application.Union(resultRange, constantErrors)
        End If
    End If
    Set GetErrorCells = resultRange
End Function

Private Function GetFormulaErrorCells(ByVal target As Range) As Range
    Dim foundCells As Range

    On Error Resume Next
    Set foundCells = target.SpecialCells(xlCellTypeFormulas, xlErrors)
    If Not foundCells Is Nothing Then
        Set GetFormulaErrorCells = Application.Intersect(target, foundCells)
    End If
    On Error GoTo 0
End Function

Private Function GetTextCells(ByVal target As Range) As Range
    Dim formulaText As Range
    Dim constantText As Range
    Dim resultRange As Range
    On Error Resume Next
    Set formulaText = target.SpecialCells(xlCellTypeFormulas, xlTextValues)
    Set constantText = target.SpecialCells(xlCellTypeConstants, xlTextValues)
    If Not formulaText Is Nothing Then
        Set formulaText = Application.Intersect(target, formulaText)
    End If
    If Not constantText Is Nothing Then
        Set constantText = Application.Intersect(target, constantText)
    End If
    On Error GoTo 0
    If Not formulaText Is Nothing Then Set resultRange = formulaText
    If Not constantText Is Nothing Then
        If resultRange Is Nothing Then
            Set resultRange = constantText
        Else
            Set resultRange = Application.Union(resultRange, constantText)
        End If
    End If
    Set GetTextCells = resultRange
End Function

Private Function ContainsZeroOrBlank(ByVal target As Range) As Boolean
    Dim data As Variant
    Dim r As Long, c As Long
    Dim val As Variant
    Dim hasBlank As Boolean
    Dim hasZero As Boolean

    If target Is Nothing Then Exit Function

    On Error GoTo Failed

    If RangeCount(target) = 1 Then
        val = target.Value2
        If IsError(val) Then Exit Function
        If IsEmpty(val) Or Len(CStr(val)) = 0 Then
            ContainsZeroOrBlank = True
        ElseIf IsNumeric(val) Then
            ContainsZeroOrBlank = (CDbl(val) = 0)
        End If
        Exit Function
    End If

    data = target.Value2
    If IsArray(data) Then
        For r = LBound(data, 1) To UBound(data, 1)
            For c = LBound(data, 2) To UBound(data, 2)
                If IsEmpty(data(r, c)) Or (VarType(data(r, c)) = vbString And Len(CStr(data(r, c))) = 0) Then
                    hasBlank = True
                ElseIf IsNumeric(data(r, c)) Then
                    If data(r, c) = 0 Then hasZero = True
                End If
                If hasBlank Or hasZero Then
                    ContainsZeroOrBlank = True
                    Exit Function
                End If
            Next c
        Next r
    Else
        val = data
        If IsEmpty(val) Or (VarType(val) = vbString And Len(CStr(val)) = 0) Then
            ContainsZeroOrBlank = True
        ElseIf IsNumeric(val) And val = 0 Then
            ContainsZeroOrBlank = True
        End If
    End If
    Exit Function

Failed:
    ContainsZeroOrBlank = False
End Function

Private Function ErrorRangeNote(ByVal errorCells As Range) As String
    Dim itemCount As Double
    Dim firstCell As Range
    itemCount = RangeCount(errorCells)
    Set firstCell = errorCells.Areas(1).Cells(1, 1)
    If itemCount = 1 Then
        ErrorRangeNote = "错误：" & GetErrorText(firstCell)
    Else
        ErrorRangeNote = Format$(itemCount, "#,##0") & "个错误，首个" & GetErrorText(firstCell)
    End If
End Function

Private Function ChooseAutoRootCandidate(ByVal wb As Workbook, _
                                         ByRef refs() As FormulaRef, _
                                         ByRef candidates() As TraceCandidate, _
                                         ByVal candidateCount As Long, _
                                         ByRef drillCell As Range) As Long
    Dim i As Long
    Dim refIndex As Long
    Dim oneDrillCell As Range
    Dim priorityValue As Long
    Dim bestPriority As Long
    Dim bestIndex As Long

    Set drillCell = Nothing

    For i = 1 To candidateCount
        If candidates(i).candidateKind = CAND_BROKEN Or _
           candidates(i).candidateKind = CAND_INVALID Then
            ChooseAutoRootCandidate = i
            Exit Function
        End If
    Next i

    For i = 1 To candidateCount
        refIndex = candidates(i).refIndex
        Set oneDrillCell = FindDrillCell(wb, refs(refIndex), candidates(i))
        If Not oneDrillCell Is Nothing Then
            Set drillCell = oneDrillCell
            ChooseAutoRootCandidate = i
            Exit Function
        End If
    Next i

    bestPriority = 999
    For i = 1 To candidateCount
        Select Case candidates(i).candidateKind
            Case CAND_ERROR
                priorityValue = 1
            Case CAND_ZERO_BLANK
                priorityValue = 2
            Case CAND_TEXT
                priorityValue = 3
            Case CAND_ALL
                priorityValue = 4
            Case Else
                priorityValue = 10
        End Select

        If priorityValue < bestPriority Then
            bestPriority = priorityValue
            bestIndex = i
        End If
    Next i

    ChooseAutoRootCandidate = bestIndex
End Function

Private Function GetAutoRootSuspect(ByVal wb As Workbook, _
                                    ByRef item As FormulaRef, _
                                    ByRef candidate As TraceCandidate) As Range
    Dim target As Range
    Dim suspect As Range

    If Not item.HasTarget Then Exit Function
    If Not ResolveRef(wb, item, target) Then Exit Function

    Select Case candidate.candidateKind
        Case CAND_ERROR
            Set suspect = GetErrorCells(target)
            If Not suspect Is Nothing Then
                Set GetAutoRootSuspect = suspect.Areas(1).Cells(1, 1)
            End If
        Case CAND_ZERO_BLANK
            Set GetAutoRootSuspect = FindFirstZeroOrBlankCell(target)
        Case CAND_TEXT
            Set suspect = GetTextCells(target)
            If Not suspect Is Nothing Then
                Set GetAutoRootSuspect = suspect.Areas(1).Cells(1, 1)
            End If
        Case Else
            Set GetAutoRootSuspect = target.Cells(1, 1)
    End Select
End Function

Private Function FindFirstZeroOrBlankCell(ByVal target As Range) As Range
    Dim numericCells As Range
    Dim blankCells As Range
    Dim foundCell As Range

    Set numericCells = GetNumericCells(target)
    Set foundCell = FindFirstNumericZero(numericCells)
    If Not foundCell Is Nothing Then
        Set FindFirstZeroOrBlankCell = foundCell
        Exit Function
    End If

    Set blankCells = GetBlankCells(target)
    If Not blankCells Is Nothing Then
        Set FindFirstZeroOrBlankCell = blankCells.Areas(1).Cells(1, 1)
        Exit Function
    End If

    On Error Resume Next
    Set foundCell = target.Find(What:="", LookIn:=xlValues, _
                                LookAt:=xlWhole, SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, MatchCase:=False)
    On Error GoTo 0
    If Not foundCell Is Nothing Then
        Set FindFirstZeroOrBlankCell = foundCell
    Else
        Set FindFirstZeroOrBlankCell = target.Cells(1, 1)
    End If
End Function

Private Function BuildAutoRootSummary(ByVal sourceCell As Range, _
                                      ByRef item As FormulaRef, _
                                      ByRef candidate As TraceCandidate, _
                                      ByVal suspectCell As Range, _
                                      ByVal overrideReason As String) As String
    Dim PathText As String
    Dim targetText As String
    Dim reasonText As String

    PathText = FullTracePath(m_PathCells)

    If Not suspectCell Is Nothing Then
        targetText = GetCellLabel(suspectCell)
        If Not SameCell(suspectCell, sourceCell) Then
            PathText = PathText & " → " & targetText
        End If
    Else
        targetText = item.displayText
        PathText = PathText & " → " & targetText
    End If

    If Len(overrideReason) > 0 Then
        reasonText = overrideReason
    Else
        Select Case candidate.candidateKind
            Case CAND_BROKEN
                reasonText = "公式包含失效引用，Excel无法取得对应单元格。"
            Case CAND_INVALID
                reasonText = "引用的工作表或范围当前不可访问。"
            Case CAND_ERROR
                If suspectCell Is Nothing Then
                    reasonText = "引用范围包含错误值。"
                ElseIf HasFormula(suspectCell) Then
                    reasonText = targetText & " 返回 " & _
                                 ActualCellValueText(suspectCell) & _
                                 "，且没有可继续深入的明显引用。"
                Else
                    reasonText = targetText & " 本身是错误值 " & _
                                 ActualCellValueText(suspectCell) & "。"
                End If
            Case CAND_ZERO_BLANK
                If suspectCell Is Nothing Then
                    reasonText = "引用中包含0值或空白，可能导致除零。"
                ElseIf IsBlankLikeCell(suspectCell) Then
                    reasonText = targetText & " 为空白，可能被当作0参与除法。"
                Else
                    reasonText = targetText & " 的值为 " & _
                                 ActualCellValueText(suspectCell) & _
                                 "，可能导致除零。"
                End If
            Case CAND_TEXT
                If suspectCell Is Nothing Then
                    reasonText = "引用中包含文本，可能参与了数值运算。"
                Else
                    reasonText = targetText & " 是文本值 " & _
                                 ActualCellValueText(suspectCell) & _
                                 "，可能导致#VALUE!。"
                End If
            Case Else
                reasonText = "未发现更深的明确错误，当前引用是最可能的根因候选。"
        End Select
    End If

    BuildAutoRootSummary = "路径：" & PathText & vbCrLf & _
                           "根因候选：" & targetText & vbCrLf & _
                           "判断：" & reasonText
End Function

Private Function BuildAutoRootNoCandidateSummary(ByVal sourceCell As Range) As String
    Dim reasonText As String

    If IsErrorCell(sourceCell) Then
        reasonText = "当前公式返回 " & GetErrorText(sourceCell) & _
                     "，但直接引用中未发现可继续追踪的错误、0/空白或文本；" & _
                     "请检查函数参数、常量、定义名称、条件或匹配方式。"
    Else
        reasonText = "当前公式没有可见错误，也未发现更深的明显异常引用。"
    End If

    BuildAutoRootNoCandidateSummary = _
        "路径：" & FullTracePath(m_PathCells) & vbCrLf & _
        "停留公式：" & GetCellLabel(sourceCell) & vbCrLf & _
        "判断：" & reasonText
End Function

Private Function FullTracePath(ByVal pathCells As Collection) As String
    Dim i As Long
    Dim oneCell As Range
    Dim resultText As String

    For i = 1 To pathCells.count
        Set oneCell = pathCells.item(i)
        If i > 1 Then resultText = resultText & " → "
        resultText = resultText & GetCellLabel(oneCell)
    Next i

    FullTracePath = resultText
End Function

Private Function IsBlankLikeCell(ByVal target As Range) As Boolean
    Dim valueData As Variant

    On Error GoTo Failed
    valueData = target.Value2
    IsBlankLikeCell = IsEmpty(valueData) Or _
                      (VarType(valueData) = vbString And Len(CStr(valueData)) = 0)
    Exit Function

Failed:
    IsBlankLikeCell = False
End Function

Private Function GetTraceFocusCell() As Range
    Dim refIndex As Long
    Dim target As Range
    Dim activeCell As Range

    If Not m_Initialized Or m_SourceCell Is Nothing Then Exit Function
    Set GetTraceFocusCell = m_SourceCell

    If m_ShowingSource Or m_CurrentIndex < 1 Or _
       m_CurrentIndex > m_CandidateCount Then Exit Function

    On Error Resume Next
    Set activeCell = Application.activeCell
    On Error GoTo 0
    If Not activeCell Is Nothing Then
        Set GetTraceFocusCell = activeCell
        Exit Function
    End If

    refIndex = m_Candidates(m_CurrentIndex).refIndex
    If ResolveRef(m_SourceCell.Worksheet.Parent, m_Refs(refIndex), target) Then
        Set GetTraceFocusCell = target.Cells(1, 1)
    End If
End Function

Private Function BuildDeepSeekPrompt() As String
    Dim valueContext As String
    Dim promptText As String
    Dim focusCell As Range
    Dim selectStart As Long
    Dim selectLength As Long

    If Not m_Initialized Or m_SourceCell Is Nothing Then Exit Function

    valueContext = BuildValuePerspective( _
                       m_SourceCell, m_Refs, m_RefCount, _
                       m_Candidates, m_CandidateCount, _
                       m_CurrentIndex, m_ShowingSource, _
                       selectStart, selectLength)

    If Len(valueContext) > AI_MAX_CONTEXT Then
        valueContext = Left$(valueContext, AI_MAX_CONTEXT - 18) & _
                       vbCrLf & "……值证据过长，已截断。"
    End If

    Set focusCell = GetTraceFocusCell()

    promptText = _
        "请分析下面这个Excel公式为什么得到当前结果或错误。" & vbCrLf & _
        "请定位具体出错子表达式和相关单元格；不要只泛泛建议检查引用。" & vbCrLf & vbCrLf & _
        "当前单元格：" & GetCellLabel(focusCell) & vbCrLf & _
        "Source formula cell: " & GetCellLabel(m_SourceCell) & vbCrLf & _
        "当前公式：" & m_FormulaText & vbCrLf & _
        "显示结果：" & SafeCellText(m_SourceCell) & vbCrLf & _
        "实际结果：" & ActualCellValueText(m_SourceCell) & vbCrLf & _
        "错误状态：" & GetErrorText(m_SourceCell) & vbCrLf & _
        "追踪路径：" & FullTracePath(m_PathCells) & vbCrLf & vbCrLf & _
        "以下是VBA在本地整理的引用和值证据：" & vbCrLf & _
        valueContext

    BuildDeepSeekPrompt = promptText
End Function

Private Function BuildValuePerspective(ByVal sourceCell As Range, _
                                       ByRef refs() As FormulaRef, _
                                       ByVal refCount As Long, _
                                       ByRef candidates() As TraceCandidate, _
                                       ByVal candidateCount As Long, _
                                       ByVal currentIndex As Long, _
                                       ByVal showingSource As Boolean, _
                                       ByRef selectStart As Long, _
                                       ByRef selectLength As Long) As String
    Dim resultText As String
    Dim blockText As String
    Dim currentRefIndex As Long
    Dim displayCount As Long
    Dim i As Long

    selectStart = 0
    selectLength = 0

    If Len(m_RootCauseSummary) > 0 Then
        resultText = "【自动根因候选】" & vbCrLf & m_RootCauseSummary & _
                     vbCrLf & vbCrLf
    End If

    resultText = resultText & "【最终结果】" & vbCrLf
    selectStart = Len(resultText)
    selectLength = Len(GetCellLabel(sourceCell))
    resultText = resultText & GetCellLabel(sourceCell) & vbCrLf & _
                 BuildSingleCellValueDetails(sourceCell, "  ") & vbCrLf & vbCrLf & _
                 "【直接引用：" & CStr(refCount) & " 项】" & vbCrLf

    If Not showingSource And currentIndex >= 1 And _
       currentIndex <= candidateCount Then
        currentRefIndex = candidates(currentIndex).refIndex
    End If

    If refCount = 0 Then
        resultText = resultText & _
                     "未解析到直接单元格引用；公式可能使用了名称、结构化引用或常量。"
    Else
        displayCount = refCount
        If displayCount > MAX_VALUE_ITEMS Then displayCount = MAX_VALUE_ITEMS

        For i = 1 To displayCount
            blockText = BuildReferenceValueBlock( _
                            sourceCell.Worksheet.Parent, refs(i), i, _
                            (i = currentRefIndex))
            If i = currentRefIndex Then
                selectStart = Len(resultText)
                selectLength = FirstLineLength(blockText)
            End If
            resultText = resultText & blockText
            If i < displayCount Then resultText = resultText & vbCrLf
        Next i

        If refCount > displayCount Then
            resultText = resultText & vbCrLf & _
                         "……为保证速度，前 " & CStr(displayCount) & _
                         " 项后其余 " & CStr(refCount - displayCount) & " 项已折叠。"

            If currentRefIndex > displayCount Then
                blockText = BuildReferenceValueBlock( _
                                sourceCell.Worksheet.Parent, _
                                refs(currentRefIndex), currentRefIndex, True)
                resultText = resultText & vbCrLf & "当前候选：" & vbCrLf
                selectStart = Len(resultText)
                selectLength = FirstLineLength(blockText)
                resultText = resultText & blockText
            End If
        End If
    End If

    If Len(resultText) > MAX_VALUE_VIEW_LENGTH Then
        resultText = Left$(resultText, MAX_VALUE_VIEW_LENGTH - 18) & _
                     vbCrLf & "……内容过长，已截断。"
        If selectStart + selectLength > Len(resultText) Then
            selectStart = 0
            selectLength = 0
        End If
    End If

    ' Selection positions stay in raw display-text coordinates until the final render.
    BuildValuePerspective = resultText
End Function

Private Function BuildReferenceValueBlock(ByVal wb As Workbook, _
                                          ByRef item As FormulaRef, _
                                          ByVal itemIndex As Long, _
                                          ByVal isCurrent As Boolean) As String
    Dim target As Range
    Dim markerText As String
    Dim headerText As String

    If isCurrent Then
        markerText = ">> "
    Else
        markerText = "  "
    End If

    headerText = markerText & CStr(itemIndex) & ". " & item.displayText

    If Not item.HasTarget Then
        BuildReferenceValueBlock = headerText & vbCrLf & "     类型：失效引用"
        Exit Function
    End If

    If Not ResolveRef(wb, item, target) Then
        BuildReferenceValueBlock = headerText & vbCrLf & "     类型：范围不可用"
        Exit Function
    End If

    If RangeCount(target) = 1 Then
        BuildReferenceValueBlock = headerText & vbCrLf & _
                                   BuildSingleCellValueDetails(target, "     ")
    Else
        BuildReferenceValueBlock = headerText & vbCrLf & _
                                   BuildRangeValueDetails(target, "     ")
    End If
End Function

Private Function BuildSingleCellValueDetails(ByVal target As Range, _
                                             ByVal indentText As String) As String
    Dim displayText As String
    Dim actualText As String
    Dim typeText As String
    Dim formulaText As String
    Dim resultText As String

    displayText = SafeCellText(target)
    actualText = ActualCellValueText(target)
    typeText = CellValueType(target)

    resultText = indentText & "显示：" & displayText & _
                 "｜实际：" & actualText & _
                 "｜类型：" & typeText

    formulaText = SafeCellFormula(target)
    If Len(formulaText) > 0 Then
        resultText = resultText & vbCrLf & indentText & _
                     "自身公式：" & Shorten(CleanOneLine(formulaText), 1000)
    End If

    BuildSingleCellValueDetails = resultText
End Function

Private Function BuildRangeValueDetails(ByVal target As Range, _
                                        ByVal indentText As String) As String
    Dim summaryRange As Range
    Dim errorCells As Range
    Dim textCells As Range
    Dim numericCells As Range
    Dim formulaCells As Range
    Dim blankCells As Range
    Dim firstProblem As Range
    Dim totalCount As Double
    Dim summaryCount As Double
    Dim errorCount As Double
    Dim blankCount As Double
    Dim textCount As Double
    Dim zeroCount As Double
    Dim formulaCount As Double
    Dim scopeText As String
    Dim resultText As String

    totalCount = RangeCount(target)
    PrepareValueSummaryRange target, summaryRange, scopeText

    resultText = indentText & "范围：" & Format$(totalCount, "#,##0") & " 格"
    If Len(scopeText) > 0 Then resultText = resultText & "｜" & scopeText

    If summaryRange Is Nothing Then
        BuildRangeValueDetails = resultText & vbCrLf & indentText & _
                                 "已用区域内没有数据，可视为空白引用。"
        Exit Function
    End If

    summaryCount = RangeCount(summaryRange)
    Set errorCells = GetErrorCells(summaryRange)
    Set textCells = GetTextCells(summaryRange)
    Set numericCells = GetNumericCells(summaryRange)
    Set formulaCells = GetFormulaCells(summaryRange)

    If Not errorCells Is Nothing Then errorCount = RangeCount(errorCells)
    If Not textCells Is Nothing Then textCount = RangeCount(textCells)
    If Not formulaCells Is Nothing Then formulaCount = RangeCount(formulaCells)
    zeroCount = CountNumericZeros(numericCells)

    On Error Resume Next
    blankCount = Application.WorksheetFunction.CountBlank(summaryRange)
    On Error GoTo 0

    resultText = resultText & vbCrLf & indentText & _
                 "统计 " & Format$(summaryCount, "#,##0") & " 格：" & _
                 "错误 " & Format$(errorCount, "#,##0") & _
                 "｜空白 " & Format$(blankCount, "#,##0") & _
                 "｜文本 " & Format$(textCount, "#,##0") & _
                 "｜0值 " & Format$(zeroCount, "#,##0") & _
                 "｜公式 " & Format$(formulaCount, "#,##0")

    If Not errorCells Is Nothing Then
        Set firstProblem = errorCells.Areas(1).Cells(1, 1)
    ElseIf Not textCells Is Nothing Then
        Set firstProblem = textCells.Areas(1).Cells(1, 1)
    ElseIf zeroCount > 0 Then
        Set firstProblem = FindFirstNumericZero(numericCells)
    ElseIf blankCount > 0 Then
        Set blankCells = GetBlankCells(summaryRange)
        If Not blankCells Is Nothing Then
            Set firstProblem = blankCells.Areas(1).Cells(1, 1)
        End If
    End If

    If Not firstProblem Is Nothing Then
        resultText = resultText & vbCrLf & indentText & _
                     "首个异常：" & GetCellLabel(firstProblem) & _
                     " = " & ActualCellValueText(firstProblem)
    End If

    BuildRangeValueDetails = resultText
End Function

Private Sub PrepareValueSummaryRange(ByVal target As Range, _
                                     ByRef summaryRange As Range, _
                                     ByRef scopeText As String)
    Dim usedPart As Range

    Set summaryRange = Nothing
    scopeText = vbNullString

    If RangeCount(target) <= VALUE_FULL_SCAN_LIMIT Then
        Set summaryRange = target
        Exit Sub
    End If

    On Error Resume Next
    Set usedPart = Application.Intersect(target, target.Worksheet.UsedRange)
    On Error GoTo 0

    If usedPart Is Nothing Then
        scopeText = "按已用区域统计"
    ElseIf RangeCount(usedPart) > VALUE_FULL_SCAN_LIMIT Then
        Set summaryRange = LimitRangeForScan(usedPart, VALUE_FULL_SCAN_LIMIT)
        If summaryRange Is Nothing Then Set summaryRange = usedPart.Cells(1, 1)
        scopeText = "已用交集过大，仅统计前 " & _
                    Format$(RangeCount(summaryRange), "#,##0") & " 格"
    Else
        Set summaryRange = usedPart
        scopeText = "按已用交集 " & Format$(RangeCount(usedPart), "#,##0") & " 格统计"
    End If
End Sub

Private Function GetNumericCells(ByVal target As Range) As Range
    Dim formulaNumbers As Range
    Dim constantNumbers As Range
    Dim resultRange As Range

    On Error Resume Next
    Set formulaNumbers = target.SpecialCells(xlCellTypeFormulas, xlNumbers)
    Set constantNumbers = target.SpecialCells(xlCellTypeConstants, xlNumbers)
    If Not formulaNumbers Is Nothing Then
        Set formulaNumbers = Application.Intersect(target, formulaNumbers)
    End If
    If Not constantNumbers Is Nothing Then
        Set constantNumbers = Application.Intersect(target, constantNumbers)
    End If
    On Error GoTo 0

    If Not formulaNumbers Is Nothing Then Set resultRange = formulaNumbers
    If Not constantNumbers Is Nothing Then
        If resultRange Is Nothing Then
            Set resultRange = constantNumbers
        Else
            Set resultRange = Application.Union(resultRange, constantNumbers)
        End If
    End If

    Set GetNumericCells = resultRange
End Function

Private Function GetFormulaCells(ByVal target As Range) As Range
    Dim resultRange As Range

    On Error Resume Next
    Set resultRange = target.SpecialCells(xlCellTypeFormulas)
    If Not resultRange Is Nothing Then
        Set resultRange = Application.Intersect(target, resultRange)
    End If
    On Error GoTo 0

    Set GetFormulaCells = resultRange
End Function

Private Function GetBlankCells(ByVal target As Range) As Range
    Dim resultRange As Range

    On Error Resume Next
    Set resultRange = target.SpecialCells(xlCellTypeBlanks)
    If Not resultRange Is Nothing Then
        Set resultRange = Application.Intersect(target, resultRange)
    End If
    On Error GoTo 0

    Set GetBlankCells = resultRange
End Function

Private Function CountNumericZeros(ByVal numericCells As Range) As Double
    Dim data As Variant
    Dim r As Long, c As Long
    Dim total As Double

    If numericCells Is Nothing Then Exit Function

    On Error GoTo Failed
    data = numericCells.Value2
    If IsArray(data) Then
        For r = LBound(data, 1) To UBound(data, 1)
            For c = LBound(data, 2) To UBound(data, 2)
                If IsNumeric(data(r, c)) Then
                    If data(r, c) = 0 Then total = total + 1
                End If
            Next c
        Next r
    Else
        If IsNumeric(data) Then
            If data = 0 Then total = 1
        End If
    End If
    CountNumericZeros = total
    Exit Function

Failed:
    CountNumericZeros = total
End Function

Private Function FindFirstNumericZero(ByVal numericCells As Range) As Range
    Dim oneArea As Range
    Dim foundCell As Range

    If numericCells Is Nothing Then Exit Function

    On Error Resume Next
    For Each oneArea In numericCells.Areas
        Set foundCell = Nothing
        Set foundCell = oneArea.Find(What:=0, LookIn:=xlValues, _
                                     LookAt:=xlWhole, SearchOrder:=xlByRows, _
                                     SearchDirection:=xlNext, MatchCase:=False)
        If Not foundCell Is Nothing Then
            Set FindFirstNumericZero = foundCell
            Exit Function
        End If
    Next oneArea
    On Error GoTo 0
End Function

Private Function SafeCellText(ByVal target As Range) As String
    Dim textValue As Variant

    On Error GoTo Failed
    textValue = target.text
    If IsNull(textValue) Then
        SafeCellText = "(多值)"
    ElseIf Len(CStr(textValue)) = 0 Then
        SafeCellText = "(空白)"
    Else
        SafeCellText = Shorten(CStr(textValue), 300)
    End If
    Exit Function

Failed:
    SafeCellText = "(无法读取)"
End Function

Private Function ActualCellValueText(ByVal target As Range) As String
    Dim valueData As Variant

    On Error GoTo Failed
    valueData = target.Value2

    If IsError(valueData) Then
        ActualCellValueText = SafeCellText(target)
    ElseIf IsEmpty(valueData) Then
        ActualCellValueText = "(空白)"
    ElseIf VarType(valueData) = vbString Then
        ActualCellValueText = """" & Shorten(CStr(valueData), 300) & """"
    ElseIf VarType(valueData) = vbBoolean Then
        If CBool(valueData) Then
            ActualCellValueText = "TRUE"
        Else
            ActualCellValueText = "FALSE"
        End If
    Else
        ActualCellValueText = Shorten(CStr(valueData), 300)
    End If
    Exit Function

Failed:
    ActualCellValueText = "(无法读取)"
End Function

Private Function CellValueType(ByVal target As Range) As String
    Dim valueData As Variant
    Dim resultText As String

    On Error GoTo Failed
    valueData = target.Value2

    If IsError(valueData) Then
        resultText = "错误"
    Else
        Select Case VarType(valueData)
            Case vbEmpty
                resultText = "空白"
            Case vbString
                If Len(CStr(valueData)) = 0 Then
                    resultText = "空字符串"
                Else
                    resultText = "文本"
                End If
            Case vbBoolean
                resultText = "逻辑值"
            Case vbDate
                resultText = "日期"
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                resultText = "数值"
            Case Else
                resultText = "其他"
        End Select
    End If

    If HasFormula(target) Then resultText = "公式结果·" & resultText
    CellValueType = resultText
    Exit Function

Failed:
    CellValueType = "未知"
End Function

Private Function SafeCellFormula(ByVal target As Range) As String
    Dim formulaValue As Variant

    On Error GoTo Failed
    If Not HasFormula(target) Then Exit Function

    On Error Resume Next
    Err.Clear
    formulaValue = target.Formula2
    If Err.Number <> 0 Then
        Err.Clear
        formulaValue = target.Formula
    End If
    On Error GoTo Failed

    SafeCellFormula = CStr(formulaValue)
    Exit Function

Failed:
    SafeCellFormula = vbNullString
End Function

Private Function ValueTextBoxSelectionStart(ByVal textValue As String, _
                                            ByVal stringPosition As Long) As Long
    Dim prefixText As String
    Dim lineBreakCount As Long

    If stringPosition <= 0 Then Exit Function
    prefixText = Left$(textValue, stringPosition)
    lineBreakCount = Len(prefixText) - _
                     Len(Replace(prefixText, vbCrLf, vbLf))
    ValueTextBoxSelectionStart = stringPosition - lineBreakCount
End Function

Private Function FirstLineLength(ByVal textValue As String) As Long
    Dim lineEnd As Long

    lineEnd = InStr(1, textValue, vbCrLf, vbBinaryCompare)
    If lineEnd = 0 Then
        FirstLineLength = Len(textValue)
    Else
        FirstLineLength = lineEnd - 1
    End If
End Function

Private Function BuildPrompt(ByVal sourceCell As Range, ByVal formulaText As String, _
                             ByRef refs() As FormulaRef, _
                             ByRef candidates() As TraceCandidate, _
                             ByVal candidateCount As Long, ByVal currentIndex As Long, _
                             ByVal showingSource As Boolean, ByVal errorOnly As Boolean, _
                             ByVal fallbackAll As Boolean, ByVal diagnosis As String, _
                             ByVal pathCells As Collection) As String
    Dim modeText As String
    Dim targetText As String
    Dim noteText As String
    Dim contextText As String
    Dim progressText As String
    Dim refIndex As Long
    Dim resultText As String

    If errorOnly Then
        If fallbackAll Then
            modeText = "错误优先→全部"
        Else
            modeText = "错误优先"
        End If
    Else
        modeText = "全部引用"
    End If

    If showingSource Or currentIndex = 0 Then
        progressText = "0/" & CStr(candidateCount)
        targetText = GetCellLabel(sourceCell)
        noteText = diagnosis
        contextText = Shorten(CleanOneLine(formulaText), 65)
    Else
        progressText = CStr(currentIndex) & "/" & CStr(candidateCount)
        refIndex = candidates(currentIndex).refIndex
        If refs(refIndex).HasTarget Then
            targetText = refs(refIndex).sheetName & "!" & refs(refIndex).addressText
        Else
            targetText = refs(refIndex).displayText
        End If
        noteText = candidates(currentIndex).Note
        contextText = formulaContext(formulaText, refs(refIndex))
    End If

    resultText = GetErrorText(sourceCell) & " | " & modeText & _
                 " | 深度" & CStr(pathCells.count) & vbCrLf & _
                 "路径：" & PathText(pathCells) & vbCrLf & _
                 progressText & "：" & ShortenMiddle(targetText, 42) & vbCrLf & _
                 "提示：" & Shorten(noteText, 30) & vbCrLf & _
                 "确定=N；D深入；B返回；A模式；R根；Q退出" & vbCrLf & _
                 "公式：" & contextText

    If Len(resultText) > MAX_PROMPT Then
        resultText = Left$(resultText, MAX_PROMPT - 3) & "..."
    End If
    BuildPrompt = resultText
End Function

Private Function formulaContext(ByVal formulaText As String, ByRef item As FormulaRef) As String
    Dim beforeText As String
    Dim tokenText As String
    Dim afterText As String

    If item.tokenStart < 1 Or item.tokenLength < 1 Then
        formulaContext = Shorten(CleanOneLine(formulaText), 65)
        Exit Function
    End If

    beforeText = CleanOneLine(Left$(formulaText, item.tokenStart - 1))
    tokenText = CleanOneLine(Mid$(formulaText, item.tokenStart, item.tokenLength))
    afterText = CleanOneLine(Mid$(formulaText, item.tokenStart + item.tokenLength))
    If Len(beforeText) > 10 Then beforeText = "..." & Right$(beforeText, 10)
    If Len(afterText) > 10 Then afterText = Left$(afterText, 10) & "..."
    formulaContext = beforeText & "【" & Shorten(tokenText, 40) & "】" & afterText
End Function

Private Function GetDiagnosis(ByVal sourceCell As Range, ByVal formulaText As String) As String
    Dim errorName As String
    Dim upperFormula As String
    errorName = UCase$(GetErrorText(sourceCell))
    upperFormula = UCase$(formulaText)

    If Not IsErrorCell(sourceCell) Then
        If InStr(upperFormula, "IFERROR") > 0 Or InStr(upperFormula, "IFNA") > 0 Then
            GetDiagnosis = "无可见错误；IFERROR可能隐藏错误"
        Else
            GetDiagnosis = "当前结果无错误，检查上游引用"
        End If
    ElseIf InStr(errorName, "#DIV/0!") > 0 Then
        GetDiagnosis = "优先检查0值或空白除数"
    ElseIf InStr(errorName, "#N/A") > 0 Then
        GetDiagnosis = "检查查找值、匹配区域和上游#N/A"
    ElseIf InStr(errorName, "#VALUE!") > 0 Then
        GetDiagnosis = "检查文本运算或数组尺寸不一致"
    ElseIf InStr(errorName, "#REF!") > 0 Then
        GetDiagnosis = "公式存在失效引用，查看【#REF!】"
    ElseIf InStr(errorName, "#NAME?") > 0 Then
        GetDiagnosis = "检查函数名、定义名称和引号"
    ElseIf InStr(errorName, "#NUM!") > 0 Then
        GetDiagnosis = "检查超出允许范围的数值参数"
    ElseIf InStr(errorName, "#SPILL!") > 0 Then
        GetDiagnosis = "检查动态数组输出区域的阻挡单元格"
    Else
        GetDiagnosis = "优先检查含错误的直接引用"
    End If
End Function

Private Function MaskStringLiterals(ByVal formulaText As String) As String
    Dim i As Long
    Dim ch As String
    Dim resultText As String
    Dim insideString As Boolean

    i = 1
    Do While i <= Len(formulaText)
        ch = Mid$(formulaText, i, 1)
        If ch = """" Then
            If insideString And i < Len(formulaText) And _
               Mid$(formulaText, i + 1, 1) = """" Then
                resultText = resultText & "  "
                i = i + 2
            Else
                insideString = Not insideString
                resultText = resultText & " "
                i = i + 1
            End If
        Else
            If insideString Then
                resultText = resultText & " "
            Else
                resultText = resultText & ch
            End If
            i = i + 1
        End If
    Loop
    MaskStringLiterals = resultText
End Function

Private Function IsInsideBrackets(ByVal formulaText As String, _
                                  ByVal oneBasedPosition As Long) As Boolean
    Dim i As Long
    Dim level As Long
    Dim ch As String
    For i = 1 To oneBasedPosition - 1
        ch = Mid$(formulaText, i, 1)
        If ch = "[" Then
            level = level + 1
        ElseIf ch = "]" And level > 0 Then
            level = level - 1
        End If
    Next i
    IsInsideBrackets = (level > 0)
End Function

Private Function PathText(ByVal pathCells As Collection) As String
    Dim rootCell As Range
    Dim currentCell As Range
    Set rootCell = pathCells.item(1)
    If pathCells.count = 1 Then
        PathText = ShortenMiddle(GetCellLabel(rootCell), 38)
    Else
        Set currentCell = pathCells.item(pathCells.count)
        PathText = ShortenMiddle(GetCellLabel(rootCell), 15) & "→" & _
                   IIf(pathCells.count > 2, "...→", vbNullString) & _
                   ShortenMiddle(GetCellLabel(currentCell), 15)
    End If
End Function

Private Function IsRangeInPath(ByVal target As Range, ByVal pathCells As Collection) As Boolean
    Dim i As Long
    Dim oneCell As Range
    For i = 1 To pathCells.count
        Set oneCell = pathCells.item(i)
        If SameCell(target, oneCell) Then
            IsRangeInPath = True
            Exit Function
        End If
    Next i
End Function

Private Function SameCell(ByVal firstCell As Range, ByVal secondCell As Range) As Boolean
    On Error GoTo Failed
    If firstCell.Worksheet Is secondCell.Worksheet Then
        SameCell = (firstCell.Address(False, False, xlA1) = _
                    secondCell.Address(False, False, xlA1))
    End If
    Exit Function
Failed:
    SameCell = False
End Function

Private Function HasFormula(ByVal target As Range) As Boolean
    On Error GoTo Failed
    HasFormula = CBool(target.HasFormula)
    Exit Function
Failed:
    HasFormula = False
End Function

Private Function IsErrorCell(ByVal target As Range) As Boolean
    Dim valueData As Variant
    On Error GoTo Failed
    valueData = target.Value2
    IsErrorCell = IsError(valueData)
    Exit Function
Failed:
    IsErrorCell = False
End Function

Private Function GetErrorText(ByVal target As Range) As String
    On Error GoTo Failed
    If IsErrorCell(target) Then
        GetErrorText = CStr(target.text)
    Else
        GetErrorText = "无可见错误"
    End If
    Exit Function
Failed:
    GetErrorText = "未知错误"
End Function

Private Function GetCellLabel(ByVal target As Range) As String
    On Error GoTo Failed
    GetCellLabel = target.Worksheet.Name & "!" & target.Address(False, False, xlA1)
    Exit Function
Failed:
    GetCellLabel = "未知单元格"
End Function

Private Function RangeCount(ByVal target As Range) As Double
    On Error Resume Next
    RangeCount = CDbl(target.CountLarge)
    If Err.Number <> 0 Then
        Err.Clear
        RangeCount = CDbl(target.count)
    End If
    On Error GoTo 0
End Function

Private Function CleanOneLine(ByVal textValue As String) As String
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    textValue = Replace(textValue, vbTab, " ")
    CleanOneLine = textValue
End Function

Private Function Shorten(ByVal textValue As String, ByVal maximumLength As Long) As String
    If Len(textValue) <= maximumLength Then
        Shorten = textValue
    ElseIf maximumLength <= 3 Then
        Shorten = Left$(textValue, maximumLength)
    Else
        Shorten = Left$(textValue, maximumLength - 3) & "..."
    End If
End Function

Private Function ShortenMiddle(ByVal textValue As String, ByVal maximumLength As Long) As String
    Dim leftLength As Long
    Dim rightLength As Long
    If Len(textValue) <= maximumLength Then
        ShortenMiddle = textValue
        Exit Function
    End If
    If maximumLength <= 3 Then
        ShortenMiddle = Left$(textValue, maximumLength)
        Exit Function
    End If
    leftLength = (maximumLength - 3) \ 2
    rightLength = maximumLength - 3 - leftLength
    ShortenMiddle = Left$(textValue, leftLength) & "..." & Right$(textValue, rightLength)
End Function

Private Function WrapTraceDisplayText(ByVal sourceText As String, _
                                      ByRef selectStart As Long, _
                                      ByRef selectLength As Long) As String
    Const TRACE_DISPLAY_LINE_LENGTH As Long = 96
    Dim resultText As String
    Dim i As Long
    Dim textLength As Long
    Dim lineLength As Long
    Dim ch As String
    Dim originalSelStart As Long
    Dim originalSelEnd As Long
    Dim displaySelStart As Long
    Dim displaySelEnd As Long
    Dim insideString As Boolean

    textLength = Len(sourceText)
    originalSelStart = selectStart
    originalSelEnd = selectStart + selectLength
    displaySelStart = 0
    displaySelEnd = 0

    i = 1
    Do While i <= textLength
        ch = Mid$(sourceText, i, 1)

        ' 记录原始选择起点在显示文本中的位置
        If i - 1 = originalSelStart Then displaySelStart = Len(resultText)

        resultText = resultText & ch

        ' 处理字符串内的连续双引号
        If ch = """" Then
            If insideString And i < textLength And _
               Mid$(sourceText, i + 1, 1) = """" Then
                resultText = resultText & Mid$(sourceText, i + 1, 1)
                i = i + 1
            Else
                insideString = Not insideString
            End If
        End If

        ' 更新行长度计数
        If ch = vbLf Then
            lineLength = 0
        ElseIf ch = vbCr Then
            ' 忽略 CR，行长度不增加（CRLF 中 LF 会重置）
        Else
            lineLength = lineLength + 1
        End If

        ' 记录原始选择终点在显示文本中的位置（在可能的换行插入前）
        If i - 1 = originalSelEnd Then displaySelEnd = Len(resultText)

        ' 超过行长度且在边界字符处才换行
        If ch <> vbCr And ch <> vbLf And i < textLength And _
           lineLength >= TRACE_DISPLAY_LINE_LENGTH Then
            If (Not insideString And IsTraceWrapBoundary(ch)) Or _
               lineLength >= TRACE_DISPLAY_LINE_LENGTH + 20 Then
                If Mid$(sourceText, i + 1, 1) <> vbCr And _
                   Mid$(sourceText, i + 1, 1) <> vbLf Then
                    resultText = resultText & vbCrLf
                    lineLength = 0
                End If
            End If
        End If

        i = i + 1
    Loop

    ' 若原选择范围有效，将显示位置转换为文本框坐标（排除多字节问题）
    If originalSelEnd > originalSelStart Then
        displaySelStart = ValueTextBoxSelectionStart(resultText, displaySelStart)
        displaySelEnd = ValueTextBoxSelectionStart(resultText, displaySelEnd)
        selectStart = displaySelStart
        selectLength = displaySelEnd - displaySelStart
    Else
        selectStart = 0
        selectLength = 0
    End If

    WrapTraceDisplayText = resultText
End Function

Private Function IsTraceWrapBoundary(ByVal ch As String) As Boolean
    IsTraceWrapBoundary = (InStr(1, ",;+-*/^=&()!:", ch, vbBinaryCompare) > 0)
End Function



