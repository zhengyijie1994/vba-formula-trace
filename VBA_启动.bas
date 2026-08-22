'''''' 模块: VBA/启动  类型: 启动.bas  行数: 1366
Attribute VB_Name = "启动"
' =============================================
' RunLauncherBot.bas — WPS 启动机器人 v14 外部启动器版
' =============================================
' 核心改变：
' 1. 不再让 WPS/VBA 直接 ShellExecute 快捷方式
' 2. 不再解析 .lnk 的 TargetPath，避免丢失参数和工作目录
' 3. 改为临时生成外部 .vbs 启动器，由 wscript.exe 模拟系统打开
' 4. 添加启动项时优先手动输入路径，避免 WPS 文件选择器自动解析 .lnk
'
' 支持：
'    - .lnk 快捷方式
'    - .exe / .bat / .cmd / .py / .ps1 / .vbs / .js / 任意文件
'    - 文件夹
'    - URL
'    - cmd: 自定义命令
'
' 示例：
'    C:\Users\hf\Desktop\demo.lnk
'    C:\Tools\app.exe
'    C:\Users\hf\Desktop\project
'    https://chat.openai.com
'    cmd:python C:\bot\main.py
'    cmd:python C:\bot\main.py {doc}
'    cmd:cd /d C:\project && npm run dev
'
' 占位符：
'    {doc}      当前 WPS 文档路径，自动加引号
'    {doc_raw}  当前 WPS 文档路径，不加引号
' =============================================

Const CONFIG_PATH = "C:\Users\hf\Desktop\wps-python-bridge\scripts\scripts_config.json"
Const MAX_SCRIPTS = 100

' =================== 主入口 ===================
Public Sub RunPythonScript()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(CONFIG_PATH) Then
        AddScript
    End If

    If fso.FileExists(CONFIG_PATH) Then
        ShowMainMenu
    End If
End Sub

Public Sub RunLauncherBot()
    RunPythonScript
End Sub

' =================== 文件夹辅助 ===================
Private Sub EnsureFolder(ByVal folderPath)
    On Error Resume Next

    Dim fso
    Dim parentFolder

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(folderPath) = 0 Then Exit Sub
    If fso.FolderExists(folderPath) Then Exit Sub

    parentFolder = fso.GetParentFolderName(folderPath)

    If Len(parentFolder) > 0 Then
        If Not fso.FolderExists(parentFolder) Then
            EnsureFolder parentFolder
        End If
    End If

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

Private Sub EnsureParentFolder(ByVal filePath)
    Dim fso
    Dim folderPath

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = fso.GetParentFolderName(filePath)

    If Len(folderPath) > 0 Then
        EnsureFolder folderPath
    End If
End Sub

' =================== 文件读写 ===================
Private Function ReadUtf8File(ByVal fp)
    On Error GoTo Fallback

    Dim s
    Dim content

    Set s = CreateObject("ADODB.Stream")
    s.Type = 2
    s.Charset = "utf-8"
    s.Open
    s.LoadFromFile fp
    content = s.ReadText(-1)
    s.Close
    Set s = Nothing

    If Len(content) > 0 Then
        If AscW(Left(content, 1)) = &HFEFF Then
            content = Mid(content, 2)
        End If
    End If

    ReadUtf8File = content
    Exit Function

Fallback:
    On Error Resume Next

    Dim fso
    Dim ts

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(fp) Then
        Set ts = fso.OpenTextFile(fp, 1, False)
        content = ts.ReadAll
        ts.Close
    Else
        content = ""
    End If

    If Len(content) > 0 Then
        If AscW(Left(content, 1)) = &HFEFF Then
            content = Mid(content, 2)
        End If
    End If

    ReadUtf8File = content
End Function

Private Sub WriteUtf8File(ByVal fp, ByVal content)
    On Error GoTo Fallback

    EnsureParentFolder fp

    Dim s
    Set s = CreateObject("ADODB.Stream")

    s.Type = 2
    s.Charset = "utf-8"
    s.Open
    s.WriteText content
    s.SaveToFile fp, 2
    s.Close

    Set s = Nothing
    Exit Sub

Fallback:
    On Error Resume Next

    Dim fso
    Dim ts

    EnsureParentFolder fp

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(fp, 2, True)
    ts.Write content
    ts.Close
End Sub

Private Sub WriteUnicodeFile(ByVal fp, ByVal content)
    On Error GoTo Fallback

    EnsureParentFolder fp

    Dim s
    Set s = CreateObject("ADODB.Stream")

    s.Type = 2
    s.Charset = "unicode"
    s.Open
    s.WriteText content
    s.SaveToFile fp, 2
    s.Close

    Set s = Nothing
    Exit Sub

Fallback:
    On Error Resume Next

    Dim fso
    Dim ts

    EnsureParentFolder fp

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(fp, 2, True)
    ts.Write content
    ts.Close
End Sub

' =================== JSON 工具 ===================
Private Function IsHex4(ByVal s)
    Dim i
    Dim ch

    If Len(s) <> 4 Then
        IsHex4 = False
        Exit Function
    End If

    For i = 1 To 4
        ch = Mid(s, i, 1)

        If Not ((ch >= "0" And ch <= "9") Or _
                (ch >= "A" And ch <= "F") Or _
                (ch >= "a" And ch <= "f")) Then
            IsHex4 = False
            Exit Function
        End If
    Next

    IsHex4 = True
End Function

Private Function JsonEsc(ByVal t)
    Dim r
    Dim i
    Dim ch

    r = ""

    For i = 1 To Len(t)
        ch = Mid(t, i, 1)

        If ch = Chr(92) Then
            r = r & Chr(92) & Chr(92)
        ElseIf ch = Chr(34) Then
            r = r & Chr(92) & Chr(34)
        ElseIf ch = vbCr Then
            r = r & Chr(92) & "r"
        ElseIf ch = vbLf Then
            r = r & Chr(92) & "n"
        ElseIf ch = vbTab Then
            r = r & Chr(92) & "t"
        Else
            r = r & ch
        End If
    Next

    JsonEsc = r
End Function

Private Function JsonExtract(ByVal block, ByVal key)
    Dim k
    Dim p
    Dim c
    Dim i
    Dim ch
    Dim nextCh
    Dim hex4
    Dim r

    k = Chr(34) & key & Chr(34)
    p = InStr(1, block, k, vbTextCompare)

    If p = 0 Then
        JsonExtract = ""
        Exit Function
    End If

    c = InStr(p + Len(k), block, ":")

    If c = 0 Then
        JsonExtract = ""
        Exit Function
    End If

    i = c + 1

    Do While i <= Len(block) And Mid(block, i, 1) <= " "
        i = i + 1
    Loop

    If i > Len(block) Then
        JsonExtract = ""
        Exit Function
    End If

    r = ""

    If Mid(block, i, 1) = Chr(34) Then
        i = i + 1

        Do While i <= Len(block)
            ch = Mid(block, i, 1)

            If ch = Chr(92) Then
                If i + 1 <= Len(block) Then
                    nextCh = Mid(block, i + 1, 1)

                    Select Case nextCh
                        Case Chr(34)
                            r = r & Chr(34)
                            i = i + 2

                        Case Chr(92)
                            r = r & Chr(92)
                            i = i + 2

                        Case "/"
                            r = r & "/"
                            i = i + 2

                        Case "b"
                            r = r & Chr(8)
                            i = i + 2

                        Case "f"
                            r = r & Chr(12)
                            i = i + 2

                        Case "n"
                            r = r & vbLf
                            i = i + 2

                        Case "r"
                            r = r & vbCr
                            i = i + 2

                        Case "t"
                            r = r & vbTab
                            i = i + 2

                        Case "u"
                            If i + 5 <= Len(block) Then
                                hex4 = Mid(block, i + 2, 4)

                                If IsHex4(hex4) Then
                                    r = r & ChrW(CLng("&H" & hex4))
                                    i = i + 6
                                Else
                                    r = r & nextCh
                                    i = i + 2
                                End If
                            Else
                                r = r & nextCh
                                i = i + 2
                            End If

                        Case Else
                            r = r & nextCh
                            i = i + 2
                    End Select
                Else
                    i = i + 1
                End If

            ElseIf ch = Chr(34) Then
                Exit Do

            Else
                r = r & ch
                i = i + 1
            End If
        Loop

    Else
        Do While i <= Len(block)
            ch = Mid(block, i, 1)

            If ch = "," Or ch = "}" Or ch = "]" Then
                Exit Do
            End If

            If ch > " " Then
                r = r & ch
            End If

            i = i + 1
        Loop
    End If

    JsonExtract = r
End Function

Private Function FindMatchingChar(ByVal txt, ByVal openPos, ByVal openChar, ByVal closeChar)
    Dim i
    Dim depth
    Dim ch
    Dim inString
    Dim escaped

    depth = 0
    inString = False
    escaped = False

    For i = openPos To Len(txt)
        ch = Mid(txt, i, 1)

        If inString Then
            If escaped Then
                escaped = False
            ElseIf ch = Chr(92) Then
                escaped = True
            ElseIf ch = Chr(34) Then
                inString = False
            End If
        Else
            If ch = Chr(34) Then
                inString = True
            ElseIf ch = openChar Then
                depth = depth + 1
            ElseIf ch = closeChar Then
                depth = depth - 1

                If depth = 0 Then
                    FindMatchingChar = i
                    Exit Function
                End If
            End If
        End If
    Next

    FindMatchingChar = 0
End Function

Private Function FindNextCharOutsideString(ByVal txt, ByVal targetChar, ByVal startPos, ByVal endPos)
    Dim i
    Dim ch
    Dim inString
    Dim escaped

    inString = False
    escaped = False

    If endPos <= 0 Or endPos > Len(txt) Then
        endPos = Len(txt)
    End If

    For i = startPos To endPos
        ch = Mid(txt, i, 1)

        If inString Then
            If escaped Then
                escaped = False
            ElseIf ch = Chr(92) Then
                escaped = True
            ElseIf ch = Chr(34) Then
                inString = False
            End If
        Else
            If ch = Chr(34) Then
                inString = True
            ElseIf ch = targetChar Then
                FindNextCharOutsideString = i
                Exit Function
            End If
        End If
    Next

    FindNextCharOutsideString = 0
End Function

' =================== 配置读写 ===================
Private Function ReadScripts(ids, names, remarks, paths)
    Dim fso
    Dim json
    Dim scriptsKey
    Dim arrStart
    Dim arrEnd
    Dim pos
    Dim objStart
    Dim objEnd
    Dim blk
    Dim idv
    Dim nmv
    Dim rmv
    Dim ptv
    Dim n

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(CONFIG_PATH) Then
        ReadScripts = 0
        Exit Function
    End If

    json = ReadUtf8File(CONFIG_PATH)

    If Len(Trim(json)) = 0 Then
        ReadScripts = 0
        Exit Function
    End If

    scriptsKey = InStr(1, json, Chr(34) & "scripts" & Chr(34), vbTextCompare)

    If scriptsKey = 0 Then
        ReadScripts = 0
        Exit Function
    End If

    arrStart = InStr(scriptsKey, json, "[")

    If arrStart = 0 Then
        ReadScripts = 0
        Exit Function
    End If

    arrEnd = FindMatchingChar(json, arrStart, "[", "]")

    If arrEnd = 0 Or arrEnd <= arrStart Then
        ReadScripts = 0
        Exit Function
    End If

    n = 0
    pos = arrStart + 1

    Do While pos < arrEnd And n < MAX_SCRIPTS
        objStart = FindNextCharOutsideString(json, "{", pos, arrEnd - 1)

        If objStart = 0 Then Exit Do

        objEnd = FindMatchingChar(json, objStart, "{", "}")

        If objEnd = 0 Or objEnd > arrEnd Then Exit Do

        blk = Mid(json, objStart, objEnd - objStart + 1)

        idv = JsonExtract(blk, "id")
        nmv = JsonExtract(blk, "name")
        rmv = JsonExtract(blk, "remark")
        ptv = JsonExtract(blk, "path")

        If IsNumeric(idv) And Len(nmv) > 0 Then
            n = n + 1
            ids(n) = CLng(idv)
            names(n) = nmv
            remarks(n) = rmv
            paths(n) = ptv
        End If

        pos = objEnd + 1
    Loop

    ReadScripts = n
End Function

Private Function GetNextId(ByVal json)
    Dim v

    v = JsonExtract(json, "next_id")

    If IsNumeric(v) Then
        If CLng(v) > 0 Then
            GetNextId = CLng(v)
        Else
            GetNextId = 1
        End If
    Else
        GetNextId = 1
    End If
End Function

Private Function MaxScriptId(ids, ByVal sc)
    Dim i
    Dim m

    m = 0

    For i = 1 To sc
        If ids(i) > m Then
            m = ids(i)
        End If
    Next

    MaxScriptId = m
End Function

Private Sub WriteScriptsConfig(ids, names, remarks, paths, ByVal sc, ByVal nextId)
    Dim json
    Dim i
    Dim q

    q = Chr(34)

    json = "{" & q & "scripts" & q & ":[" & vbCrLf

    For i = 1 To sc
        json = json & "  {" & _
                      q & "id" & q & ":" & ids(i) & "," & _
                      q & "name" & q & ":" & q & JsonEsc(names(i)) & q & "," & _
                      q & "path" & q & ":" & q & JsonEsc(paths(i)) & q & "," & _
                      q & "remark" & q & ":" & q & JsonEsc(remarks(i)) & q & "}"

        If i < sc Then json = json & ","

        json = json & vbCrLf
    Next

    json = json & "]," & q & "next_id" & q & ":" & nextId & "}"

    WriteUtf8File CONFIG_PATH, json
End Sub

Private Sub SaveNewScript(ByVal fp, ByVal fn, ByVal rm)
    Dim ids(MAX_SCRIPTS)
    Dim names(MAX_SCRIPTS)
    Dim remarks(MAX_SCRIPTS)
    Dim paths(MAX_SCRIPTS)

    Dim fso
    Dim json
    Dim sc
    Dim nid
    Dim maxId

    Set fso = CreateObject("Scripting.FileSystemObject")

    sc = ReadScripts(ids, names, remarks, paths)

    If sc >= MAX_SCRIPTS Then
        MsgBox "启动项数量已达到上限：" & MAX_SCRIPTS, vbExclamation
        Exit Sub
    End If

    If fso.FileExists(CONFIG_PATH) Then
        json = ReadUtf8File(CONFIG_PATH)
        nid = GetNextId(json)
    Else
        nid = 1
    End If

    maxId = MaxScriptId(ids, sc)

    If nid <= maxId Then nid = maxId + 1

    sc = sc + 1

    ids(sc) = nid
    names(sc) = fn
    remarks(sc) = rm
    paths(sc) = fp

    WriteScriptsConfig ids, names, remarks, paths, sc, nid + 1
End Sub

Private Sub RemoveById(ByVal targetId)
    Dim ids(MAX_SCRIPTS)
    Dim names(MAX_SCRIPTS)
    Dim remarks(MAX_SCRIPTS)
    Dim paths(MAX_SCRIPTS)

    Dim newIds(MAX_SCRIPTS)
    Dim newNames(MAX_SCRIPTS)
    Dim newRemarks(MAX_SCRIPTS)
    Dim newPaths(MAX_SCRIPTS)

    Dim fso
    Dim json
    Dim nextId
    Dim sc
    Dim i
    Dim n

    Set fso = CreateObject("Scripting.FileSystemObject")

    sc = ReadScripts(ids, names, remarks, paths)

    If fso.FileExists(CONFIG_PATH) Then
        json = ReadUtf8File(CONFIG_PATH)
        nextId = GetNextId(json)
    Else
        nextId = MaxScriptId(ids, sc) + 1
    End If

    n = 0

    For i = 1 To sc
        If CLng(ids(i)) <> CLng(targetId) Then
            n = n + 1
            newIds(n) = ids(i)
            newNames(n) = names(i)
            newRemarks(n) = remarks(i)
            newPaths(n) = paths(i)
        End If
    Next

    If nextId <= MaxScriptId(newIds, n) Then
        nextId = MaxScriptId(newIds, n) + 1
    End If

    WriteScriptsConfig newIds, newNames, newRemarks, newPaths, n, nextId
End Sub

Private Sub UpdateRemark(ByVal targetId, ByVal newRm)
    Dim ids(MAX_SCRIPTS)
    Dim names(MAX_SCRIPTS)
    Dim remarks(MAX_SCRIPTS)
    Dim paths(MAX_SCRIPTS)

    Dim fso
    Dim json
    Dim nextId
    Dim sc
    Dim i

    Set fso = CreateObject("Scripting.FileSystemObject")

    sc = ReadScripts(ids, names, remarks, paths)

    If sc = 0 Then Exit Sub

    For i = 1 To sc
        If CLng(ids(i)) = CLng(targetId) Then
            remarks(i) = newRm
            Exit For
        End If
    Next

    If fso.FileExists(CONFIG_PATH) Then
        json = ReadUtf8File(CONFIG_PATH)
        nextId = GetNextId(json)
    Else
        nextId = MaxScriptId(ids, sc) + 1
    End If

    If nextId <= MaxScriptId(ids, sc) Then
        nextId = MaxScriptId(ids, sc) + 1
    End If

    WriteScriptsConfig ids, names, remarks, paths, sc, nextId
End Sub

' =================== 启动辅助 ===================
Private Function StripOuterQuotes(ByVal s)
    s = Trim(CStr(s))

    If Len(s) >= 2 Then
        If Left(s, 1) = Chr(34) And Right(s, 1) = Chr(34) Then
            s = Mid(s, 2, Len(s) - 2)
        End If
    End If

    StripOuterQuotes = s
End Function

Private Function ExpandEnv(ByVal s)
    Dim wsh

    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    ExpandEnv = wsh.ExpandEnvironmentStrings(CStr(s))

    If Err.Number <> 0 Then
        Err.Clear
        ExpandEnv = CStr(s)
    End If

    On Error GoTo 0
End Function

Private Function QuoteCmdArg(ByVal s)
    QuoteCmdArg = Chr(34) & Replace(CStr(s), Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Private Function VbsQuote(ByVal s)
    VbsQuote = Chr(34) & Replace(CStr(s), Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Private Function IsCmdTarget(ByVal s)
    s = LCase(Trim(CStr(s)))
    IsCmdTarget = False

    If Left(s, 4) = "cmd:" Then
        IsCmdTarget = True
    End If
End Function

Private Function IsUrlTarget(ByVal s)
    s = LCase(Trim(CStr(s)))
    IsUrlTarget = False

    If Left(s, 7) = "http://" Then IsUrlTarget = True
    If Left(s, 8) = "https://" Then IsUrlTarget = True
    If Left(s, 7) = "file://" Then IsUrlTarget = True
    If Left(s, 7) = "mailto:" Then IsUrlTarget = True
End Function

Private Function IsValidLaunchTarget(ByVal s)
    Dim fso
    Dim t

    Set fso = CreateObject("Scripting.FileSystemObject")

    s = Trim(CStr(s))
    t = ExpandEnv(StripOuterQuotes(s))

    If Len(s) = 0 Then
        IsValidLaunchTarget = False
        Exit Function
    End If

    If IsCmdTarget(s) Then
        IsValidLaunchTarget = True
        Exit Function
    End If

    If IsUrlTarget(s) Then
        IsValidLaunchTarget = True
        Exit Function
    End If

    If fso.FileExists(t) Then
        IsValidLaunchTarget = True
        Exit Function
    End If

    If fso.FolderExists(t) Then
        IsValidLaunchTarget = True
        Exit Function
    End If

    IsValidLaunchTarget = False
End Function

Private Function GetLaunchTargetName(ByVal s)
    Dim fso
    Dim t

    Set fso = CreateObject("Scripting.FileSystemObject")

    t = Trim(CStr(s))

    If IsCmdTarget(t) Then
        t = Trim(Mid(t, 5))

        If Len(t) > 40 Then
            t = Left(t, 40) & "..."
        End If

        GetLaunchTargetName = "命令：" & t
        Exit Function
    End If

    If IsUrlTarget(t) Then
        If Len(t) > 50 Then
            GetLaunchTargetName = "网址：" & Left(t, 50) & "..."
        Else
            GetLaunchTargetName = "网址：" & t
        End If

        Exit Function
    End If

    t = ExpandEnv(StripOuterQuotes(t))

    If fso.FileExists(t) Then
        GetLaunchTargetName = fso.GetFileName(t)
        Exit Function
    End If

    If fso.FolderExists(t) Then
        GetLaunchTargetName = fso.GetFolder(t).Name
        Exit Function
    End If

    GetLaunchTargetName = s
End Function

Private Function GetCurrentDocPath()
    Dim p

    p = ""

    On Error Resume Next
    Err.Clear

    p = ActiveWorkbook.FullName

    If Err.Number <> 0 Then
        Err.Clear
        p = ""
    End If

    If p = "" Then
        p = ActiveDocument.FullName

        If Err.Number <> 0 Then
            Err.Clear
            p = ""
        End If
    End If

    On Error GoTo 0

    GetCurrentDocPath = p
End Function

Private Function GetTempVbsPath()
    Dim tempDir
    Dim fileName

    Randomize

    tempDir = Environ("TEMP")

    If Len(tempDir) = 0 Then
        tempDir = Environ("TMP")
    End If

    If Len(tempDir) = 0 Then
        tempDir = "C:\Windows\Temp"
    End If

    fileName = "wps_launcher_" & Format(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd() * 100000)) & ".vbs"

    GetTempVbsPath = tempDir & "\" & fileName
End Function

Private Function RunExternalVbs(ByVal vbsText)
    On Error GoTo ErrHandler

    Dim vbsPath
    Dim cmdLine

    vbsPath = GetTempVbsPath()

    WriteUnicodeFile vbsPath, vbsText

    cmdLine = "wscript.exe //B //Nologo " & QuoteCmdArg(vbsPath)

    shell cmdLine, vbHide

    RunExternalVbs = True
    Exit Function

ErrHandler:
    MsgBox "外部启动器调用失败：" & vbCrLf & Err.Description, vbExclamation
    RunExternalVbs = False
End Function

Private Function BuildOpenTargetVbs(ByVal target)
    Dim v

    v = ""
    v = v & "On Error Resume Next" & vbCrLf
    v = v & "Dim sh, app, fso, target, ext, cmd1, cmd2" & vbCrLf
    v = v & "target = " & VbsQuote(target) & vbCrLf
    v = v & "Set sh = CreateObject(""WScript.Shell"")" & vbCrLf
    v = v & "Set app = CreateObject(""Shell.Application"")" & vbCrLf
    v = v & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    v = v & "ext = LCase(fso.GetExtensionName(target))" & vbCrLf
    v = v & "If ext = ""lnk"" Then" & vbCrLf
    v = v & "    Err.Clear" & vbCrLf
    v = v & "    sh.Run Chr(34) & target & Chr(34), 1, False" & vbCrLf
    v = v & "    If Err.Number <> 0 Then" & vbCrLf
    v = v & "        Err.Clear" & vbCrLf
    v = v & "        app.ShellExecute target, """", """", ""open"", 1" & vbCrLf
    v = v & "    End If" & vbCrLf
    v = v & "    If Err.Number <> 0 Then" & vbCrLf
    v = v & "        Err.Clear" & vbCrLf
    v = v & "        cmd1 = ""rundll32.exe url.dll,FileProtocolHandler "" & Chr(34) & target & Chr(34)" & vbCrLf
    v = v & "        sh.Run cmd1, 1, False" & vbCrLf
    v = v & "    End If" & vbCrLf
    v = v & "Else" & vbCrLf
    v = v & "    Err.Clear" & vbCrLf
    v = v & "    app.ShellExecute target, """", """", ""open"", 1" & vbCrLf
    v = v & "    If Err.Number <> 0 Then" & vbCrLf
    v = v & "        Err.Clear" & vbCrLf
    v = v & "        cmd2 = ""rundll32.exe url.dll,FileProtocolHandler "" & Chr(34) & target & Chr(34)" & vbCrLf
    v = v & "        sh.Run cmd2, 1, False" & vbCrLf
    v = v & "    End If" & vbCrLf
    v = v & "End If" & vbCrLf
    v = v & "Err.Clear" & vbCrLf
    v = v & "fso.DeleteFile WScript.ScriptFullName, True" & vbCrLf

    BuildOpenTargetVbs = v
End Function

Private Function BuildCmdTargetVbs(ByVal cmdText)
    Dim v

    v = ""
    v = v & "On Error Resume Next" & vbCrLf
    v = v & "Dim sh, fso, cmdText, cmdExe, fullCmd" & vbCrLf
    v = v & "Set sh = CreateObject(""WScript.Shell"")" & vbCrLf
    v = v & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    v = v & "cmdText = " & VbsQuote(cmdText) & vbCrLf
    v = v & "cmdExe = sh.ExpandEnvironmentStrings(""%COMSPEC%"")" & vbCrLf
    v = v & "If Len(cmdExe) = 0 Then cmdExe = ""cmd.exe""" & vbCrLf
    v = v & "fullCmd = Chr(34) & cmdExe & Chr(34) & "" /d /c "" & cmdText" & vbCrLf
    v = v & "sh.Run fullCmd, 0, False" & vbCrLf
    v = v & "Err.Clear" & vbCrLf
    v = v & "fso.DeleteFile WScript.ScriptFullName, True" & vbCrLf

    BuildCmdTargetVbs = v
End Function

Private Function LaunchTarget(ByVal rawTarget)
    Dim fso
    Dim target
    Dim cmd
    Dim docPath
    Dim vbsText

    Set fso = CreateObject("Scripting.FileSystemObject")

    target = Trim(CStr(rawTarget))
    docPath = GetCurrentDocPath()

    If target = "" Then
        MsgBox "启动目标为空。", vbExclamation
        LaunchTarget = False
        Exit Function
    End If

    On Error GoTo RunErr

    If IsCmdTarget(target) Then
        cmd = Trim(Mid(target, 5))

        If cmd = "" Then
            MsgBox "cmd: 后面没有命令。", vbExclamation
            LaunchTarget = False
            Exit Function
        End If

        cmd = Replace(cmd, "{doc}", QuoteCmdArg(docPath))
        cmd = Replace(cmd, "{doc_raw}", docPath)

        vbsText = BuildCmdTargetVbs(cmd)

        LaunchTarget = RunExternalVbs(vbsText)
        Exit Function
    End If

    target = Replace(target, "{doc}", docPath)
    target = Replace(target, "{doc_raw}", docPath)
    target = StripOuterQuotes(target)
    target = ExpandEnv(target)

    If Not IsUrlTarget(target) Then
        If Not fso.FileExists(target) And Not fso.FolderExists(target) Then
            MsgBox "启动目标不存在：" & vbCrLf & target, vbExclamation
            LaunchTarget = False
            Exit Function
        End If
    End If

    vbsText = BuildOpenTargetVbs(target)

    LaunchTarget = RunExternalVbs(vbsText)
    Exit Function

RunErr:
    MsgBox "启动失败：" & Err.Description, vbExclamation
    LaunchTarget = False
End Function

' =================== 主菜单 ===================
Private Sub ShowMainMenu()
    Dim ids(MAX_SCRIPTS)
    Dim names(MAX_SCRIPTS)
    Dim remarks(MAX_SCRIPTS)
    Dim paths(MAX_SCRIPTS)

    Dim sc
    Dim choice
    Dim menu
    Dim i
    Dim uc
    Dim idx
    Dim fsoTest

    Do
        sc = ReadScripts(ids, names, remarks, paths)

        If sc = 0 Then
            If MsgBox("还没有注册启动项。是否添加？", vbYesNo + vbQuestion) = vbNo Then
                Exit Do
            End If

            AddScript

            Set fsoTest = CreateObject("Scripting.FileSystemObject")

            If Not fsoTest.FileExists(CONFIG_PATH) Then
                Exit Do
            End If

        Else
            menu = "WPS 启动机器人" & vbCrLf
            menu = menu & "[A]添加  [D]删除  [M]备注  [R]刷新  [Q]退出" & vbCrLf
            menu = menu & "----------------------------------------" & vbCrLf & vbCrLf

            For i = 1 To sc
                menu = menu & " [" & i & "] " & names(i) & vbCrLf

                If Len(remarks(i)) > 0 Then
                    menu = menu & "    " & remarks(i) & vbCrLf
                End If

                menu = menu & "    目标：" & paths(i) & vbCrLf
            Next

            menu = menu & vbCrLf & "请输入数字启动项目："

            choice = InputBox(menu, "WPS 启动机器人")

            If choice = "" Then
                Exit Do
            End If

            uc = UCase(Trim(choice))

            If uc = "Q" Then
                Exit Do

            ElseIf uc = "A" Then
                AddScript

            ElseIf uc = "D" Then
                DeleteScript

            ElseIf uc = "M" Then
                ModifyRemark

            ElseIf uc = "R" Then
                ' 刷新

            ElseIf IsNumeric(choice) Then
                idx = CLng(choice)

                If idx >= 1 And idx <= sc Then
                    RunOneScript paths(idx)
                    Exit Do
                Else
                    MsgBox "请输入 1~" & sc & " 之间的数字", vbExclamation
                End If

            Else
                MsgBox "请输入数字或 A/D/M/R/Q", vbExclamation
            End If
        End If
    Loop
End Sub

' =================== 添加启动项 ===================
Private Sub AddScript()
    Dim path
    Dim targetName
    Dim remark

    path = PickTarget()

    If path = "" Then
        Exit Sub
    End If

    If Not IsValidLaunchTarget(path) Then
        MsgBox "启动目标无效：" & vbCrLf & path & vbCrLf & vbCrLf & _
               "可以是文件、文件夹、网址、快捷方式，或以 cmd: 开头的命令。", vbExclamation
        Exit Sub
    End If

    targetName = InputBox("启动项名称：", "添加启动项", GetLaunchTargetName(path))

    If targetName = "" Then
        Exit Sub
    End If

    remark = InputBox("备注说明：", "添加启动项", "")

    If remark = "" Then
        remark = "(无备注)"
    End If

    SaveNewScript path, targetName, remark

    MsgBox "已添加启动项：" & targetName, vbInformation
End Sub

Private Function PickTarget()
    Dim p
    Dim sel

    p = InputBox( _
        "请直接粘贴启动目标路径。" & vbCrLf & vbCrLf & _
        "重要提示：" & vbCrLf & _
        "如果你要添加桌面快捷方式，请手动粘贴 .lnk 路径，" & vbCrLf & _
        "不要优先使用文件选择器，否则 WPS 可能会自动解析快捷方式。" & vbCrLf & vbCrLf & _
        "示例：" & vbCrLf & _
        "C:\Users\hf\Desktop\demo.lnk" & vbCrLf & _
        "C:\Tools\app.exe" & vbCrLf & _
        "C:\Users\hf\Desktop\project" & vbCrLf & _
        "https://example.com" & vbCrLf & _
        "cmd:python C:\bot\main.py" & vbCrLf & vbCrLf & _
        "如果留空，将询问是否打开文件选择器。", _
        "添加启动项", "")

    If p <> "" Then
        p = StripOuterQuotes(Trim(CStr(p)))
        PickTarget = p
        Exit Function
    End If

    If MsgBox("是否改用文件选择器？" & vbCrLf & vbCrLf & _
              "注意：文件选择器可能无法正确选择 .lnk 快捷方式本身。", _
              vbYesNo + vbQuestion) = vbNo Then
        PickTarget = ""
        Exit Function
    End If

    On Error Resume Next
    Err.Clear

    sel = Application.GetOpenFilename("所有文件 (*.*),*.*", 1, "选择要启动的文件")

    If Err.Number = 0 Then
        If Not IsEmpty(sel) Then
            If VarType(sel) = vbString Then
                If CStr(sel) <> "False" And Len(CStr(sel)) > 0 Then
                    PickTarget = CStr(sel)
                    Exit Function
                End If
            End If
        End If
    End If

    Err.Clear
    On Error GoTo 0

    PickTarget = ""
End Function

' =================== 删除启动项 ===================
Private Sub DeleteScript()
    Dim ids(MAX_SCRIPTS)
    Dim names(MAX_SCRIPTS)
    Dim remarks(MAX_SCRIPTS)
    Dim paths(MAX_SCRIPTS)

    Dim sc
    Dim menu
    Dim i
    Dim choice
    Dim idx

    sc = ReadScripts(ids, names, remarks, paths)

    If sc = 0 Then
        MsgBox "没有启动项可删除", vbInformation
        Exit Sub
    End If

    menu = "选择要删除的启动项：" & vbCrLf

    For i = 1 To sc
        menu = menu & " [" & i & "] " & names(i)

        If Len(remarks(i)) > 0 Then
            menu = menu & " — " & remarks(i)
        End If

        menu = menu & vbCrLf
        menu = menu & "    目标：" & paths(i) & vbCrLf
    Next

    choice = InputBox(menu, "删除启动项")

    If choice = "" Then
        Exit Sub
    End If

    If Not IsNumeric(choice) Then
        MsgBox "请输入数字", vbExclamation
        Exit Sub
    End If

    idx = CLng(choice)

    If idx < 1 Or idx > sc Then
        MsgBox "请输入 1~" & sc, vbExclamation
        Exit Sub
    End If

    If MsgBox("确定删除 [" & idx & "] " & names(idx) & "？", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    RemoveById ids(idx)

    MsgBox "已删除：" & names(idx), vbInformation
End Sub

' =================== 修改备注 ===================
Private Sub ModifyRemark()
    Dim ids(MAX_SCRIPTS)
    Dim names(MAX_SCRIPTS)
    Dim remarks(MAX_SCRIPTS)
    Dim paths(MAX_SCRIPTS)

    Dim sc
    Dim menu
    Dim i
    Dim choice
    Dim idx
    Dim nr

    sc = ReadScripts(ids, names, remarks, paths)

    If sc = 0 Then
        MsgBox "没有启动项可修改", vbInformation
        Exit Sub
    End If

    menu = "选择要修改备注的启动项：" & vbCrLf

    For i = 1 To sc
        menu = menu & " [" & i & "] " & names(i)

        If Len(remarks(i)) > 0 Then
            menu = menu & " — " & remarks(i)
        End If

        menu = menu & vbCrLf
    Next

    choice = InputBox(menu, "修改备注")

    If choice = "" Then
        Exit Sub
    End If

    If Not IsNumeric(choice) Then
        MsgBox "请输入数字", vbExclamation
        Exit Sub
    End If

    idx = CLng(choice)

    If idx < 1 Or idx > sc Then
        MsgBox "请输入 1~" & sc, vbExclamation
        Exit Sub
    End If

    nr = InputBox("新的备注：", "修改备注", remarks(idx))

    If nr = "" Then
        nr = "(无备注)"
    End If

    UpdateRemark ids(idx), nr

    MsgBox "备注已更新", vbInformation
End Sub

' =================== 启动一个项目 ===================
Private Sub RunOneScript(ByVal fullPath)
    LaunchTarget fullPath
End Sub

