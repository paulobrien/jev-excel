Attribute VB_Name = "modJev"
'==============================================================================
'  Jev for Excel - VBA edition  (Windows desktop Excel)
'
'  Worksheet functions
'    =JEV(mode, text, [arg1], [arg2], ...)
'        mode : "choice" | "score" | "noul", optionally ":output"
'               e.g. "choice:probs", "score:label", "noul:bool:0.7"
'        text : the text (or cell / range) Jev should evaluate
'        args : choice -> option labels ("label | description" or a 2-col range)
'               score  -> rubric levels, lowest (level 0) first
'               noul   -> the yes/no question, then optional TRUE / FALSE descriptions
'               Any argument starting with "?" is used as the question.
'    =JEV_STATE(headers, values)   -> JSON object text for named-field state
'    =JEV_MODELS()                 -> spills name | description | release date
'
'  Macros (Alt+F8)
'    JevSetApiKey, JevSetModel, JevSetBaseUrl, JevClearCache, JevShowStats,
'    JevAddModeNames, JevToggleErrorStyle, JevRegisterFunctions
'
'  Settings live in HKCU\Software\VB and VBA Program Settings\JevForExcel.
'  The API key falls back to the TYPESAFE_API_KEY environment variable.
'
'  Unofficial; not affiliated with TypeSafe AI.  MIT licence.
'==============================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const JEV_VERSION As String = "1.0.1"
Private Const DEFAULT_BASE_URL As String = "https://api.typesafe.ai"
Private Const DEFAULT_MODEL As String = "jev-latest"
Private Const REG_APP As String = "JevForExcel"
Private Const REG_SECTION As String = "Settings"
Private Const MAX_RETRIES As Long = 3

' MSXML2.XMLHTTP.6.0 uses the Windows/Internet Options proxy settings, which suits
' most corporate networks, but it has no timeout: a stalled server blocks Excel until
' Windows gives up. "MSXML2.ServerXMLHTTP.6.0" and "WinHttp.WinHttpRequest.5.1" honour
' the timeouts below but use the WinHTTP proxy settings (netsh winhttp) instead.
Private Const HTTP_CLASS As String = "MSXML2.XMLHTTP.6.0"
Private Const TIMEOUT_MS As Long = 15000
Private Const MAX_WAIT_MS As Long = 10000     ' Excel is blocked while we wait, so keep it short
Private Const FAIL_FAST_SECONDS As Long = 30
Private Const JEV_ERR As Long = vbObjectError + 513

Private mCache As Object          ' Scripting.Dictionary: cache key -> answer Dictionary
Private mJson As String           ' JSON parser state
Private mPos As Long
Private mCalls As Long, mHits As Long, mRequests As Long, mTokens As Double, mLastModel As String

' After an authentication failure, or when Jev can't be reached, the remaining cells fail
' fast with the same message instead of each one retrying. Changing the key or base URL,
' or running JevClearCache, resets this.
Private mFailMsg As String, mFailFor As String, mFailUntil As Date

'==============================================================================
'  Worksheet functions
'==============================================================================

Public Function JEV(Mode As Variant, Text As Variant, ParamArray Args() As Variant) As Variant
    On Error GoTo Fail
    mCalls = mCalls + 1

    Dim modeText As String, modeVal As Variant
    modeVal = ScalarOf(Mode)
    If IsError(modeVal) Then
        ' Best effort for =JEV(choice, ...) without quotes: an undefined name arrives as #NAME?
        If modeVal = CVErr(xlErrName) Then modeText = ModeFromFormula()
        If Len(modeText) = 0 Then Raise "Mode is required: ""choice"", ""score"" or ""noul"" (in quotes)."
    Else
        modeText = CStr(modeVal)
    End If

    Dim qType As String, outKind As String, threshold As Double
    ParseMode modeText, qType, outKind, threshold

    Dim stateJson As String
    stateJson = StateToJson(Text)
    If Len(stateJson) = 0 Then JEV = "": Exit Function     ' blank in -> blank out, no API call

    Dim levels As Collection, argCopy As Variant
    Set levels = New Collection
    argCopy = Args
    Dim questionJson As String
    questionJson = BuildQuestionJson(qType, argCopy, levels)

    Dim answer As Object
    Set answer = GetAnswer(stateJson, questionJson)

    JEV = FormatAnswer(answer, qType, outKind, threshold, levels)
    Exit Function
Fail:
    JEV = ErrorResult(ErrMessage())
End Function

Public Function JEV_STATE(Headers As Variant, Values As Variant) As Variant
    On Error GoTo Fail
    Dim h As Collection, v As Collection, i As Long, parts() As String, k As Variant, fields As Object
    Set h = New Collection: Set v = New Collection
    FlattenAll Headers, h
    FlattenAll Values, v
    If h.Count <> v.Count Then Raise "JEV_STATE needs the same number of headers (" & h.Count & ") and values (" & v.Count & ")."
    Set fields = CreateObject("Scripting.Dictionary")
    fields.CompareMode = 0
    For i = 1 To h.Count
        If Not IsBlankValue(h(i)) Then fields(TrimWs(ValueText(h(i)))) = ValueJson(v(i))   ' last duplicate wins
    Next
    If fields.Count = 0 Then Raise "JEV_STATE needs at least one non-empty header."
    ReDim parts(0 To fields.Count - 1)
    i = 0
    For Each k In fields.Keys
        parts(i) = JStr(CStr(k)) & ":" & fields(k)
        i = i + 1
    Next
    JEV_STATE = "{" & Join(parts, ",") & "}"
    Exit Function
Fail:
    JEV_STATE = ErrorResult(ErrMessage())
End Function

Public Function JEV_MODELS() As Variant
    On Error GoTo Fail
    Dim data As Object, models As Object, i As Long, m As Object, out() As Variant
    Set data = JsonParseObject(HttpRequest("GET", "/v1/models", ""))
    Set models = DictObj(data, "models")
    If models.Count = 0 Then JEV_MODELS = "(no models)": Exit Function
    ReDim out(1 To models.Count, 1 To 3)
    For i = 1 To models.Count
        Set m = models(i)
        out(i, 1) = DictText(m, "name")
        out(i, 2) = DictText(m, "description")
        out(i, 3) = DictText(m, "release_date")
    Next
    JEV_MODELS = out
    Exit Function
Fail:
    JEV_MODELS = ErrorResult(ErrMessage())
End Function

'==============================================================================
'  Macros
'==============================================================================

Public Sub JevSetApiKey()
    Dim k As String
    k = InputBox("Paste your TypeSafe API key." & vbCrLf & vbCrLf & _
                 "It is stored in your Windows user registry on this PC." & vbCrLf & _
                 "Leave blank to fall back to the TYPESAFE_API_KEY environment variable.", _
                 "Jev for Excel - API key")
    If StrPtr(k) = 0 Then Exit Sub   ' cancelled
    SaveSetting REG_APP, REG_SECTION, "ApiKey", Trim$(k)
    MsgBox "API key saved. Recalculating open workbooks...", vbInformation, "Jev for Excel"
    Application.CalculateFull
End Sub

Public Sub JevSetModel()
    Dim m As String
    m = InputBox("Model to use (e.g. jev-latest, jev-preview, jev-1.13.0):", "Jev for Excel - model", ModelName())
    If StrPtr(m) = 0 Then Exit Sub
    SaveSetting REG_APP, REG_SECTION, "Model", Trim$(m)
    Application.CalculateFull
End Sub

Public Sub JevSetBaseUrl()
    Dim u As String
    u = InputBox("API base URL (default " & DEFAULT_BASE_URL & ", or your proxy URL):", "Jev for Excel - base URL", BaseUrl())
    If StrPtr(u) = 0 Then Exit Sub
    u = Trim$(u)
    If Len(u) > 0 And Not IsSafeUrl(u) Then
        MsgBox "The base URL must start with https:// (http:// is allowed only for localhost).", vbExclamation, "Jev for Excel"
        Exit Sub
    End If
    SaveSetting REG_APP, REG_SECTION, "BaseUrl", u
    Application.CalculateFull
End Sub

Public Sub JevClearCache()
    Set mCache = Nothing
    mFailMsg = ""
    mCalls = 0: mHits = 0: mRequests = 0: mTokens = 0
    MsgBox "Jev cache cleared. The next recalculation (Ctrl+Alt+F9) will call Jev again.", vbInformation, "Jev for Excel"
End Sub

Public Sub JevShowStats()
    Dim cached As Long
    If Not mCache Is Nothing Then cached = mCache.Count
    MsgBox "Jev for Excel " & JEV_VERSION & vbCrLf & vbCrLf & _
           "Model: " & IIf(Len(mLastModel) > 0, mLastModel, ModelName()) & vbCrLf & _
           "Base URL: " & BaseUrl() & vbCrLf & _
           "API key: " & IIf(Len(ApiKey()) > 0, "set", "NOT SET") & vbCrLf & vbCrLf & _
           "Function calls: " & mCalls & vbCrLf & _
           "Served from cache: " & mHits & vbCrLf & _
           "HTTP requests: " & mRequests & vbCrLf & _
           "Input tokens: " & Format$(mTokens, "#,##0") & vbCrLf & _
           "Cached answers: " & cached, vbInformation, "Jev for Excel"
End Sub

' Adds workbook names choice / score / noul so =JEV(choice, A2, ...) works without quotes.
' Existing names with those spellings are replaced.
Public Sub JevAddModeNames()
    Dim wb As Workbook, nm As Variant
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    On Error GoTo Fail
    For Each nm In Array("choice", "score", "noul")
        wb.Names.Add Name:=CStr(nm), RefersTo:="=""" & nm & """"
    Next
    MsgBox "Added names choice, score and noul to " & wb.Name & ".", vbInformation, "Jev for Excel"
    Exit Sub
Fail:
    MsgBox "Could not add the names to " & wb.Name & ": " & Err.Description, vbExclamation, "Jev for Excel"
End Sub

' Errors as "#JEV! message" text (default, easier to debug) or as #VALUE!.
Public Sub JevToggleErrorStyle()
    Dim asText As Boolean
    asText = (GetSetting(REG_APP, REG_SECTION, "ErrorsAsText", "1") = "1")
    SaveSetting REG_APP, REG_SECTION, "ErrorsAsText", IIf(asText, "0", "1")
    MsgBox "Jev errors will now show as " & IIf(asText, "#VALUE!", """#JEV! message"" text") & ".", vbInformation, "Jev for Excel"
End Sub

' Adds descriptions to the Insert Function dialog. Call from ThisWorkbook.Workbook_Open.
' Best effort: MacroOptions can fail (e.g. while Excel is starting), and that's harmless.
Public Sub JevRegisterFunctions()
    On Error Resume Next
    Dim book As String
    book = "'" & ThisWorkbook.Name & "'!"
    Application.MacroOptions Macro:=book & "JEV", Category:="Jev", _
        Description:="Ask Jev a typed question about text. Mode: choice | score | noul (optionally :output).", _
        ArgumentDescriptions:=Array( _
            "choice | score | noul, optionally :output e.g. choice:probs, score:label, noul:bool:0.7", _
            "The text or cell Jev should evaluate. Blank returns blank.", _
            "Options, rubric levels or the yes/no question. Prefix with ? to set the question.")
    Application.MacroOptions Macro:=book & "JEV_STATE", Category:="Jev", _
        Description:="Build a JSON object of named fields for Jev from a header range and a value range.", _
        ArgumentDescriptions:=Array("Field names, e.g. $A$1:$C$1", "Field values, e.g. A2:C2")
    Application.MacroOptions Macro:=book & "JEV_MODELS", Category:="Jev", _
        Description:="List the Jev models available to your API key."
End Sub

'==============================================================================
'  Mode parsing
'==============================================================================

Private Sub ParseMode(ByVal modeText As String, ByRef qType As String, ByRef outKind As String, ByRef threshold As Double)
    Dim parts() As String, t As String, o As String
    parts = Split(LCase$(TrimWs(modeText)), ":")
    If UBound(parts) < 0 Then Raise "Mode is required: ""choice"", ""score"" or ""noul""."
    If UBound(parts) > 2 Then Raise "Mode """ & modeText & """ has too many parts."
    t = TrimWs(parts(0))
    If UBound(parts) >= 1 Then o = TrimWs(parts(1))

    Select Case t
        Case "choice", "choose", "classify", "category", "pick", "select": qType = "choice"
        Case "score", "rate", "rating", "scale", "rubric": qType = "score"
        Case "noul", "yesno", "yn", "bool", "boolean", "binary": qType = "noul"
        Case Else: Raise "Unknown mode """ & t & """. Use ""choice"", ""score"" or ""noul""."
    End Select

    Select Case qType
        Case "choice"
            Select Case o
                Case "", "label", "choice": outKind = "label"
                Case "confidence", "conf": outKind = "confidence"
                Case "prob", "probability": outKind = "prob"
                Case "probs", "probabilities": outKind = "probs"
                Case "json": outKind = "json"
                Case Else: Raise "Unknown output """ & o & """ for choice. Valid: label, confidence, prob, probs, json."
            End Select
        Case "score"
            Select Case o
                Case "", "score": outKind = "score"
                Case "level", "round": outKind = "level"
                Case "label", "legend": outKind = "label"
                Case "confidence", "conf": outKind = "confidence"
                Case "probs", "probabilities": outKind = "probs"
                Case "json": outKind = "json"
                Case Else: Raise "Unknown output """ & o & """ for score. Valid: score, level, label, confidence, probs, json."
            End Select
        Case "noul"
            Select Case o
                Case "", "prob", "probability", "noul": outKind = "prob"
                Case "bool", "boolean": outKind = "bool"
                Case "yesno": outKind = "yesno"
                Case "json": outKind = "json"
                Case Else: Raise "Unknown output """ & o & """ for noul. Valid: prob, bool, yesno, json."
            End Select
    End Select

    threshold = 0.5
    If UBound(parts) = 2 Then
        Dim th As String
        th = TrimWs(parts(2))
        If Len(th) > 0 Then
            If qType <> "noul" Or (outKind <> "bool" And outKind <> "yesno") Then Raise "A threshold is only valid with noul:bool or noul:yesno."
            If IsJsonNumber(th) Then threshold = Val(th) Else threshold = -1
            If threshold <= 0 Or threshold >= 1 Then Raise "Threshold must be a number between 0 and 1, e.g. noul:bool:0.7."
        End If
    End If
End Sub

Private Function ModeFromFormula() As String
    On Error GoTo Nope
    Dim f As String, p As Long, q As Long, tok As String
    If TypeName(Application.Caller) <> "Range" Then Exit Function
    f = Application.Caller.Cells(1, 1).Formula
    p = InStr(1, f, "JEV(", vbTextCompare)
    If p = 0 Then Exit Function
    p = p + 4
    q = InStr(p, f, ",")
    If q = 0 Then Exit Function
    tok = Replace(Trim$(Mid$(f, p, q - p)), """", "")
    If Len(tok) = 0 Or tok Like "*[!A-Za-z0-9:._]*" Then Exit Function
    ModeFromFormula = tok
Nope:
End Function

'==============================================================================
'  State and question building
'==============================================================================

Private Function StateToJson(Text As Variant) As String
    Dim vals As Collection, i As Long, s As String, t As String, parts() As String
    Set vals = New Collection
    FlattenNonBlank Text, vals
    If vals.Count = 0 Then Exit Function

    If vals.Count = 1 Then
        s = vals(1)
        t = TrimWs(s)
        If (Left$(t, 1) = "{" And Right$(t, 1) = "}") Or (Left$(t, 1) = "[" And Right$(t, 1) = "]") Then
            If IsValidJson(t) Then StateToJson = t: Exit Function
        End If
        StateToJson = JStr(s)
    Else
        ReDim parts(0 To vals.Count - 1)
        For i = 1 To vals.Count
            parts(i - 1) = JStr(vals(i))
        Next
        StateToJson = "[" & Join(parts, ",") & "]"
    End If
End Function

' Returns the question JSON; fills `levels` (score rubric) for label lookups.
Private Function BuildQuestionJson(ByVal qType As String, ByVal argList As Variant, ByVal levels As Collection) As String
    Dim questions As Collection, labels As Collection, descs As Collection
    Set questions = New Collection: Set labels = New Collection: Set descs = New Collection

    Dim i As Long, arr As Variant, r As Long, c As Long
    Dim flat As Collection, item As Variant, s As String
    If IsArray(argList) Then
        For i = LBound(argList) To UBound(argList)
            If IsMissing(argList(i)) Then GoTo NextArg

            ' A range / array with 2+ rows and exactly 2 columns = choice label/description pairs
            If qType = "choice" Then
                If IsObject(argList(i)) Then arr = argList(i).Value2 Else arr = argList(i)
                If ArrayDims(arr) = 2 Then
                    If (UBound(arr, 1) - LBound(arr, 1) + 1) >= 2 And (UBound(arr, 2) - LBound(arr, 2) + 1) = 2 Then
                        For r = LBound(arr, 1) To UBound(arr, 1)
                            c = LBound(arr, 2)
                            If Not IsBlankValue(arr(r, c)) Then
                                s = TrimWs(ValueText(arr(r, c)))
                                If IsQuestion(s) Then
                                    questions.Add TrimWs(Mid$(s, 2))
                                Else
                                    labels.Add s
                                    If IsBlankValue(arr(r, c + 1)) Then descs.Add Null Else descs.Add TrimWs(ValueText(arr(r, c + 1)))
                                End If
                            End If
                        Next
                        GoTo NextArg
                    End If
                End If
            End If

            Set flat = New Collection
            FlattenNonBlank argList(i), flat
            For Each item In flat
                s = TrimWs(CStr(item))
                If IsQuestion(s) Then
                    questions.Add TrimWs(Mid$(s, 2))
                Else
                    labels.Add s
                    descs.Add Empty      ' Empty = not given (may come from "label | description")
                End If
            Next
NextArg:
        Next
    End If

    If questions.Count > 1 Then Raise "Only one ""?question"" argument is allowed."
    Dim instr_ As String, hasInstr As Boolean
    If questions.Count = 1 Then instr_ = questions(1): hasInstr = True

    Dim json As String, parts() As String, n As Long, lbl As String, d As Variant, p As Long
    Select Case qType
        Case "choice"
            Dim seen As Object
            Set seen = CreateObject("Scripting.Dictionary")
            seen.CompareMode = 0
            If labels.Count < 2 Then Raise "choice needs at least 2 options."
            If labels.Count > 255 Then Raise "choice supports at most 255 options."
            ReDim parts(0 To labels.Count - 1)
            For n = 1 To labels.Count
                lbl = labels(n)
                d = descs(n)
                If IsEmpty(d) Then
                    p = InStr(1, lbl, "|")
                    If p > 0 Then
                        d = TrimWs(Mid$(lbl, p + 1))
                        If Len(d) = 0 Then d = Null
                        lbl = TrimWs(Left$(lbl, p - 1))
                    Else
                        d = Null
                    End If
                End If
                If Len(lbl) = 0 Then Raise "Choice labels cannot be empty."
                If seen.Exists(lbl) Then Raise "Duplicate choice label """ & lbl & """."
                seen.Add lbl, True
                If IsNull(d) Then
                    parts(n - 1) = JStr(lbl) & ":null"
                Else
                    parts(n - 1) = JStr(lbl) & ":" & JStr(CStr(d))
                End If
            Next
            json = "{""type"":""choice"""
            If hasInstr Then json = json & ",""instructions"":" & JStr(instr_)
            json = json & ",""criteria"":{" & Join(parts, ",") & "}}"

        Case "score"
            If labels.Count < 2 Then Raise "score needs at least 2 rubric levels."
            If labels.Count > 10 Then Raise "score supports at most 10 rubric levels."
            ReDim parts(0 To labels.Count - 1)
            For n = 1 To labels.Count
                parts(n - 1) = JStr(labels(n))
                levels.Add labels(n)
            Next
            json = "{""type"":""score"""
            If hasInstr Then json = json & ",""instructions"":" & JStr(instr_)
            json = json & ",""criteria"":[" & Join(parts, ",") & "]}"

        Case "noul"
            Dim first As Long
            first = 1
            If Not hasInstr Then
                If labels.Count = 0 Then Raise "noul needs a yes/no question, e.g. ""Is the customer asking for a refund?"""
                instr_ = labels(1): hasInstr = True: first = 2
            End If
            If labels.Count - first + 1 > 2 Then Raise "noul takes a question plus at most two descriptions (TRUE, then FALSE)."
            json = "{""type"":""noul"",""instructions"":" & JStr(instr_)
            If labels.Count >= first Then
                json = json & ",""criteria"":{""true"":" & JStr(labels(first))
                If labels.Count >= first + 1 Then json = json & ",""false"":" & JStr(labels(first + 1))
                json = json & "}"
            End If
            json = json & "}"
    End Select
    BuildQuestionJson = json
End Function

'==============================================================================
'  API calls and cache
'==============================================================================

Private Function GetAnswer(ByVal stateJson As String, ByVal questionJson As String) As Object
    Dim model As String, key As String
    model = ModelName()
    key = model & vbNullChar & stateJson & vbNullChar & questionJson
    If mCache Is Nothing Then
        Set mCache = CreateObject("Scripting.Dictionary")
        mCache.CompareMode = 0
    End If
    If mCache.Exists(key) Then
        mHits = mHits + 1
        Set GetAnswer = mCache(key)
        Exit Function
    End If

    Dim body As String, resp As Object, answers As Object
    body = "{""model"":" & JStr(model) & ",""state"":" & stateJson & ",""questions"":{""q0"":" & questionJson & "}}"
    Set resp = JsonParseObject(HttpRequest("POST", "/v1/systemone", body))

    If resp.Exists("usage") Then
        If IsObject(resp("usage")) Then mTokens = mTokens + Val(DictText(resp("usage"), "input_tokens"))
    End If
    If resp.Exists("model") Then mLastModel = DictText(resp, "model")
    Set answers = DictObj(resp, "answers")
    If Not answers.Exists("q0") Then Raise "Jev returned no answer for this question."
    If Not IsObject(answers("q0")) Then Raise "Empty answer from Jev."

    If mCache.Count > 20000 Then mCache.RemoveAll
    mCache.Add key, answers("q0")
    Set GetAnswer = answers("q0")
End Function

Private Function HttpRequest(ByVal method As String, ByVal path As String, ByVal body As String) As String
    Dim base As String, url As String, key As String, target As String
    base = BaseUrl()
    url = base & path
    key = ApiKey()
    If Not IsSafeUrl(base) Then Raise "The base URL must start with https:// (http:// is allowed only for localhost). Run JevSetBaseUrl."
    If Len(key) = 0 Then Raise "No Jev API key. Run the JevSetApiKey macro (Alt+F8) or set TYPESAFE_API_KEY."

    ' Fail fast while an earlier auth or network failure for this key and URL still stands.
    target = base & vbNullChar & key
    If Len(mFailMsg) > 0 Then
        If mFailFor = target And Now < mFailUntil Then Raise mFailMsg
        mFailMsg = ""
    End If

    Dim attempt As Long, http As Object, status As Long, resp As String, netErr As String, waitMs As Long
    For attempt = 0 To MAX_RETRIES
        mRequests = mRequests + 1
        netErr = ""
        Set http = CreateObject(HTTP_CLASS)
        On Error Resume Next
        http.setTimeouts TIMEOUT_MS, TIMEOUT_MS, TIMEOUT_MS, TIMEOUT_MS   ' not supported by XMLHTTP; ignored
        Err.Clear
        http.Open method, url, False
        http.setRequestHeader "Accept", "application/json"
        http.setRequestHeader "Cache-Control", "no-cache"
        If Len(body) > 0 Then http.setRequestHeader "Content-Type", "application/json"
        http.setRequestHeader "Authorization", "Bearer " & key
        If Len(body) > 0 Then http.send body Else http.send
        If Err.Number <> 0 Then netErr = Err.Description: Err.Clear
        On Error GoTo 0

        If Len(netErr) = 0 Then
            status = http.Status
            ' WinINet reports some network failures (timeout, DNS, refused) as a 12xxx status.
            If status >= 12000 And status < 13000 Then netErr = "Windows network error " & status
        End If

        If Len(netErr) = 0 Then
            resp = http.responseText
            If status >= 200 And status < 300 Then
                If Len(TrimWs(resp)) = 0 Then resp = "{}"
                HttpRequest = resp
                Exit Function
            End If
            If Not (status = 408 Or status = 429 Or status >= 500) Or attempt = MAX_RETRIES Then
                If status = 401 Or status = 403 Then FailFast target, HttpMessage(status, resp)
                Raise HttpMessage(status, resp)
            End If
            waitMs = RetryAfterMs(http)
        Else
            If attempt = MAX_RETRIES Then FailFast target, "Could not reach " & base & ": " & TrimWs(netErr)
            waitMs = -1
        End If
        If waitMs < 0 Then waitMs = 500 * (2 ^ attempt)
        If waitMs > MAX_WAIT_MS Then waitMs = MAX_WAIT_MS
        Sleep waitMs
    Next
End Function

' Remembers a failure so the remaining cells fail fast, then raises it.
Private Sub FailFast(ByVal target As String, ByVal msg As String)
    mFailMsg = msg
    mFailFor = target
    mFailUntil = Now + TimeSerial(0, 0, FAIL_FAST_SECONDS)
    Raise msg
End Sub

' Returns the server's requested wait in ms, or -1 if it didn't give a usable one.
Private Function RetryAfterMs(ByVal http As Object) As Long
    Dim h As String
    RetryAfterMs = -1
    On Error Resume Next
    h = Trim$(http.getResponseHeader("retry-after-ms"))
    If IsDigits(h) Then RetryAfterMs = CLng(Left$(h, 9)): Exit Function
    h = Trim$(http.getResponseHeader("retry-after"))
    If IsDigits(h) And Len(h) <= 6 Then RetryAfterMs = CLng(h) * 1000
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    IsDigits = Len(s) > 0 And Not s Like "*[!0-9]*"
End Function

' https:// anywhere; http:// only for this PC.
Private Function IsSafeUrl(ByVal u As String) As Boolean
    Dim l As String
    l = LCase$(u)
    IsSafeUrl = l Like "https://?*" Or l Like "http://localhost[:/]*" Or l = "http://localhost" _
             Or l Like "http://127.0.0.1[:/]*" Or l = "http://127.0.0.1"
End Function

Private Function HttpMessage(ByVal status As Long, ByVal resp As String) As String
    Dim prefix As String, detail As String, o As Object
    Select Case status
        Case 400: prefix = "Bad request"
        Case 401: prefix = "Invalid or missing API key - run JevSetApiKey"
        Case 403: prefix = "Access denied for this API key"
        Case 404: prefix = "Not found - check the base URL and model name"
        Case 422: prefix = "Jev rejected the question"
        Case 429: prefix = "Rate limited by Jev - try again shortly"
        Case 529: prefix = "Jev is overloaded - try again shortly"
        Case Else
            If status >= 500 Then prefix = "Jev server error" Else prefix = "Jev request failed"
    End Select

    On Error Resume Next
    Set o = JsonParseObject(resp)
    If Not o Is Nothing Then
        If o.Exists("error") Then
            If IsObject(o("error")) Then detail = CStr(o("error")("message")) Else detail = CStr(o("error"))
        ElseIf o.Exists("message") Then
            detail = CStr(o("message"))
        ElseIf o.Exists("detail") Then
            detail = DetailText(o("detail"))
        End If
    End If
    On Error GoTo 0
    If Len(detail) = 0 Then detail = TrimWs(resp)
    HttpMessage = prefix & " (HTTP " & status & ")" & IIf(Len(detail) > 0, ": " & Left$(detail, 200), "")
End Function

' Reads FastAPI-style error details: a string, { message }, or a list of
' { loc: [...], msg } validation errors, which becomes "loc.path: msg; ...".
Private Function DetailText(ByVal d As Variant) As String
    Dim item As Variant, parts As String, loc As String, part As Variant
    If Not IsObject(d) Then
        If Not IsNull(d) Then DetailText = CStr(d)
        Exit Function
    End If
    If TypeName(d) = "Dictionary" Then
        If d.Exists("message") Then DetailText = DictText(d, "message") Else DetailText = ToJson(d)
        Exit Function
    End If
    For Each item In d
        If IsObject(item) Then
            If TypeName(item) = "Dictionary" Then
                loc = ""
                If item.Exists("loc") Then
                    If IsObject(item("loc")) Then
                        For Each part In item("loc")
                            If Not IsObject(part) Then loc = loc & IIf(Len(loc) > 0, ".", "") & ValueText(part)
                        Next
                    End If
                End If
                If Len(parts) > 0 Then parts = parts & "; "
                parts = parts & IIf(Len(loc) > 0, loc & ": ", "") & DictText(item, "msg") & DictText(item, "message")
            End If
        End If
    Next
    DetailText = parts
End Function

'==============================================================================
'  Output formatting
'==============================================================================

Private Function FormatAnswer(ByVal a As Object, ByVal qType As String, ByVal outKind As String, _
                              ByVal threshold As Double, ByVal levels As Collection) As Variant
    If outKind = "json" Then FormatAnswer = ToJson(a): Exit Function

    Dim probs As Object, k As Variant, n As Long, i As Long
    Dim out() As Variant, keys() As Variant, vals() As Double

    Select Case qType
        Case "choice"
            Select Case outKind
                Case "label": FormatAnswer = DictText(a, "choice", True)
                Case "confidence": FormatAnswer = DictNum(a, "confidence")
                Case "prob": FormatAnswer = DictNum(DictObj(a, "probabilities"), DictText(a, "choice", True))
                Case "probs"
                    Set probs = DictObj(a, "probabilities")
                    n = probs.Count
                    If n = 0 Then Raise "Jev response is missing probabilities."
                    ReDim keys(1 To n): ReDim vals(1 To n)
                    i = 0
                    For Each k In probs.Keys
                        i = i + 1: keys(i) = k: vals(i) = DictNum(probs, CStr(k))
                    Next
                    SortPairs keys, vals, False            ' highest probability first
                    ReDim out(1 To n, 1 To 2)
                    For i = 1 To n
                        out(i, 1) = CStr(keys(i)): out(i, 2) = vals(i)
                    Next
                    FormatAnswer = out
            End Select

        Case "score"
            Dim sc As Double, lvl As Long
            sc = DictNum(a, "score")
            lvl = Int(sc + 0.5)
            If lvl < 0 Then lvl = 0
            If levels.Count > 0 And lvl > levels.Count - 1 Then lvl = levels.Count - 1
            Select Case outKind
                Case "score": FormatAnswer = sc
                Case "level": FormatAnswer = lvl
                Case "label"
                    FormatAnswer = lvl
                    If a.Exists("legend") Then
                        If IsObject(a("legend")) Then
                            If a("legend").Exists(CStr(lvl)) Then
                                If IsObject(a("legend")(CStr(lvl))) Then
                                    FormatAnswer = ToJson(a("legend")(CStr(lvl))): Exit Function
                                ElseIf Not IsNull(a("legend")(CStr(lvl))) Then
                                    FormatAnswer = ValueText(a("legend")(CStr(lvl))): Exit Function
                                End If
                            End If
                        End If
                    End If
                    If lvl + 1 <= levels.Count Then FormatAnswer = levels(lvl + 1)
                Case "confidence": FormatAnswer = DictNum(a, "confidence")
                Case "probs"
                    Set probs = DictObj(a, "probabilities")
                    n = probs.Count
                    If n = 0 Then Raise "Jev response is missing probabilities."
                    ReDim keys(1 To n): ReDim vals(1 To n)
                    i = 0
                    For Each k In probs.Keys
                        i = i + 1: keys(i) = Val(CStr(k)): vals(i) = DictNum(probs, CStr(k))
                    Next
                    SortPairs keys, vals, True             ' lowest level first
                    ReDim out(1 To n, 1 To 2)
                    For i = 1 To n
                        out(i, 1) = keys(i): out(i, 2) = vals(i)
                    Next
                    FormatAnswer = out
            End Select

        Case "noul"
            Dim pr As Double
            pr = DictNum(a, "noul")
            Select Case outKind
                Case "prob": FormatAnswer = pr
                Case "bool": FormatAnswer = (pr >= threshold)
                Case "yesno": FormatAnswer = IIf(pr >= threshold, "Yes", "No")
            End Select
    End Select
End Function

' Insertion sort of parallel arrays (1-based): by key ascending, or by value descending.
Private Sub SortPairs(keys() As Variant, vals() As Double, ByVal byKeyAscending As Boolean)
    Dim i As Long, j As Long, tk As Variant, tv As Double, inOrder As Boolean
    For i = LBound(keys) + 1 To UBound(keys)
        tk = keys(i): tv = vals(i): j = i - 1
        Do While j >= LBound(keys)
            If byKeyAscending Then inOrder = (keys(j) <= tk) Else inOrder = (vals(j) >= tv)
            If inOrder Then Exit Do
            keys(j + 1) = keys(j): vals(j + 1) = vals(j): j = j - 1
        Loop
        keys(j + 1) = tk: vals(j + 1) = tv
    Next
End Sub

'==============================================================================
'  Settings
'==============================================================================

Private Function ApiKey() As String
    ApiKey = Trim$(GetSetting(REG_APP, REG_SECTION, "ApiKey", ""))
    If Len(ApiKey) = 0 Then ApiKey = Trim$(Environ$("TYPESAFE_API_KEY"))
End Function

Private Function ModelName() As String
    ModelName = Trim$(GetSetting(REG_APP, REG_SECTION, "Model", ""))
    If Len(ModelName) = 0 Then ModelName = DEFAULT_MODEL
End Function

Private Function BaseUrl() As String
    BaseUrl = Trim$(GetSetting(REG_APP, REG_SECTION, "BaseUrl", ""))
    If Len(BaseUrl) = 0 Then BaseUrl = DEFAULT_BASE_URL
    Do While Right$(BaseUrl, 1) = "/"
        BaseUrl = Left$(BaseUrl, Len(BaseUrl) - 1)
    Loop
End Function

'==============================================================================
'  Value helpers
'==============================================================================

Private Sub Raise(ByVal msg As String)
    Err.Raise JEV_ERR, "Jev", msg
End Sub

' The current error as cell text; errors this module didn't raise get their number.
Private Function ErrMessage() As String
    If Err.Number = JEV_ERR Then
        ErrMessage = Err.Description
    Else
        ErrMessage = "Unexpected error " & Err.Number & ": " & Err.Description
    End If
End Function

' "?Which team?" is a question; a lone "?" is an ordinary option.
Private Function IsQuestion(ByVal s As String) As Boolean
    IsQuestion = Len(s) > 1 And Left$(s, 1) = "?"
End Function

Private Function ErrorResult(ByVal msg As String) As Variant
    If GetSetting(REG_APP, REG_SECTION, "ErrorsAsText", "1") = "1" Then
        ErrorResult = "#JEV! " & msg
    Else
        ErrorResult = CVErr(xlErrValue)
    End If
End Function

Private Function ScalarOf(v As Variant) As Variant
    If IsObject(v) Then
        If TypeName(v) = "Range" Then ScalarOf = v.Cells(1, 1).Value2: Exit Function
        ScalarOf = CVErr(xlErrValue): Exit Function
    End If
    If IsArray(v) Then
        Dim c As Collection
        Set c = New Collection
        FlattenAll v, c
        If c.Count > 0 Then ScalarOf = c(1) Else ScalarOf = Empty
        Exit Function
    End If
    If IsMissing(v) Then ScalarOf = CVErr(xlErrValue): Exit Function
    ScalarOf = v
End Function

Private Function IsBlankValue(v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then IsBlankValue = True: Exit Function
    If VarType(v) = vbString Then IsBlankValue = (Len(TrimWs(CStr(v))) = 0)
End Function

Private Sub CheckNotError(v As Variant)
    If IsError(v) Then Raise "An input cell contains an error value."
End Sub

Private Function ValueText(v As Variant) As String
    CheckNotError v
    If VarType(v) = vbBoolean Then
        ValueText = IIf(v, "TRUE", "FALSE")
    ElseIf IsNull(v) Or IsEmpty(v) Then
        ValueText = ""
    ElseIf (VarType(v) >= vbInteger And VarType(v) <= vbCurrency) Or VarType(v) = vbDecimal Then
        ValueText = NumJson(CDbl(v))            ' "." decimal point regardless of locale
    Else
        ValueText = CStr(v)
    End If
End Function

Private Function ValueJson(v As Variant) As String
    CheckNotError v
    If IsEmpty(v) Or IsNull(v) Then
        ValueJson = "null"
    ElseIf VarType(v) = vbString And Len(v) = 0 Then
        ValueJson = "null"
    ElseIf VarType(v) = vbBoolean Then
        ValueJson = IIf(v, "true", "false")
    ElseIf VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Or VarType(v) = vbSingle Then
        ValueJson = NumJson(CDbl(v))
    Else
        ValueJson = JStr(CStr(v))
    End If
End Function

' Trims spaces, tabs, CR, LF, VT, FF, no-break spaces and BOMs, like JavaScript's trim
' for the characters that turn up in spreadsheets.
Private Function TrimWs(ByVal s As String) As String
    Dim a As Long, b As Long, ws As String
    ws = " " & vbTab & vbCr & vbLf & vbVerticalTab & vbFormFeed & ChrW$(160) & ChrW$(&HFEFF&)
    a = 1: b = Len(s)
    Do While a <= b
        If InStr(1, ws, Mid$(s, a, 1), vbBinaryCompare) = 0 Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If InStr(1, ws, Mid$(s, b, 1), vbBinaryCompare) = 0 Then Exit Do
        b = b - 1
    Loop
    If b >= a Then TrimWs = Mid$(s, a, b - a + 1)
End Function

' JSON number grammar: -?(0|[1-9][0-9]*)(\.[0-9]+)?([eE][+-]?[0-9]+)?
Private Function IsJsonNumber(ByVal t As String) As Boolean
    Dim p As Long, n As Long, digits As Long
    n = Len(t): p = 1
    If Mid$(t, p, 1) = "-" Then p = p + 1        ' Mid$ past the end returns "", so no bounds check needed
    If p > n Then Exit Function
    If Mid$(t, p, 1) = "0" Then
        p = p + 1
    ElseIf Mid$(t, p, 1) Like "[1-9]" Then
        Do While p <= n
            If Not Mid$(t, p, 1) Like "#" Then Exit Do
            p = p + 1
        Loop
    Else
        Exit Function
    End If
    If p <= n Then
        If Mid$(t, p, 1) = "." Then
            p = p + 1: digits = 0
            Do While p <= n
                If Not Mid$(t, p, 1) Like "#" Then Exit Do
                p = p + 1: digits = digits + 1
            Loop
            If digits = 0 Then Exit Function
        End If
    End If
    If p <= n Then
        If Mid$(t, p, 1) = "e" Or Mid$(t, p, 1) = "E" Then
            p = p + 1
            If Mid$(t, p, 1) = "+" Or Mid$(t, p, 1) = "-" Then p = p + 1
            digits = 0
            Do While p <= n
                If Not Mid$(t, p, 1) Like "#" Then Exit Do
                p = p + 1: digits = digits + 1
            Loop
            If digits = 0 Then Exit Function
        End If
    End If
    IsJsonNumber = (p > n)
End Function

Private Function NumJson(ByVal d As Double) As String
    Dim s As String
    s = Trim$(Str$(d))                       ' Str$ always uses "." as the decimal point
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    NumJson = s
End Function

Private Function ArrayDims(v As Variant) As Long
    Dim d As Long, tmp As Long
    If Not IsArray(v) Then Exit Function
    On Error GoTo Done
    Do
        tmp = UBound(v, d + 1)
        d = d + 1
    Loop
Done:
    ArrayDims = d
End Function

' Flattens a scalar / range / 1-D or 2-D array (row by row) into `out`, keeping blanks.
Private Sub FlattenAll(v As Variant, ByVal out As Collection)
    Dim arr As Variant, r As Long, c As Long
    If IsObject(v) Then
        If TypeName(v) <> "Range" Then Raise "Unsupported argument type " & TypeName(v) & "."
        arr = v.Value2
    Else
        If IsMissing(v) Then Exit Sub
        arr = v
    End If
    Select Case ArrayDims(arr)
        Case 0
            out.Add arr
        Case 1
            For r = LBound(arr) To UBound(arr)
                out.Add arr(r)
            Next
        Case 2
            For r = LBound(arr, 1) To UBound(arr, 1)
                For c = LBound(arr, 2) To UBound(arr, 2)
                    out.Add arr(r, c)
                Next
            Next
        Case Else
            Raise "Arrays with more than two dimensions are not supported."
    End Select
End Sub

' Like FlattenAll but drops blanks, converts to text and rejects error values.
Private Sub FlattenNonBlank(v As Variant, ByVal out As Collection)
    Dim all As Collection, x As Variant
    Set all = New Collection
    FlattenAll v, all
    For Each x In all
        CheckNotError x
        If Not IsBlankValue(x) Then out.Add ValueText(x)
    Next
End Sub

Private Function DictText(ByVal d As Object, ByVal key As String, Optional ByVal required As Boolean = False) As String
    If d.Exists(key) Then
        If Not IsNull(d(key)) Then DictText = CStr(d(key)): Exit Function
    End If
    If required Then Raise "Jev response is missing " & key & "."
End Function

Private Function DictObj(ByVal d As Object, ByVal key As String) As Object
    If d.Exists(key) Then
        If IsObject(d(key)) Then Set DictObj = d(key): Exit Function
    End If
    Raise "Jev response is missing " & key & "."
End Function

Private Function DictNum(ByVal d As Object, ByVal key As String) As Double
    If Not d.Exists(key) Then Raise "Jev response is missing " & key & "."
    If IsNull(d(key)) Or IsObject(d(key)) Then Raise "Jev response is missing " & key & "."
    DictNum = CDbl(d(key))
End Function

'==============================================================================
'  JSON: string escaping, serialising and a small parser
'  Objects -> Scripting.Dictionary, arrays -> Collection, numbers -> Double.
'==============================================================================

Private Function JStr(ByVal s As String) As String
    Dim n As Long, buf As String, p As Long, i As Long, code As Long, esc As String
    n = Len(s)
    buf = Space$(n * 6 + 2)
    p = 1
    Mid$(buf, p, 1) = """": p = p + 1
    For i = 1 To n
        code = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case code
            Case 34: esc = "\"""
            Case 92: esc = "\\"
            Case 8: esc = "\b"
            Case 9: esc = "\t"
            Case 10: esc = "\n"
            Case 12: esc = "\f"
            Case 13: esc = "\r"
            Case 32 To 126: esc = ChrW$(code)
            Case Else: esc = "\u" & Right$("000" & Hex$(code), 4)
        End Select
        Mid$(buf, p, Len(esc)) = esc
        p = p + Len(esc)
    Next
    Mid$(buf, p, 1) = """"
    JStr = Left$(buf, p)
End Function

Private Function ToJson(v As Variant) As String
    Dim parts() As String, i As Long, k As Variant, item As Variant
    If IsObject(v) Then
        If TypeName(v) = "Dictionary" Then
            If v.Count = 0 Then ToJson = "{}": Exit Function
            ReDim parts(0 To v.Count - 1)
            For Each k In v.Keys
                parts(i) = JStr(CStr(k)) & ":" & ToJson(v(k))
                i = i + 1
            Next
            ToJson = "{" & Join(parts, ",") & "}"
        ElseIf TypeName(v) = "Collection" Then
            If v.Count = 0 Then ToJson = "[]": Exit Function
            ReDim parts(0 To v.Count - 1)
            For Each item In v
                parts(i) = ToJson(item)
                i = i + 1
            Next
            ToJson = "[" & Join(parts, ",") & "]"
        Else
            ToJson = "null"
        End If
    ElseIf IsNull(v) Or IsEmpty(v) Then
        ToJson = "null"
    ElseIf VarType(v) = vbBoolean Then
        ToJson = IIf(v, "true", "false")
    ElseIf VarType(v) = vbString Then
        ToJson = JStr(CStr(v))
    ElseIf IsNumeric(v) Then
        ToJson = NumJson(CDbl(v))
    Else
        ToJson = JStr(CStr(v))
    End If
End Function

Private Function JsonParseObject(ByVal s As String) As Object
    mJson = s
    mPos = 1
    SkipWs
    If Mid$(mJson, mPos, 1) <> "{" Then Raise "Jev returned a response that is not a JSON object."
    Set JsonParseObject = PObject()
End Function

Private Function IsValidJson(ByVal s As String) As Boolean
    Dim holder As Collection
    Set holder = New Collection
    On Error GoTo Bad
    mJson = s
    mPos = 1
    holder.Add PValue()
    SkipWs
    IsValidJson = (mPos > Len(mJson))
    Exit Function
Bad:
    IsValidJson = False
End Function

Private Sub SkipWs()
    Dim ch As String
    Do While mPos <= Len(mJson)
        ch = Mid$(mJson, mPos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        mPos = mPos + 1
    Loop
End Sub

Private Function PValue() As Variant
    SkipWs
    If mPos > Len(mJson) Then Raise "Invalid JSON (unexpected end)."
    Select Case Mid$(mJson, mPos, 1)
        Case "{": Set PValue = PObject()
        Case "[": Set PValue = PArray()
        Case """": PValue = PString()
        Case "t": PWord "true": PValue = True
        Case "f": PWord "false": PValue = False
        Case "n": PWord "null": PValue = Null
        Case Else: PValue = PNumber()
    End Select
End Function

Private Sub PWord(ByVal w As String)
    If Mid$(mJson, mPos, Len(w)) <> w Then Raise "Invalid JSON near position " & mPos & "."
    mPos = mPos + Len(w)
End Sub

Private Function PObject() As Object
    Dim d As Object, k As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0
    mPos = mPos + 1
    SkipWs
    If Mid$(mJson, mPos, 1) = "}" Then mPos = mPos + 1: Set PObject = d: Exit Function
    Do
        SkipWs
        If Mid$(mJson, mPos, 1) <> """" Then Raise "Invalid JSON (expected a key) near position " & mPos & "."
        k = PString()
        SkipWs
        If Mid$(mJson, mPos, 1) <> ":" Then Raise "Invalid JSON (expected :) near position " & mPos & "."
        mPos = mPos + 1
        If d.Exists(k) Then d.Remove k
        d.Add k, PValue()
        SkipWs
        Select Case Mid$(mJson, mPos, 1)
            Case ",": mPos = mPos + 1
            Case "}": mPos = mPos + 1: Exit Do
            Case Else: Raise "Invalid JSON (expected , or }) near position " & mPos & "."
        End Select
    Loop
    Set PObject = d
End Function

Private Function PArray() As Object
    Dim c As Collection
    Set c = New Collection
    mPos = mPos + 1
    SkipWs
    If Mid$(mJson, mPos, 1) = "]" Then mPos = mPos + 1: Set PArray = c: Exit Function
    Do
        c.Add PValue()
        SkipWs
        Select Case Mid$(mJson, mPos, 1)
            Case ",": mPos = mPos + 1
            Case "]": mPos = mPos + 1: Exit Do
            Case Else: Raise "Invalid JSON (expected , or ]) near position " & mPos & "."
        End Select
    Loop
    Set PArray = c
End Function

Private Function PString() As String
    Dim buf As String, startPos As Long, ch As String, e As String
    mPos = mPos + 1                         ' opening quote
    startPos = mPos
    Do
        If mPos > Len(mJson) Then Raise "Invalid JSON (unterminated string)."
        ch = Mid$(mJson, mPos, 1)
        If AscW(ch) >= 0 And AscW(ch) < 32 Then Raise "Invalid JSON (control character in string)."
        If ch = """" Then
            buf = buf & Mid$(mJson, startPos, mPos - startPos)
            mPos = mPos + 1
            Exit Do
        ElseIf ch = "\" Then
            buf = buf & Mid$(mJson, startPos, mPos - startPos)
            e = Mid$(mJson, mPos + 1, 1)
            Select Case e
                Case """", "\", "/": buf = buf & e: mPos = mPos + 2
                Case "b": buf = buf & ChrW$(8): mPos = mPos + 2
                Case "f": buf = buf & ChrW$(12): mPos = mPos + 2
                Case "n": buf = buf & vbLf: mPos = mPos + 2
                Case "r": buf = buf & vbCr: mPos = mPos + 2
                Case "t": buf = buf & vbTab: mPos = mPos + 2
                Case "u"
                    If Not Mid$(mJson, mPos + 2, 4) Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
                        Raise "Invalid JSON escape near position " & mPos & "."
                    End If
                    buf = buf & ChrW$(CLng("&H" & Mid$(mJson, mPos + 2, 4)))
                    mPos = mPos + 6
                Case Else: Raise "Invalid JSON escape near position " & mPos & "."
            End Select
            startPos = mPos
        Else
            mPos = mPos + 1
        End If
    Loop
    PString = buf
End Function

Private Function PNumber() As Double
    Dim startPos As Long, ch As String
    startPos = mPos
    Do While mPos <= Len(mJson)
        ch = Mid$(mJson, mPos, 1)
        If InStr(1, "+-0123456789.eE", ch, vbBinaryCompare) = 0 Then Exit Do
        mPos = mPos + 1
    Loop
    If Not IsJsonNumber(Mid$(mJson, startPos, mPos - startPos)) Then Raise "Invalid JSON number near position " & startPos & "."
    PNumber = Val(Mid$(mJson, startPos, mPos - startPos))   ' Val always uses "." as the decimal point
End Function
