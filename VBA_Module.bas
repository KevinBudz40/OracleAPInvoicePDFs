Attribute VB_Name = "modAPInvoicePDFs"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════════════
'  AP Invoice PDF Downloader
'  Oracle Fusion Cloud — SCANNED_INVOICE_IMAGE attachments
'
'  Two entry points:
'    GetInvoiceListFromBIP  — populates Invoice_List via BIP REST report
'    DownloadInvoicePDFs    — downloads PDFs for every row in Invoice_List
' ═══════════════════════════════════════════════════════════════════════════════

' ── Sheet / cell constants ───────────────────────────────────────────────────
Private Const CFG     As String = "Config"
Private Const INV_LST As String = "Invoice_List"
Private Const INV_LOG As String = "Invoice_Log"

Private Const C_HOST  As String = "C4"
Private Const C_USER  As String = "C5"
Private Const C_PASS  As String = "C6"
Private Const C_PROJ  As String = "C9"
Private Const C_FROM  As String = "C10"
Private Const C_TO    As String = "C11"
Private Const C_FOLD  As String = "C14"
Private Const C_BIP   As String = "C17"

' ── Colour constants ─────────────────────────────────────────────────────────
Private Const CLR_GREEN  As Long = 13561798  ' RGB(198,239,206)
Private Const CLR_RED    As Long = 13375687  ' RGB(255,199,206)
Private Const CLR_YELLOW As Long = 10079232  ' RGB(255,235,156)
Private Const CLR_WHITE  As Long = 16777215

' ═══════════════════════════════════════════════════════════════════════════════
'  PUBLIC ENTRY POINT 1 — populate Invoice_List from BIP
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub GetInvoiceListFromBIP()
    Dim cfg   As Object
    If Not ReadConfig(cfg) Then Exit Sub

    Dim bipPath As String
    bipPath = Trim(ThisWorkbook.Sheets(CFG).Range(C_BIP).Value)
    If bipPath = "" Then
        MsgBox "No BIP report path is configured (Config row 17)." & vbLf & _
               "Please enter the path or use Manual mode (paste into Invoice_List).", _
               vbExclamation, "BIP Path Missing"
        Exit Sub
    End If

    Application.StatusBar = "Calling BIP report..."
    Dim csv As String
    csv = BipReportCsv(cfg("Host"), cfg("User"), cfg("Pass"), bipPath, _
                       cfg("Project"), cfg("FromDate"), cfg("ToDate"))

    If csv = "" Then
        MsgBox "BIP report returned no data or could not be reached." & vbLf & _
               "Check the BIP path and try Manual mode as a fallback.", _
               vbExclamation, "BIP Error"
        Application.StatusBar = False
        Exit Sub
    End If

    ' Write rows to Invoice_List (starting row 3, clear old data first)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(INV_LST)
    ws.Range("A3:C" & ws.Rows.Count).ClearContents

    Dim lines() As String
    lines = Split(csv, vbLf)
    Dim r As Long : r = 3
    Dim i As Long
    For i = 1 To UBound(lines)   ' skip header row (i=0)
        Dim cols() As String
        cols = Split(lines(i), ",")
        If UBound(cols) >= 0 And Trim(cols(0)) <> "" Then
            ws.Cells(r, 1).Value = Trim(cols(0))  ' Invoice Number
            ws.Cells(r, 2).Value = IIf(UBound(cols) >= 1, Trim(cols(1)), "")
            ws.Cells(r, 3).Value = IIf(UBound(cols) >= 2, Trim(cols(2)), "")
            r = r + 1
        End If
    Next i

    Application.StatusBar = False
    MsgBox (r - 3) & " invoice(s) loaded into the Invoice_List sheet.", _
           vbInformation, "BIP Complete"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
'  PUBLIC ENTRY POINT 2 — download PDFs for every invoice in Invoice_List
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub DownloadInvoicePDFs()
    Dim cfg As Object
    If Not ReadConfig(cfg) Then Exit Sub

    ' Make sure save folder exists
    If Not EnsureFolder(cfg("Folder")) Then
        MsgBox "Could not create save folder: " & cfg("Folder"), _
               vbCritical, "Folder Error"
        Exit Sub
    End If

    ' Read Invoice_List
    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets(INV_LST)
    Dim lastRow As Long
    lastRow = wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "Invoice_List is empty. Run  Get Invoice List from BIP  " & _
               "or paste invoice numbers manually (starting at row 3).", _
               vbExclamation, "No Invoices"
        Exit Sub
    End If

    InitLogSheet

    Dim logRow As Long : logRow = 3
    Dim r As Long
    For r = 3 To lastRow
        Dim invNum As String
        invNum = Trim(wsInv.Cells(r, 1).Value)
        If invNum = "" Then GoTo NextRow

        Application.StatusBar = "Processing invoice " & invNum & _
                                 "  (" & (r - 2) & " of " & (lastRow - 2) & ")"

        ' Step A — resolve InvoiceNumber → InvoiceId
        Dim invId As String
        invId = ResolveInvoiceId(cfg, invNum)
        If invId = "" Then
            LogRow logRow, invNum, "", "", "Invoice not found via REST", ""
            logRow = logRow + 1
            GoTo NextRow
        End If

        ' Step B — fetch SCANNED_INVOICE_IMAGE attachments
        Dim attJson As String
        Dim attUrl As String
        attUrl = cfg("Host") & "/fscmRestApi/resources/11.13.18.05/invoices/" & _
                 invId & "/child/attachments" & _
                 "?q=CategoryName%3D%27SCANNED_INVOICE_IMAGE%27"
        attJson = HttpGet(attUrl, cfg("User"), cfg("Pass"))

        If attJson = "" Then
            LogRow logRow, invNum, invId, "", "REST call failed for attachments", ""
            logRow = logRow + 1
            GoTo NextRow
        End If

        ' Step C — parse items and download each file
        Dim items() As String
        Dim itemCount As Long
        itemCount = ParseJsonItems(attJson, items)

        If itemCount = 0 Then
            LogRow logRow, invNum, invId, "", "No SCANNED_INVOICE_IMAGE attachments", ""
            logRow = logRow + 1
            GoTo NextRow
        End If

        Dim j As Long
        For j = 0 To itemCount - 1
            Dim fName As String
            Dim fUrl  As String
            fName = JsonValue(items(j), "FileName")
            fUrl  = FileContentsHref(items(j))

            If fUrl = "" Then
                LogRow logRow, invNum, invId, fName, "No FileContents link", ""
            Else
                Dim savePath As String
                savePath = cfg("Folder") & SafeName(fName, invNum & "_att" & j & ".pdf")
                Dim bytes As Long
                bytes = DownloadBinary(fUrl, cfg("User"), cfg("Pass"), savePath)
                If bytes > 0 Then
                    LogRow logRow, invNum, invId, fName, _
                           "Saved  " & Format(bytes, "#,##0") & " bytes", savePath, True
                Else
                    LogRow logRow, invNum, invId, fName, "Download failed", "", False
                End If
            End If
            logRow = logRow + 1
        Next j

NextRow:
        DoEvents
    Next r

    Application.StatusBar = False
    MsgBox "Done!  " & (logRow - 3) & " row(s) written to Invoice_Log.", _
           vbInformation, "Download Complete"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
'  BIP REST — run report, return CSV text
'
'  Expects a BIP report with parameters:  P_PROJECT, P_FROM_DATE, P_TO_DATE
'  Returns CSV output (outputFormat=csv).
' ═══════════════════════════════════════════════════════════════════════════════
Private Function BipReportCsv(host As String, user As String, pass As String, _
                               reportPath As String, _
                               proj As String, fromDt As String, toDt As String) _
                               As String
    ' URL-encode the report path (replace / with %2F is NOT needed for xmlpserver)
    Dim url As String
    url = host & "/xmlpserver/services/rest/v1/reports" & _
          "?reportAbsolutePath=" & reportPath & _
          "&P_PROJECT=" & EncodeUrl(proj) & _
          "&P_FROM_DATE=" & EncodeUrl(fromDt) & _
          "&P_TO_DATE=" & EncodeUrl(toDt) & _
          "&_output=csv"

    Dim result As String
    result = HttpGet(url, user, pass)

    ' xmlpserver may wrap CSV in a JSON envelope; try to extract raw CSV
    If Left(result, 1) = "{" Then
        ' {"reportBytes":"base64..."} — decode the base64
        Dim b64 As String
        b64 = JsonValue(result, "reportBytes")
        If b64 <> "" Then
            result = Base64ToString(b64)
        End If
    End If

    BipReportCsv = result
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  ORACLE REST HELPERS
' ═══════════════════════════════════════════════════════════════════════════════
Private Function ResolveInvoiceId(cfg As Object, invNum As String) As String
    Dim url As String
    url = cfg("Host") & "/fscmRestApi/resources/11.13.18.05/invoices" & _
          "?q=InvoiceNumber%3D%27" & EncodeUrl(invNum) & "%27" & _
          "&fields=InvoiceId%2CInvoiceNumber"
    Dim json As String
    json = HttpGet(url, cfg("User"), cfg("Pass"))
    If json = "" Then Exit Function

    Dim items() As String
    If ParseJsonItems(json, items) > 0 Then
        ResolveInvoiceId = JsonValue(items(0), "InvoiceId")
    End If
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  HTTP
' ═══════════════════════════════════════════════════════════════════════════════
Private Function HttpGet(url As String, user As String, pass As String) As String
    On Error GoTo ErrH
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Basic " & B64Encode(user & ":" & pass)
    http.SetRequestHeader "Accept", "application/json"
    http.Send
    If http.Status = 200 Then
        HttpGet = http.ResponseText
    Else
        Debug.Print "HTTP " & http.Status & " ← " & Left(url, 120)
        Debug.Print http.ResponseText
    End If
    Exit Function
ErrH:
    Debug.Print "HttpGet error: " & Err.Description
End Function

Private Function DownloadBinary(url As String, user As String, pass As String, _
                                 savePath As String) As Long
    On Error GoTo ErrH
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Basic " & B64Encode(user & ":" & pass)
    http.SetRequestHeader "Accept", "*/*"
    http.Send
    If http.Status <> 200 Then
        Debug.Print "DownloadBinary HTTP " & http.Status
        Exit Function
    End If
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1  ' adTypeBinary
    st.Open
    st.Write http.ResponseBody
    st.SaveToFile savePath, 2  ' adSaveCreateOverWrite
    st.Close
    DownloadBinary = FileLen(savePath)
    Exit Function
ErrH:
    Debug.Print "DownloadBinary error: " & Err.Description
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  JSON HELPERS  (no external library — handles Oracle FSCM response shapes)
' ═══════════════════════════════════════════════════════════════════════════════
' Split top-level "items":[{...},{...}] into individual JSON object strings
Private Function ParseJsonItems(json As String, ByRef items() As String) As Long
    ReDim items(0 To 500)
    Dim arrStart As Long
    arrStart = InStr(json, """items"":[")
    If arrStart = 0 Then Exit Function

    Dim pos As Long
    pos = InStr(arrStart, json, "[") + 1
    Dim depth As Long : depth = 0
    Dim objStart As Long : objStart = 0
    Dim count As Long : count = 0

    Do While pos <= Len(json)
        Dim ch As String
        ch = Mid(json, pos, 1)
        Select Case ch
            Case "{"
                If depth = 0 Then objStart = pos
                depth = depth + 1
            Case "}"
                depth = depth - 1
                If depth = 0 And objStart > 0 Then
                    items(count) = Mid(json, objStart, pos - objStart + 1)
                    count = count + 1
                    objStart = 0
                End If
            Case "]"
                If depth = 0 Then Exit Do
        End Select
        pos = pos + 1
    Loop
    ParseJsonItems = count
End Function

' Extract a scalar value (string, number, null) by key from a JSON object string
Private Function JsonValue(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """:"
    Dim sp As Long
    sp = InStr(json, pattern)
    If sp = 0 Then Exit Function
    sp = sp + Len(pattern)

    ' skip whitespace
    Do While sp <= Len(json) And Mid(json, sp, 1) = " " : sp = sp + 1 : Loop

    Dim ch As String
    ch = Mid(json, sp, 1)
    If ch = """" Then
        sp = sp + 1
        Dim ep As Long
        ep = sp
        Do While ep <= Len(json)
            If Mid(json, ep, 1) = """" And Mid(json, ep - 1, 1) <> "\" Then Exit Do
            ep = ep + 1
        Loop
        JsonValue = Mid(json, sp, ep - sp)
    ElseIf Mid(json, sp, 4) = "null" Then
        JsonValue = ""
    Else
        Dim e1 As Long, e2 As Long
        e1 = InStr(sp, json, ",") : If e1 = 0 Then e1 = Len(json)
        e2 = InStr(sp, json, "}") : If e2 = 0 Then e2 = Len(json)
        Dim endPos As Long
        endPos = IIf(e1 < e2, e1, e2)
        JsonValue = Trim(Mid(json, sp, endPos - sp))
    End If
End Function

' Find the href of the FileContents link inside an attachment JSON item
Private Function FileContentsHref(itemJson As String) As String
    Dim pos As Long : pos = 1
    Do
        Dim np As Long
        np = InStr(pos, itemJson, """name"":""FileContents""")
        If np = 0 Then Exit Do
        ' Look back in the same link object for the href
        Dim hp As Long
        hp = InStrRev(itemJson, """href"":""", np)
        If hp > 0 Then
            Dim hs As Long : hs = hp + 8
            Dim he As Long : he = InStr(hs, itemJson, """")
            FileContentsHref = Mid(itemJson, hs, he - hs)
            Exit Function
        End If
        pos = np + 1
    Loop
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  BASE-64
' ═══════════════════════════════════════════════════════════════════════════════
Private Function B64Encode(s As String) As String
    Dim xml  As Object : Set xml  = CreateObject("MSXML2.DOMDocument")
    Dim node As Object : Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = StringToBytes(s)
    B64Encode = Replace(node.Text, vbLf, "")
End Function

Private Function Base64ToString(b64 As String) As String
    Dim xml  As Object : Set xml  = CreateObject("MSXML2.DOMDocument")
    Dim node As Object : Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = b64
    Dim st As Object : Set st = CreateObject("ADODB.Stream")
    st.Type = 1 : st.Open : st.Write node.nodeTypedValue
    st.Position = 0 : st.Type = 2 : st.CharSet = "utf-8"
    Base64ToString = st.ReadText : st.Close
End Function

Private Function StringToBytes(s As String) As Variant
    Dim st As Object : Set st = CreateObject("ADODB.Stream")
    st.Type = 2 : st.CharSet = "us-ascii" : st.Open : st.WriteText s
    st.Position = 0 : st.Type = 1
    StringToBytes = st.Read : st.Close
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  UTILITY
' ═══════════════════════════════════════════════════════════════════════════════
Private Function EncodeUrl(s As String) As String
    ' Minimal URL encoding for query-string values
    s = Replace(s, " ", "%20")
    s = Replace(s, "'", "%27")
    s = Replace(s, """", "%22")
    s = Replace(s, "&", "%26")
    EncodeUrl = s
End Function

Private Function SafeName(name As String, default_ As String) As String
    If Trim(name) = "" Then
        SafeName = default_
        Exit Function
    End If
    Dim s As String : s = name
    Dim bad As Variant
    For Each bad In Array("/", "\", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, bad, "_")
    Next bad
    SafeName = s
End Function

Private Function EnsureFolder(path As String) As Boolean
    On Error GoTo ErrH
    If Right(path, 1) <> "\" Then path = path & "\"
    If Dir(path, vbDirectory) = "" Then MkDir path
    EnsureFolder = True
    Exit Function
ErrH:
    EnsureFolder = False
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  CONFIG READER
' ═══════════════════════════════════════════════════════════════════════════════
Private Function ReadConfig(ByRef cfg As Object) As Boolean
    Set cfg = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CFG)

    cfg("Host")    = Trim(ws.Range(C_HOST).Value)
    cfg("User")    = Trim(ws.Range(C_USER).Value)
    cfg("Pass")    = Trim(ws.Range(C_PASS).Value)
    cfg("Project") = Trim(ws.Range(C_PROJ).Value)
    cfg("FromDate")= Format(ws.Range(C_FROM).Value, "YYYY-MM-DD")
    cfg("ToDate")  = Format(ws.Range(C_TO).Value, "YYYY-MM-DD")
    cfg("Folder")  = Trim(ws.Range(C_FOLD).Value)

    If cfg("Host") = "" Or cfg("User") = "" Or cfg("Pass") = "" Then
        MsgBox "Please fill in Host URL, Username, and Password on the Config sheet.", _
               vbExclamation, "Config Incomplete"
        Exit Function
    End If
    If cfg("Project") = "" Then
        MsgBox "Please enter a Project Number on the Config sheet.", _
               vbExclamation, "Config Incomplete"
        Exit Function
    End If
    If Right(cfg("Folder"), 1) <> "\" Then cfg("Folder") = cfg("Folder") & "\"
    ReadConfig = True
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  LOG SHEET
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub InitLogSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(INV_LOG)
    ws.Range("A3:G" & ws.Rows.Count).ClearContents
    ws.Range("A3:G" & ws.Rows.Count).Interior.ColorIndex = xlNone
End Sub

Private Sub LogRow(rowNum As Long, invNum As String, invId As String, _
                   fName As String, status As String, savedPath As String, _
                   Optional success As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(INV_LOG)

    ws.Cells(rowNum, 1).Value = invNum
    ws.Cells(rowNum, 2).Value = invId
    ws.Cells(rowNum, 3).Value = fName
    ws.Cells(rowNum, 4).Value = status
    ws.Cells(rowNum, 5).Value = savedPath
    ws.Cells(rowNum, 7).Value = Now()

    ' Colour the status cell
    Dim clr As Long
    If IsMissing(success) Then
        clr = CLR_YELLOW
    ElseIf CBool(success) = True Then
        clr = CLR_GREEN
    Else
        clr = CLR_RED
    End If
    ws.Cells(rowNum, 4).Interior.Color = clr
    DoEvents
End Sub
