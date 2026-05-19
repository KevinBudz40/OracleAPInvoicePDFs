Attribute VB_Name = "Module1"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════════════
'  AP Invoice PDF Downloader
'  Oracle Fusion Cloud — Payables
'
'  Oracle Fusion Cloud REST API — Payables invoices + attachments endpoints.
'  CommandButton1 on the Parameters sheet calls DownloadInvoicePDFs.
'
'  Workbook sheets:
'    Parameters   — host URL, username, date range, save folder, invoice list
'    Invoice_Log  — auto-populated download results
' ═══════════════════════════════════════════════════════════════════════════════

' ── Sheet names ──────────────────────────────────────────────────────────────
Private Const SH_PARAMS As String = "Parameters"
Private Const SH_LOG    As String = "Invoice_Log"

' ── Parameter cell addresses on the Parameters sheet ─────────────────────────
Private Const P_HOST   As String = "B3"   ' Oracle Cloud host URL
Private Const P_USER   As String = "B4"   ' Username
' Password is prompted at runtime (masked) — not stored on the sheet
Private Const P_FROM   As String = "B9"   ' From Date (YYYY-MM-DD)
Private Const P_TO     As String = "B10"  ' To Date   (YYYY-MM-DD)
Private Const P_FOLDER As String = "B13"  ' PDF save folder
' Invoice numbers start at row 17, column A (below a small header at row 16)

' ── Colours ──────────────────────────────────────────────────────────────────
Private Const CLR_GREEN  As Long = 13561798   ' RGB(198,239,206)
Private Const CLR_RED    As Long = 13375687   ' RGB(255,199,206)
Private Const CLR_YELLOW As Long = 10079232   ' RGB(255,235,156)

' ═══════════════════════════════════════════════════════════════════════════════
'  ENTRY POINT  —  called by CommandButton1_Click on Sheet1
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub DownloadInvoicePDFs()

    ' ── read and validate config ──────────────────────────────────────────────
    Dim host   As String, user  As String, pass_  As String
    Dim fromD  As String, toD   As String
    Dim folder As String

    With ThisWorkbook.Sheets(SH_PARAMS)
        host   = Trim(.Range(P_HOST).Value)
        user   = Trim(.Range(P_USER).Value)
        fromD  = Format(.Range(P_FROM).Value, "YYYY-MM-DD")
        toD    = Format(.Range(P_TO).Value,   "YYYY-MM-DD")
        folder = Trim(.Range(P_FOLDER).Value)
    End With

    If host = "" Or user = "" Then
        MsgBox "Please fill in Host URL and Username on the Parameters sheet.", _
               vbExclamation, "Configuration Incomplete"
        Exit Sub
    End If

    ' ── prompt for password (masked) — not stored anywhere on the sheet ───────
    pass_ = PasswordForm.GetPassword()
    If PasswordForm.Cancelled Or pass_ = "" Then
        MsgBox "No password entered — download cancelled.", _
               vbExclamation, "Cancelled"
        Exit Sub
    End If
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Dir(folder, vbDirectory) = "" Then MkDir folder

    ' ── read invoice list (col A, rows 17+) ───────────────────────────────────
    Dim wsPar As Worksheet
    Set wsPar = ThisWorkbook.Sheets(SH_PARAMS)
    Dim lastRow As Long
    lastRow = wsPar.Cells(wsPar.Rows.Count, 1).End(xlUp).Row

    If lastRow < 17 Then
        MsgBox "No invoice numbers found." & vbLf & _
               "Enter invoice numbers in column A starting at row 17 of the Parameters sheet.", _
               vbExclamation, "No Invoices"
        Exit Sub
    End If

    ' ── initialise log sheet ──────────────────────────────────────────────────
    InitLog

    Dim logRow As Long : logRow = 3
    Dim r As Long

    For r = 17 To lastRow
        Dim invNum As String
        invNum = Trim(wsPar.Cells(r, 1).Value)
        If invNum = "" Then GoTo NextInv

        Application.StatusBar = "Invoice " & invNum & "  (" & (r - 16) & " of " & (lastRow - 16) & ")"

        ' ── resolve InvoiceNumber → InvoiceId + metadata ─────────────────────
        Dim invId As String, suppName As String, invAmt As String
        invId = ResolveInvoiceId(host, user, pass_, invNum, suppName, invAmt)
        If invId = "" Then
            WriteLog logRow, invNum, "", "", "Invoice not found", "", False
            logRow = logRow + 1
            GoTo NextInv
        End If

        ' ── fetch all attachments (CategoryName is not a filterable field) ──────
        Dim attUrl As String
        attUrl = host & "/fscmRestApi/resources/11.13.18.05/invoices/" & invId & _
                 "/child/attachments"
        Dim attJson As String
        attJson = HttpGet(attUrl, user, pass_)

        If attJson = "" Then
            WriteLog logRow, invNum, invId, "", "Attachments call failed", "", False
            logRow = logRow + 1
            GoTo NextInv
        End If

        Dim items() As String
        Dim nItems  As Long
        nItems = ParseItems(attJson, items)

        If nItems = 0 Then
            WriteLog logRow, invNum, invId, "", "No attachments found", "", Null
            logRow = logRow + 1
            GoTo NextInv
        End If

        ' ── download each Scanned Invoice Image attachment ────────────────────
        Dim j As Long
        Dim nDownloaded As Long : nDownloaded = 0
        For j = 0 To nItems - 1
            Dim fName As String : fName = JVal(items(j), "FileName")
            Dim fCat  As String : fCat  = JVal(items(j), "Category")

            ' Skip anything that isn't the scanned invoice image
            If LCase(Trim(fCat)) <> "scanned invoice image" Then GoTo NextAtt

            Dim fHref As String : fHref = FileContentsHref(items(j))

            If fHref = "" Then
                WriteLog logRow, invNum, invId, fName, "No FileContents link", "", False
            Else
                Dim savePath As String
                savePath = folder & InvoiceFileName(suppName, invNum, invAmt, nDownloaded)
                Dim bytes As Long
                bytes = SaveBinary(fHref, user, pass_, savePath)
                If bytes > 0 Then
                    WriteLog logRow, invNum, invId, fName, _
                             "Saved  " & Format(bytes, "#,##0") & " bytes", savePath, True
                Else
                    WriteLog logRow, invNum, invId, fName, "Download failed", "", False
                End If
            End If
            nDownloaded = nDownloaded + 1
            logRow = logRow + 1
NextAtt:
        Next j

        If nDownloaded = 0 Then
            WriteLog logRow, invNum, invId, "", "No Scanned Invoice Image attachments", "", Null
            logRow = logRow + 1
        End If

NextInv:
        DoEvents
    Next r

    Application.StatusBar = False
    ThisWorkbook.Sheets(SH_LOG).Activate
    MsgBox "Done.  " & (logRow - 3) & " result row(s) written to Invoice_Log.", _
           vbInformation, "AP Invoice PDF Downloader"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
'  ORACLE REST HELPERS
' ═══════════════════════════════════════════════════════════════════════════════
Private Function ResolveInvoiceId(host As String, user As String, pass_ As String, _
                                   invNum As String, _
                                   ByRef suppName As String, _
                                   ByRef invAmt As String) As String
    ' No fields= restriction — requesting an unrecognised field name causes
    ' Oracle to 400 the whole call, breaking invoice lookup entirely.
    ' Fetch all fields and parse what we need; unknown names just return "".
    Dim url As String
    url = host & "/fscmRestApi/resources/11.13.18.05/invoices" & _
          "?q=InvoiceNumber%3D%27" & invNum & "%27&limit=1"
    Dim json As String : json = HttpGet(url, user, pass_)
    If json = "" Then Exit Function

    Dim items() As String
    If ParseItems(json, items) > 0 Then
        ResolveInvoiceId = JVal(items(0), "InvoiceId")
        suppName = JVal(items(0), "Supplier")
        invAmt = JVal(items(0), "InvoiceAmount")
    End If
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  HTTP
' ═══════════════════════════════════════════════════════════════════════════════
Private Function HttpGet(url As String, user As String, pass_ As String) As String
    ' Uses MSXML2.XMLHTTP which honours the Windows/IE proxy settings
    ' automatically — required on corporate networks where WinHttp's
    ' separate proxy config isn't set up.
    On Error GoTo ErrH
    Dim h As Object
    Set h = CreateObject("MSXML2.XMLHTTP.6.0")
    h.Open "GET", url, False
    h.setRequestHeader "Authorization", "Basic " & B64(user & ":" & pass_)
    h.setRequestHeader "Accept", "application/json"
    h.Send
    If h.Status = 200 Then HttpGet = h.responseText Else _
        Debug.Print "HTTP " & h.Status & " " & Left(url, 100)
    Exit Function
ErrH: Debug.Print "HttpGet: " & Err.Description
End Function

Private Function SaveBinary(url As String, user As String, pass_ As String, _
                             path As String) As Long
    On Error GoTo ErrH
    Dim h As Object
    Set h = CreateObject("MSXML2.XMLHTTP.6.0")
    h.Open "GET", url, False
    h.setRequestHeader "Authorization", "Basic " & B64(user & ":" & pass_)
    h.setRequestHeader "Accept", "*/*"
    h.Send
    If h.Status <> 200 Then
        Debug.Print "SaveBinary HTTP " & h.Status
        Exit Function
    End If
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1          ' adTypeBinary
    st.Open
    st.Write h.responseBody   ' Byte array from MSXML2
    st.SaveToFile path, 2     ' adSaveCreateOverWrite
    st.Close
    SaveBinary = FileLen(path)
    Exit Function
ErrH: Debug.Print "SaveBinary: " & Err.Description
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  JSON HELPERS
' ═══════════════════════════════════════════════════════════════════════════════
Private Function ParseItems(json As String, ByRef items() As String) As Long
    ' Extracts each top-level object from an Oracle "items":[{...},{...}] array.
    '
    ' Two earlier bugs fixed here:
    '   1. Braces inside JSON string VALUES (e.g. the DownloadInfo field whose
    '      value is itself a JSON blob) were being counted, causing premature
    '      extraction before fields like "Category" were reached.
    '   2. The depth counter was only incremented when depth=0, so nested
    '      objects never increased depth beyond 1.
    ' Fix: track whether the scanner is inside a quoted string and skip
    ' all structural characters ({, }, [, ]) while inside one.
    ReDim items(0 To 500)
    Dim p As Long : p = InStr(json, """items""")
    If p = 0 Then Exit Function
    p = InStr(p, json, "[") + 1

    Dim depth  As Long
    Dim oStart As Long
    Dim count  As Long
    Dim insideStr As Boolean : insideStr = False

    Do While p <= Len(json)
        Dim c As String : c = Mid(json, p, 1)

        If insideStr Then
            ' Inside a string value — only care about escape and closing quote
            If c = "\" Then
                p = p + 1           ' skip the escaped character
            ElseIf c = """" Then
                insideStr = False
            End If
        Else
            Select Case c
                Case """"
                    insideStr = True    ' entering a string value
                Case "{"
                    If depth = 0 Then oStart = p
                    depth = depth + 1       ' always increment (fixes bug #2)
                Case "}"
                    depth = depth - 1
                    If depth = 0 And oStart > 0 Then
                        items(count) = Mid(json, oStart, p - oStart + 1)
                        count = count + 1
                        oStart = 0
                    End If
                Case "]"
                    If depth = 0 Then Exit Do
            End Select
        End If

        p = p + 1
    Loop
    ParseItems = count
End Function

Private Function JVal(json As String, key As String) As String
    ' Finds "key" then skips optional whitespace + colon + whitespace before value.
    ' Handles both compact JSON ("key":"val") and pretty-printed ("key" : "val").
    Dim pat As String : pat = """" & key & """"
    Dim sp As Long    : sp = InStr(json, pat)
    If sp = 0 Then Exit Function
    sp = sp + Len(pat)
    ' skip whitespace before colon
    Do While sp <= Len(json) And Mid(json, sp, 1) = " " : sp = sp + 1 : Loop
    ' expect colon
    If Mid(json, sp, 1) <> ":" Then Exit Function
    sp = sp + 1
    ' skip whitespace after colon
    Do While sp <= Len(json) And Mid(json, sp, 1) = " " : sp = sp + 1 : Loop
    Dim ch As String : ch = Mid(json, sp, 1)
    If ch = """" Then
        sp = sp + 1
        Dim ep As Long : ep = sp
        Do While ep <= Len(json)
            If Mid(json, ep, 1) = """" And Mid(json, ep - 1, 1) <> "\" Then Exit Do
            ep = ep + 1
        Loop
        JVal = Mid(json, sp, ep - sp)
    ElseIf Left(Mid(json, sp), 4) = "null" Then
        JVal = ""
    Else
        Dim e1 As Long : e1 = InStr(sp, json, ",") : If e1 = 0 Then e1 = Len(json)
        Dim e2 As Long : e2 = InStr(sp, json, "}") : If e2 = 0 Then e2 = Len(json)
        JVal = Trim(Mid(json, sp, IIf(e1 < e2, e1, e2) - sp))
    End If
End Function

Private Function FileContentsHref(itemJson As String) As String
    ' Locates the link object whose "name" is "FileContents" and returns its "href".
    ' Tolerates pretty-printed JSON with spaces around colons.
    Dim p As Long : p = 1
    Do
        ' Find "name" key with value "FileContents" (allow spaces around colon)
        Dim np As Long : np = InStr(p, itemJson, """name""")
        If np = 0 Then Exit Do
        ' skip to colon then to value
        Dim cp As Long : cp = InStr(np + 6, itemJson, ":")
        If cp = 0 Then Exit Do
        cp = cp + 1
        Do While cp <= Len(itemJson) And Mid(itemJson, cp, 1) = " " : cp = cp + 1 : Loop
        If Mid(itemJson, cp, 14) = """FileContents""" Then
            ' found the right link object — now look backward for the href value
            Dim hp As Long : hp = InStrRev(itemJson, """href""", np)
            If hp > 0 Then
                Dim hc As Long : hc = InStr(hp + 6, itemJson, ":")
                If hc > 0 Then
                    hc = hc + 1
                    Do While hc <= Len(itemJson) And Mid(itemJson, hc, 1) = " " : hc = hc + 1 : Loop
                    If Mid(itemJson, hc, 1) = """" Then
                        hc = hc + 1
                        Dim he As Long : he = InStr(hc, itemJson, """")
                        FileContentsHref = Mid(itemJson, hc, he - hc)
                        Exit Function
                    End If
                End If
            End If
        End If
        p = np + 1
    Loop
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  BASE-64  (MSXML2 — same COM library Oracle tools use)
' ═══════════════════════════════════════════════════════════════════════════════
Private Function B64(s As String) As String
    Dim xml  As Object : Set xml  = CreateObject("MSXML2.DOMDocument")
    Dim node As Object : Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    Dim st As Object : Set st = CreateObject("ADODB.Stream")
    st.Type = 2 : st.CharSet = "us-ascii" : st.Open : st.WriteText s
    st.Position = 0 : st.Type = 1
    node.nodeTypedValue = st.Read : st.Close
    B64 = Replace(node.Text, vbLf, "")
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  UTILITY
' ═══════════════════════════════════════════════════════════════════════════════
Private Function InvoiceFileName(suppName As String, invNum As String, _
                                  invAmt As String, attachIndex As Long) As String
    ' Builds:  Supplier Name - InvoiceNumber - Amount.pdf
    ' If there are multiple scanned images on one invoice, appends _2, _3, ...
    Dim amt As String
    If invAmt <> "" Then
        On Error Resume Next
        amt = Format(CDbl(invAmt), "0.00")
        If Err.Number <> 0 Then amt = invAmt
        On Error GoTo 0
    End If

    Dim base As String
    If suppName <> "" Then base = suppName & "-"
    base = base & invNum
    If amt <> "" Then base = base & "-" & amt

    Dim suffix As String
    If attachIndex > 0 Then suffix = "_" & (attachIndex + 1) Else suffix = ""

    ' SafeName strips chars illegal in Windows filenames; suffix and .pdf added after
    InvoiceFileName = SafeName(base, invNum) & suffix & ".pdf"
End Function

Private Function SafeName(name As String, default_ As String) As String
    If Trim(name) = "" Then SafeName = default_ : Exit Function
    Dim s As String : s = name
    Dim x As Variant
    For Each x In Array("/", "\", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, x, "_")
    Next x
    SafeName = s
End Function

' ═══════════════════════════════════════════════════════════════════════════════
'  LOG SHEET
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub InitLog()
    Dim ws As Worksheet : Set ws = ThisWorkbook.Sheets(SH_LOG)
    ws.Range("A3:D" & ws.Rows.Count).ClearContents
    ws.Range("A3:D" & ws.Rows.Count).Interior.ColorIndex = xlNone
End Sub

Private Sub WriteLog(rowNum As Long, invNum As String, invId As String, _
                     fName As String, status As String, savedPath As String, _
                     Optional success As Variant)
    Dim ws As Worksheet : Set ws = ThisWorkbook.Sheets(SH_LOG)
    ws.Cells(rowNum, 1).Value = invNum
    ws.Cells(rowNum, 2).Value = fName
    ws.Cells(rowNum, 3).Value = status
    ws.Cells(rowNum, 4).Value = savedPath
    Dim clr As Long
    If IsMissing(success) Or IsNull(success) Then
        clr = CLR_YELLOW
    ElseIf CBool(success) Then
        clr = CLR_GREEN
    Else
        clr = CLR_RED
    End If
    ws.Cells(rowNum, 3).Interior.Color = clr
    DoEvents
End Sub
