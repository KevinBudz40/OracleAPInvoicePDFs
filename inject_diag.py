"""Injects a DiagTest sub into Module1 so the user can run it from the VBA IDE
to see exactly what URL is called, what the auth header looks like, and what
Oracle returns."""

import os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import win32com.client as win32

PATH = r"D:\Nuke\AP_Invoice_Downloader.xlsm"

DIAG = r"""
Public Sub DiagTest()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Parameters")
    Dim host As String, user As String, pass_ As String, invNum As String
    host   = Trim(ws.Range("B3").Value)
    user   = Trim(ws.Range("B4").Value)
    pass_  = Trim(ws.Range("B5").Value)
    invNum = Trim(CStr(ws.Cells(17, 1).Value))

    Dim url As String
    url = host & "/fscmRestApi/resources/11.13.18.05/invoices" & _
          "?q=InvoiceNumber%3D%27" & invNum & "%27" & _
          "&fields=InvoiceId%2CInvoiceNumber&limit=1"

    Dim authHdr As String
    authHdr = "Basic " & B64(user & ":" & pass_)

    Dim msg As String
    msg = "invNum = [" & invNum & "]" & vbLf
    msg = msg & "URL = " & url & vbLf
    msg = msg & "Auth (first 40) = " & Left(authHdr, 40) & vbLf & vbLf

    Dim h As Object
    Set h = CreateObject("MSXML2.XMLHTTP.6.0")
    On Error Resume Next
    h.Open "GET", url, False
    h.setRequestHeader "Authorization", authHdr
    h.setRequestHeader "Accept", "application/json"
    h.Send
    msg = msg & "HTTP Status = " & h.Status & vbLf
    If Err.Number <> 0 Then msg = msg & "Error: " & Err.Description & vbLf
    On Error GoTo 0
    Dim resp As String
    resp = h.ResponseText
    msg = msg & "Response length = " & Len(resp) & vbLf
    msg = msg & "Response start = " & Left(resp, 300)
    MsgBox msg, vbInformation, "DiagTest"
End Sub
"""

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AutomationSecurity = 3

wb = xl.Workbooks.Open(os.path.abspath(PATH), UpdateLinks=0, ReadOnly=False)
xl.AutomationSecurity = 1

vbp = wb.VBProject
m1  = vbp.VBComponents("Module1")
cm  = m1.CodeModule
cm.InsertLines(cm.CountOfLines + 1, DIAG)
print(f"DiagTest added. Module1 now has {cm.CountOfLines} lines.")

wb.Save()
wb.Close(False)
xl.Quit()
print("Saved OK.")
