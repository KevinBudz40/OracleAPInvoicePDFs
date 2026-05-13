VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PasswordForm
   Caption         =   "Oracle AP Invoice Downloader"
   ClientHeight    =   1635
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   3780
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PasswordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Controls are created at runtime so no companion .frx binary is required.
Private WithEvents btnOK     As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private txtPwd               As MSForms.TextBox

Public Cancelled As Boolean

' ── build the form the first time the predeclared instance is used ────────────
Private Sub UserForm_Initialize()
    Me.Width  = 300
    Me.Height = 155

    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblPrompt")
    lbl.Caption = "Oracle Cloud password:"
    lbl.Left = 12 : lbl.Top = 12 : lbl.Width = 264 : lbl.Height = 18

    Set txtPwd = Me.Controls.Add("Forms.TextBox.1", "txtPwd")
    txtPwd.PasswordChar = "*"
    txtPwd.Left = 12 : txtPwd.Top = 36 : txtPwd.Width = 264 : txtPwd.Height = 24

    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    btnOK.Caption = "OK" : btnOK.Default = True
    btnOK.Left = 96 : btnOK.Top = 72 : btnOK.Width = 72 : btnOK.Height = 24

    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    btnCancel.Caption = "Cancel" : btnCancel.Cancel = True
    btnCancel.Left = 180 : btnCancel.Top = 72 : btnCancel.Width = 72 : btnCancel.Height = 24
End Sub

' ── called by DownloadInvoicePDFs; returns "" if user cancels ────────────────
Public Function GetPassword() As String
    Cancelled = False
    If Not txtPwd Is Nothing Then txtPwd.Text = ""   ' clear any previous entry
    Me.Show vbModal
    If Not Cancelled Then GetPassword = txtPwd.Text
End Function

' ── event handlers ────────────────────────────────────────────────────────────
Private Sub btnOK_Click()
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Treat the red-X close button the same as Cancel
    If CloseMode = vbFormControlMenu Then
        Cancelled = True
    End If
End Sub
