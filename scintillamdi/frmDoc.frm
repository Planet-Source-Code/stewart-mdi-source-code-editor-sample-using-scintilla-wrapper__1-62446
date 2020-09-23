VERSION 5.00
Begin VB.Form frmDoc 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents sciMain As clsScintilla
Attribute sciMain.VB_VarHelpID = -1
Public strFile As String
Public lHighlighter As Long

Private Sub Form_Activate()
  sciMain.SetFocus
  sciMain_UpdateUI
End Sub

Private Sub Form_Load()
  Set sciMain = New clsScintilla
  sciMain.CreateScintilla Me, True, frmMain
  sciMain.ScrollWidth = 10000
  sciMain.Attach Me
  sciMain.Folding = True
  sciMain.ShowCallTips = True
  sciMain.LineNumbers = True
  sciMain.AutoIndent = True
  sciMain.SetMarginWidth MarginLineNumbers, 50
  Call SetHighlighter(sciMain, "CPP")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim msgResp As VbMsgBoxResult
  If sciMain.Modified = True Then
    msgResp = MsgBox("File [" & Me.Caption & "]" & vbCrLf & "has been modified.  Do you wish to save?", vbYesNoCancel, "Modified")
    Select Case msgResp
      Case vbYes
        frmMain.SaveDoc
      Case vbCancel
        Cancel = True
    End Select
  End If
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  sciMain.SizeScintilla 0, 0, Me.ScaleWidth / Screen.TwipsPerPixelX, (Me.ScaleHeight / Screen.TwipsPerPixelY)
End Sub

Private Sub sciMain_SavePointLeft()
  frmMain.stb.Panels(2).Text = "Modified"
End Sub

Private Sub sciMain_SavePointReached()
  frmMain.stb.Panels(2).Text = ""
End Sub

Private Sub sciMain_UpdateUI()
  frmMain.stb.Panels(1).Text = "CurrentLine: " & sciMain.GetCurLine & " Column: " & sciMain.GetColumn & " Lines: " & sciMain.GetLineCount
End Sub

