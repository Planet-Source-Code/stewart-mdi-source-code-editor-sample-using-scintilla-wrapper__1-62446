VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Scintilla Class MDI App"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10350
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6930
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7488
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4200
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09DC
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AEE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C00
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D12
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E24
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F36
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1048
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":115A
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":126C
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":137E
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1490
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A2
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Document"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Document"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Document &As"
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExportHTML 
         Caption         =   "&Export to HTML"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Document"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Co&py"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSyntax 
         Caption         =   "&Syntax Settings"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTimeDate 
         Caption         =   "&Time Date"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFindPrevious 
         Caption         =   "Find Previous"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "&Goto"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuHighlighters 
      Caption         =   "&Highlighter"
      Begin VB.Menu mnuHighlighter 
         Caption         =   "Highlighter"
         Index           =   0
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Window"
      Begin VB.Menu mnuHor 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnuVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowList 
         Caption         =   "&Window List"
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
  Setup
End Sub

Public Sub Setup()
  On Error Resume Next
  LoadDirectory App.Path & "\highlighters"
  Call SetupMenu
  Me.WindowState = GetSetting("ScintillaMDI", "Settings", "MDIWState", 0)
  Me.Left = GetSetting("ScintillaMDI", "Settings", "MDILeft", (Screen.Width - Me.Width) \ 2)
  Me.Top = GetSetting("ScintillaMDI", "Settings", "MDITop", (Screen.Height - Me.Height) \ 2)
  Me.Width = GetSetting("ScintillaMDI", "Settings", "MDIWidth", Me.Width)
  Me.Height = GetSetting("ScintillaMDI", "Settings", "MDIHeight", Me.Height)
  NewDoc "New Document"
  Me.Arrange vbCascade
End Sub

Public Function AddMenu(sCaption As String, sTag As String, iIndex As Integer) As Integer
  On Error Resume Next
  If iIndex > 0 Then Load mnuHighlighter(iIndex)
  mnuHighlighter(iIndex).Caption = sCaption ' sCaption we got from the "Identify" function on the plugin
  mnuHighlighter(iIndex).Visible = True
  mnuHighlighter(iIndex).Enabled = True
  mnuHighlighter(iIndex).Tag = sTag ' We store the interface to the plugin in here, to later use it on the event of a menu click
End Function

Public Function SetupMenu()
  Dim i As Integer
  For i = 0 To UBound(Highlighters) - 1
    AddMenu Highlighters(i).strName, Highlighters(i).strName, i
  Next i
End Function

Private Sub MDIForm_Unload(Cancel As Integer)
  If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
  SaveSetting "ScintillaMDI", "Settings", "MDIWState", Me.WindowState
  Me.WindowState = vbNormal
  SaveSetting "ScintillaMDI", "Settings", "MDILeft", Me.Left
  SaveSetting "ScintillaMDI", "Settings", "MDITop", Me.Top
  SaveSetting "ScintillaMDI", "Settings", "MDIWidth", Me.Width
  SaveSetting "ScintillaMDI", "Settings", "MDIHeight", Me.Height
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal, Me
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuArrange_Click()
  Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuCascade_Click()
  Me.Arrange vbCascade
End Sub

Private Sub mnuCopy_Click()
  On Error Resume Next
  ActiveForm.sciMain.Copy
End Sub

Private Sub mnuCut_Click()
  On Error Resume Next
  ActiveForm.sciMain.Cut
End Sub

Private Sub mnuExportHTML_Click()
  On Error Resume Next
  With cd
    .Filter = "HTML Files (*.html, *.htm)|*.html;*.htm)|All Files (*.*)|*.*"
    .ShowSave
    If .filename <> "" Then
      ExportToHTML .filename, ActiveForm.sciMain
    End If
    ActiveForm.sciMain.SetFocus
  End With
End Sub

Private Sub mnuFind_Click()
  On Error Resume Next
  ActiveForm.sciMain.DoFind
End Sub

Private Sub mnuGoto_Click()
  On Error Resume Next
  ActiveForm.sciMain.DoGoto
End Sub

Private Sub mnuHighlighter_Click(index As Integer)
  On Error Resume Next
  SetHighlighter ActiveForm.sciMain, mnuHighlighter(index).Tag
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuHor_Click()
  Me.Arrange vbHorizontal
End Sub

Private Sub mnuNew_Click()
  NewDoc "New Document"
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuOpen_Click()
  Dim i As Long
  With cd
    .Filter = "All Files (*.*)|*.*|"
    For i = 0 To UBound(Highlighters) - 1
      If Highlighters(i).strFilter <> "" Then .Filter = .Filter & Highlighters(i).strFilter
    Next i
    .ShowOpen
    If .filename <> "" Then NewDoc .filename
  End With
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub NewDoc(strFile As String)
  On Error Resume Next
  Static lDocumentCount As Long
  Dim doc As New frmDoc
  LockWindowUpdate Me.hWnd
  Load doc
  lDocumentCount = lDocumentCount + 1
  If Dir(strFile) <> "" Then
    doc.sciMain.LoadFile (strFile)
    doc.Caption = strFile
    doc.strFile = strFile
    SetHighlighter doc.sciMain, SetHighlighterBasedOnExtension(strFile)
  Else
    doc.Caption = strFile & " " & lDocumentCount
  End If
  doc.Show
  doc.sciMain.SetFocus
  LockWindowUpdate 0
End Sub

Private Sub mnuPaste_Click()
  On Error Resume Next
  ActiveForm.sciMain.Paste
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuPrint_Click()
  On Error Resume Next
  ActiveForm.sciMain.PrintDoc
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuRedo_Click()
  On Error Resume Next
  ActiveForm.sciMain.Redo
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuReplace_Click()
  On Error Resume Next
  ActiveForm.sciMain.DoReplace
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuSave_Click()
  On Error Resume Next
  If ActiveForm.strFile <> "" Then
    ActiveForm.sciMain.SaveToFile (ActiveForm.strFile)
  Else
    DoSaveAs
  End If
  ActiveForm.sciMain.SetFocus
End Sub

Public Sub SaveDoc()
  mnuSave_Click
  ActiveForm.sciMain.SetFocus
End Sub

Public Sub DoSaveAs()
  On Error Resume Next
  Dim i As Long
  With cd
    .Filter = "All Files (*.*)|*.*|"
    For i = 0 To UBound(Highlighters) - 1
      If Highlighters(i).strFilter <> "" Then .Filter = .Filter & Highlighters(i).strFilter
    Next i
    .ShowSave
    If .filename <> "" Then
      ActiveForm.sciMain.SaveToFile .filename
      ActiveForm.strFile = .filename
      ActiveForm.Caption = .filename
      SetHighlighter ActiveForm.sciMain, SetHighlighterBasedOnExtension(.filename)
    End If
  End With
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuSaveAs_Click()
  On Error Resume Next
  DoSaveAs
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuSelAll_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelectAll
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuSyntax_Click()
  DoSyntaxOptions App.Path & "\highlighters\", Me
  ResetSyntaxMDI "frmDoc"
  ActiveForm.sciMain.SetFocus
  SetupMenu
End Sub

Private Sub mnuTimeDate_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelText = Time & " | " & Date
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuUndo_Click()
  On Error Resume Next
  ActiveForm.sciMain.Undo
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuVertical_Click()
  Me.Arrange vbVertical
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Select Case LCase(Button.key)
    Case "new"
      mnuNew_Click
    Case "open"
      mnuOpen_Click
    Case "save"
      mnuSave_Click
    Case "print"
      mnuPrint_Click
    Case "cut"
      mnuCut_Click
    Case "copy"
      mnuCopy_Click
    Case "paste"
      mnuPaste_Click
  End Select
  ActiveForm.sciMain.SetFocus
End Sub
