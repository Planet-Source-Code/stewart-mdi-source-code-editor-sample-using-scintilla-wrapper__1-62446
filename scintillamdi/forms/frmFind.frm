VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDirection 
      Caption         =   "Direction"
      Height          =   855
      Left            =   3120
      TabIndex        =   10
      Top             =   600
      Width           =   1575
      Begin VB.OptionButton optDown 
         Caption         =   "&Down"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optUp 
         Caption         =   "&Up"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ComboBox cmbFind 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   150
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox chkRegExp 
      Caption         =   "Regular &expression"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1500
      Width           =   1695
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Match &case"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CheckBox chkWhole 
      Caption         =   "Match &whole word only"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   900
      Width           =   2535
   End
   Begin VB.CheckBox chkWrap 
      Caption         =   "Wrap aroun&d"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "Find what:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   210
      Width           =   735
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bFind As Boolean
Public DoWhat As Integer

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFind_Click()
  DoWhat = 1
  Me.Hide
End Sub

Private Sub Form_Load()
  DoWhat = 0
  Flatten Me
  Me.Left = GetSetting("ScintillaClass", "Settings", "FindLeft", (Screen.Width - Me.Width) \ 2)
  Me.Top = GetSetting("ScintillaClass", "Settings", "FindTop", (Screen.Height - Me.Height) \ 2)
  chkCase.value = GetSetting("ScintillaClass", "Settings", "FchkCase", 0)
  chkRegExp.value = GetSetting("ScintillaClass", "Settings", "FchkRegEx", 0)
  chkWhole.value = GetSetting("ScintillaClass", "Settings", "FchkWhole", 0)
  chkWrap.value = GetSetting("ScintillaClass", "Settings", "FchkWrap", 1)
  optUp.value = GetSetting("ScintillaClass", "Settings", "FOptUp", 0)
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting "ScintillaClass", "Settings", "FindLeft", Me.Left
  SaveSetting "ScintillaClass", "Settings", "FindTop", Me.Top
  SaveSetting "ScintillaClass", "Settings", "FchkCase", chkCase.value
  SaveSetting "ScintillaClass", "Settings", "FchkRegEx", chkRegExp.value
  SaveSetting "ScintillaClass", "Settings", "FchkWhole", chkWhole.value
  SaveSetting "ScintillaClass", "Settings", "FchkWrap", chkWrap.value
  SaveSetting "ScintillaClass", "Settings", "FOptUp", optUp.value
End Sub
