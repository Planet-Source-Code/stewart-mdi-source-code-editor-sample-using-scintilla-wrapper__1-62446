VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScintilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'+-------------------------------------------------------------------------------+
'| clsScintilla:  Initial Create Date (23 Aug 2005)                              |
'+-------------------------------------------------------------------------------+
'| The goal of this module is to provide a basic wrapper class for Scintilla     |
'| for us in Visual Basic.  It provides short easy to use functions to handle    |
'| some of the common features of Scintilla.  It also has a few hard coded in    |
'| languages.  Using it's functions you could easily add more languages.  Though |
'| as a general rule it's not as easy to use as a control it's easier to keep    |
'| updated and add more features too.  This makes it easier to keep up to date.  |
'+-------------------------------------------------------------------------------+
'| Copyrights etc.                                                               |
'+-------------------------------------------------------------------------------+
'| If you find this useful you are more than welcome to utilize it in your own   |
'| applications.  If you make any useful modifications to it feel free to email  |
'| them to admin@ceditmx.com.  You may use it in any way you want.  I take no    |
'| responsability for any illegal activities you may use it for (though I don't  |
'| see how).  If you do use it please be kind and post my name and website       |
'| http://www.ceditmx.com in your about box, and don't forget to give credit to  |
'| who it really belongs to.  The authors of scintilla.  Their website is at     |
'| htttp://www.scintilla.org.                                                    |
'+-------------------------------------------------------------------------------+
'| Bugs, etc.                                                                    |
'+-------------------------------------------------------------------------------+
'| If you find a bug or a problem with this please email me again at             |
'| admin@ceditmx.com with a description of what function has a problem so I can  |
'| take a look at it.                                                            |
'+-------------------------------------------------------------------------------+
'| With that being said please enjoy and I hope you find it useful :)            |
'+-------------------------------------------------------------------------------+

Public SCI As Long

Implements ISubclass

' Variable's where multiple options are available
Public Enum enHighlighter
  hlVB = 0
  hlCPP = 1
  hlHTML = 2
  hlPascal = 3
End Enum

Public Enum dcShiftDirection
    lLeft = -1
    lRight = 0
End Enum

Public Enum enMargins
  MarginLineNumbers = 0
  MarginFold = 1
End Enum

' Events which will be called by scintilla's result code
Public Event OnModified()
Public Event StyleNeeded()
Public Event CharAdded()
Public Event SavePointReached()
Public Event SavePointLeft()
Public Event DoubleClick()
Public Event UpdateUI()
Public Event MarginClick()
Public Event Paint()
Public Event UserListSelection()

'+----------------------------------------------------+
'| Begin Properties - These will be properties which  |
'| will control the look and feel of Scintilla        |
'+----------------------------------------------------+
Private bFold As Boolean
Private bLineNumbers As Boolean
Private strText As String
Private bContextMenu As Boolean
Private strSelText As String
Private bAutoIndent As Boolean
Private eolLineBreak
Private bWordWrap As Boolean
Private bIgnoreAutoCCase

Private bRegEx As Boolean
Private bWholeWord As Boolean
Private bWrap As Boolean
Private bWordStart As Boolean
Private bCase As Boolean
Private strFind As String
Private bFindEvent As Boolean
Private bFindInRange As Boolean
Private bFindReverse As Boolean
Private bShowCallTips As Boolean


Private APIStrings() As String

Private ActiveCallTip As Integer
Private CallTipPos As Integer

Private bRepLng As Boolean
Private bRepAll As Boolean

Private Owner As Form

Private APIStringLoaded As Boolean

Public lHighlighterNum As Long  'This is to work with the highlighter module

Public Property Get ShowCallTips() As Boolean
  ShowCallTips = bShowCallTips
End Property

Public Property Let ShowCallTips(bVal As Boolean)
  bShowCallTips = bVal
End Property

Public Property Get ViewWhiteSpace() As Boolean
  ViewWhiteSpace = GetViewWhiteSpace
End Property

Public Property Let ViewWhiteSpace(bVal As Boolean)
  SetViewWhiteSpace bVal
End Property


Public Property Get ScrollWidth() As Long
  ScrollWidth = GetScrollWidth
End Property

Public Property Let ScrollWidth(bVal As Long)
  SetScrollWidth bVal
End Property

Public Property Get IgnoreAutoCCase() As Boolean
  IgnoreAutoCCase = bIgnoreAutoCCase
End Property

Public Property Let IgnoreAutoCCase(ByVal bVal As Boolean)
  bIgnoreAutoCCase = bVal
  SetAutoCompleteIgnoreCase bVal
End Property

Public Property Get WordWrap() As Boolean
  WordWrap = bWordWrap
End Property

Public Property Let WordWrap(ByVal bVal As Boolean)
  bWordWrap = bVal
  SetWordWrap bVal
End Property

Public Property Get LineBreak() As EOL
  LineBreak = eolLineBreak
End Property

Public Property Let LineBreak(ByVal eolVal As EOL)
  eolLineBreak = eolVal
  SetLineBreak eolVal
End Property

Public Property Get ReadOnly() As Boolean
  ReadOnly = GetReadOnly
End Property

Public Property Let ReadOnly(ByVal bVal As Boolean)
  SetReadOnly CLng(bVal)
End Property

Public Property Get ContextMenu() As Boolean
  ContextMenu = bContextMenu
End Property

Public Property Let ContextMenu(ByVal bVal As Boolean)
  bContextMenu = bVal
  SetContextMenu bVal
End Property

Public Property Get AutoIndent() As Boolean
  AutoIndent = bAutoIndent
End Property

Public Property Let AutoIndent(ByVal bVal As Boolean)
  bAutoIndent = bVal
End Property

Public Property Get LineNumbers() As Boolean
  LineNumbers = bLineNumbers
End Property

Public Property Let LineNumbers(ByVal bVal As Boolean)
  bLineNumbers = bVal
  SetLineNumbers bVal
End Property

Public Property Get Folding() As Boolean
    Folding = bFold
End Property

Public Property Let Folding(ByVal bVal As Boolean)
    bFold = bVal
    SetFolding bFold
End Property

Public Property Get SelText() As String
  SelText = GetSelText
End Property

Public Property Let SelText(ByVal strVal As String)
  strSelText = strVal
  SetSelText strVal
End Property


Public Property Get Text() As String
  Text = GetText
End Property

Public Property Let Text(ByVal strVal As String)
  strText = strVal
  SetText strVal
End Property

Public Sub SetLineNumbers(bVal As Boolean)
  If bVal = True Then
    SendEditor SCI_SETMARGINWIDTHN, 0, 30
  Else
    SendEditor SCI_SETMARGINWIDTHN, 0, 0
  End If
End Sub

Public Sub SetReadOnly(bVal As Long)
  SendEditor SCI_SETREADONLY, bVal
End Sub

Public Sub SetProperty(key As String, value As String)
  Dim l1 As Long, l2 As Long
  Dim bKey() As Byte
  bKey = StrConv(key, vbFromUnicode)
  ReDim Preserve bKey(0 To UBound(bKey) + 1) As Byte
  Dim bValue() As Byte
  bValue = StrConv(value, vbFromUnicode)
  ReDim Preserve bValue(0 To UBound(bValue) + 1) As Byte
  SendMessage SCI, SCI_SETPROPERTY, ByVal VarPtr(bKey(0)), ByVal VarPtr(bValue(0))
End Sub

Public Sub ReplaceSel(strVal As String)
  SendEditor SCI_REPLACESEL, Len(strVal), strVal
End Sub

Public Sub SetSelText(strVal As String)
  ReplaceSel strVal
End Sub

Public Function GetSelText() As String
  Dim bByte() As Byte
  Dim lPtr As Long
  Dim strTmp As String
  Dim i As Long
  lPtr = SendMessage(SCI, SCI_GETSELTEXT, 0, 0)
  ReDim Preserve bByte(0 To lPtr)
  SendMessage SCI, SCI_GETSELTEXT, 0, VarPtr(bByte(0))
  strTmp = ""
  For i = 0 To UBound(bByte) - 1
    strTmp = strTmp & Chr(bByte(i))
  Next i
  GetSelText = strTmp
End Function

Public Sub SetFolding(bDoFold As Boolean)
  ' This function will just set a nice set of folding properties for
  ' scintilla to use.
  If bDoFold = True Then
    SetProperty "fold", "1"
    SetProperty "fold.compact", 0
    SetProperty "fold.html", "1"
    SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 0
    Call SendEditor(SCI_SETMARGINTYPEN, MARGIN_SCRIPT_FOLD_INDEX, SC_MARGIN_SYMBOL)
    Call SendEditor(SCI_SETMARGINMASKN, MARGIN_SCRIPT_FOLD_INDEX, SC_MASK_FOLDERS)
    SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 20
    SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDER, SC_MARK_BOXPLUS
    SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN, SC_MARK_BOXMINUS
    SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND, SC_MARK_BOXPLUSCONNECTED
    SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_BOXMINUSCONNECTED
    SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER
    
    SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB, SC_MARK_VLINE
    SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNER
    SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDER, vbWhite
    SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDER, vbBlack
    SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDEROPEN, vbWhite
    SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDEROPEN, vbBlack
    SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDERMIDTAIL, vbBlack
    SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDERSUB, vbBlack
    SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDERTAIL, vbBlack
    
    SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDEROPEN, vbWhite
    SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDERMIDTAIL, vbWhite
    SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDERSUB, vbWhite
    SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDERTAIL, vbWhite
    
    SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDEROPENMID, vbBlack
    SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDEROPENMID, vbWhite
    
    SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDEREND, vbBlack
    SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDEREND, vbWhite
    
    
    Call SendEditor(SCI_SETMARGINSENSITIVEN, MARGIN_SCRIPT_FOLD_INDEX, 1)
    Call SendEditor(SCI_SETFOLDFLAGS, 16, 0)
  Else
    SetProperty "fold", "0"
    SetProperty "fold.compact", 0
    SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 0
 End If
End Sub


Public Function SendEditor(ByVal Msg As Long, Optional ByVal wParam As Long = 0, Optional ByVal lParam = 0) As Long
    If VarType(lParam) = vbString Then
        SendEditor = SendMessageString(SCI, Msg, IIf(wParam = 0, CLng(wParam), wParam), CStr(lParam))
    Else
        SendEditor = SendMessage(SCI, Msg, IIf(wParam = 0, CLng(wParam), wParam), IIf(lParam = 0, CLng(lParam), lParam))
    End If
End Function

Public Sub SetHighlighter(iHighlighter As enHighlighter)
  StyleClearALL
  StartStyle
  Select Case iHighlighter
    Case hlVB
      SetHighlightVB
    Case hlHTML
      SetHighlightHTML
    Case hlCPP
      SetHighlightCPP
    Case hlPascal
      SetHighlightPascal
  End Select
  Colourise
End Sub

Public Function GetLineEndPosition(lLine As Long) As Long
  GetLineEndPosition = SendEditor(SCI_GETLINEENDPOSITION, lLine)
End Function

Public Function GetLineIndentation(lLine As Long) As Long
  GetLineIndentation = SendEditor(SCI_GETLINEINDENTATION, lLine)
End Function

Public Sub SetLineIndentation(lLine As Long, lIndent As Long)
  SendEditor SCI_SETLINEINDENTATION, lLine, lIndent
End Sub

Public Function GetLineIndentPosition(lLine As Long) As Long
  GetLineIndentPosition = SendEditor(SCI_GETLINEINDENTPOSITION, lLine)
End Function

Public Sub SetCurrentPosition(lval As Long)
  SendEditor SCI_SETCURRENTPOS, lval
End Sub

Public Sub SetSel(lStart As Long, lEnd As Long)
  SendEditor SCI_SETSEL, lStart, lEnd
End Sub

Public Function PositionFromLine(lLine As Long) As Long
  PositionFromLine = SendEditor(SCI_POSITIONFROMLINE, lLine)
End Function

Public Sub MaintainIndentation()
  Dim indentAmount As Long
  Dim lastLine As Long
  Dim curLine As Long
  curLine = GetCurLine + 1
  lastLine = curLine - 1
  indentAmount = 0
  While (lastLine >= 0) And (GetLineEndPosition(lastLine) - PositionFromLine(lastLine) = 0)
    lastLine = lastLine - 1
    If lastLine >= 0 Then
      indentAmount = GetLineIndentation(lastLine)
    End If
    If indentAmount > 0 Then
      Call SetLineIndentation(curLine, indentAmount)
      Call SetCurrentPosition(GetLineIndentPosition(curLine))
      Call SetSel(GetCurPos, GetCurPos)
    End If
  Wend
End Sub

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim scMsg As SCNotification
    Dim tHdr As NMHDR
    Select Case iMsg
      Case WM_NOTIFY
        CopyMemory scMsg, ByVal lParam, Len(scMsg)
        tHdr = scMsg.NotifyHeader
        If (tHdr.hwndFrom = SCI) Then
          'Scintilla has given some information.  Let's see what it is
          'and route it to the proper place.
          ' Any commented with TODO have not been implimented yet.
          Select Case tHdr.code
            Case SCN_MODIFIED
              RaiseEvent OnModified
            Case SCN_STYLENEEDED
              RaiseEvent StyleNeeded
            Case SCN_CHARADDED
              If bAutoIndent = True Then
                MaintainIndentation
              End If
              If bShowCallTips Then
                StartCallTip scMsg.ch
              End If
              
              RaiseEvent CharAdded
            Case SCN_SAVEPOINTREACHED
              RaiseEvent SavePointReached
            Case SCN_SAVEPOINTLEFT
              RaiseEvent SavePointLeft
            Case SCN_MODIFYATTEMPTRO
              'TODO
            Case SCN_DOUBLECLICK
              RaiseEvent DoubleClick
            Case SCN_UPDATEUI
              'If m_MatchBraces Then
                  Dim pos As Long, pos2 As Long
                  pos2 = INVALID_POSITION
                  If IsBrace(CharAtPos(GetCurPos)) Then
                      pos2 = GetCurPos
                  ElseIf IsBrace(CharAtPos(GetCurPos - 1)) Then
                      pos2 = GetCurPos - 1
                  End If
                  If pos2 <> INVALID_POSITION Then
                      pos = SendMessage(SCI, SCI_BRACEMATCH, pos2, CLng(0))
                      If pos = INVALID_POSITION Then
                          Call SendEditor(SCI_BRACEBADLIGHT, pos2)
                      Else
                          Call SendEditor(SCI_BRACEHIGHLIGHT, pos, pos2)
                          'If m_IndGuides Then
                              Call SendEditor(SCI_SETHIGHLIGHTGUIDE, GetColumn)
                          'End If
                      End If
                  Else
                      Call SendEditor(SCI_BRACEHIGHLIGHT, INVALID_POSITION, INVALID_POSITION)
                  End If
              'End If
              RaiseEvent UpdateUI
            Case SCN_MACRORECORD
              'TODO
            Case SCN_MARGINCLICK
              Dim lLine As Long, lMargin As Long, lPosition As Long
              lPosition = scMsg.position
              lLine = SendEditor(SCI_LINEFROMPOSITION, lPosition)
              lMargin = scMsg.margin
              If lMargin = MARGIN_SCRIPT_FOLD_INDEX Then
                Call SendEditor(SCI_TOGGLEFOLD, lLine, 0)
              End If
              RaiseEvent MarginClick
            Case SCN_NEEDSHOWN
              'TODO
            Case SCN_PAINTED
              RaiseEvent Paint
            Case SCN_USERLISTSELECTION
              RaiseEvent UserListSelection
            Case SCN_DWELLSTART
              'TODO
            Case SCN_DWELLEND
              'TODO
          End Select
      End If
    End Select
End Function

Public Function GetColumn()
  GetColumn = SendMessage(SCI, SCI_GETCOLUMN, GetCurPos, CLng(0))
End Function

Public Function GetLastLine() As Long
  GetLastLine = SendEditor(SCI_GETLINECOUNT)
End Function

Private Function IsBrace(ch As Long) As Boolean
    IsBrace = (ch = 40 Or ch = 41 Or ch = 60 Or ch = 62 Or ch = 91 Or ch = 93 Or ch = 123 Or ch = 125)
End Function

Public Function CharAtPos(position As Long) As Long
    CharAtPos = SendMessage(SCI, SCI_GETCHARAT, position, CLng(0))
End Function

Public Function GetLineCount() As Long
  GetLineCount = SendEditor(SCI_GETLINECOUNT)
End Function

Public Function GetCurPos() As Long
  GetCurPos = SendEditor(SCI_GETCURRENTPOS)
End Function

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse

End Property
  
Public Sub Detach(frm As Form)
  DetachMessage Me, frm.hWnd, WM_NOTIFY
End Sub
  
Public Sub Attach(frm As Form)
  Detach frm
  AttachMessage Me, frm.hWnd, WM_NOTIFY
End Sub

Public Sub CreateScintilla(frm As Form, Optional MDIApp As Boolean = False, Optional MDIParent As MDIForm)
  LoadLibrary ("SciLexer.DLL")
  bFindEvent = False
  If MDIApp = False Then
    Set Owner = frm
  Else
    Set Owner = MDIParent
  End If
  APIStringLoaded = False
  SCI = CreateWindowEx(WS_EX_CLIENTEDGE, "Scintilla", "SciMain", WS_CHILD Or WS_VISIBLE, 0, 0, 200, 200, frm.hWnd, 0, App.hInstance, 0)
  RemoveHotKeys
  SetFocus
End Sub

Public Sub SizeScintilla(Left As Long, Top As Long, Width As Long, Height As Long)
  SetWindowPos SCI, 0, Left, Top, Width, Height, 0
End Sub

Public Sub SetText(str As String)
  SendEditor SCI_SETTEXT, 0, str
End Sub

Public Sub LoadFile(strFile As String)
  Dim str As String
  If Dir(strFile) = "" Then Exit Sub  'We don't want to have an error if the file doesn't exist.
  str = GetFile(strFile)
  SetText str
  SetSavePoint
  ClearUndoBuffer
End Sub

Public Sub StartStyle()
  SendEditor SCI_STARTSTYLING, 0, 31
End Sub

Public Sub Colourise()
  SendEditor SCI_COLOURISE, 0, Len(GetText)
End Sub

'+-----------------------------+
'| Begin Style Setting code    |
'+-----------------------------+

Public Sub SetStyleFore(lStyleNum As Long, lColor As Long)
  SendEditor SCI_STYLESETFORE, lStyleNum, lColor
End Sub

Public Sub SetStyleBack(lStyleNum As Long, lColor As Long)
  SendEditor SCI_STYLESETBACK, lStyleNum, lColor
End Sub

Public Sub SetStyleFont(lStyleNum As Long, strFont As String)
  SendEditor SCI_STYLESETFONT, lStyleNum, strFont
End Sub

Public Sub SetStyleSize(lStyleNum As Long, lFontSize As Long)
  SendEditor SCI_STYLESETSIZE, lStyleNum, lFontSize
End Sub

Public Sub SetStyleBold(lStyleNum As Long, bBold As Long)
  SendEditor SCI_STYLESETBOLD, lStyleNum, bBold
End Sub

Public Sub SetStyleVisible(lStyleNum As Long, bVisible As Long)
  SendEditor SCI_STYLESETVISIBLE, lStyleNum, bVisible
End Sub

Public Sub SetStyleEOLFilled(lStyleNum As Long, bEOLFilled As Long)
  SendEditor SCI_STYLESETEOLFILLED, lStyleNum, bEOLFilled
End Sub

Public Sub SetStyleItalic(lStyleNum As Long, bItalic As Long)
  SendEditor SCI_STYLESETITALIC, lStyleNum, bItalic
End Sub

Public Function Modified() As Boolean
  Modified = SendEditor(SCI_GETMODIFY)
End Function

Public Sub SetStyleUnderline(lStyleNum As Long, bUnderline As Long)
  SendEditor SCI_STYLESETUNDERLINE, lStyleNum, bUnderline
End Sub

Public Function GetCurLine() As Long
  GetCurLine = SendEditor(SCI_LINEFROMPOSITION, GetCurPos)
End Function

Public Function GetLen() As Long
  GetLen = SendEditor(SCI_GETLENGTH)
End Function

Public Function GetText() As String
  Dim numChar As Long
  Dim Txt As String
  numChar = SendMessage(SCI, SCI_GETLENGTH, 0, 0) + 1
  Txt = String(numChar, "0")
  SendMessageString SCI, SCI_GETTEXT, numChar, Txt
  GetText = Txt
End Function

Public Sub SaveToFile(strFile As String)
  Dim str As String
  str = GetText
  WriteToFile strFile, str
  
  ' Remove the modified flag from scintilla
  SetSavePoint
  
End Sub

Public Sub SetSavePoint()
  SendEditor SCI_SETSAVEPOINT, 0, 0
End Sub

Public Sub StyleClearALL()
  Dim i As Long
  SetLexer 0
  For i = 0 To 127
    SetStyleBack i, vbWhite
    SetStyleFore i, vbBlack
    SetStyleBold i, False
    SetStyleItalic i, False
    SetStyleUnderline i, False
    SetStyleFont i, "Courier New"
    SetStyleSize i, 10
  Next i
  SendEditor SCI_STYLECLEARALL, 0, 0
End Sub

'+-----------------------------+
'| End Style Setting code      |
'+-----------------------------+

'+-----------------------------+
'| Begin Folding Code          |
'+-----------------------------+

Public Sub SetKeywords(iKeywordNum As Long, strKeywords As String)
  SendEditor SCI_SETKEYWORDS, iKeywordNum, strKeywords
End Sub

Public Sub SetLexer(lLexerNum As Long)
  SendEditor SCI_SETLEXER, lLexerNum
End Sub

Public Sub Cut()
  SendEditor SCI_CUT, 0, 0
End Sub

Public Sub Copy()
  SendEditor SCI_COPY, 0, 0
End Sub

Public Sub Paste()
  SendEditor SCI_PASTE, 0, 0
End Sub

Public Sub Undo()
  SendEditor SCI_UNDO, 0, 0
End Sub

Public Sub Redo()
  SendEditor SCI_REDO, 0, 0
End Sub

Public Sub SelectAll()
  SendEditor SCI_SELECTALL
End Sub

Private Sub SetHighlightVB()
  'SendEditor SCI_STYLESETBITS, 5
  SetStyleFore 1, 32512
  SetStyleFont 1, "Comic Sans MS"
  SetStyleItalic 1, True
  SetStyleSize 1, 10
  SetStyleFore 2, vbRed
  SetStyleSize 2, 10
  SetStyleFore 3, &H800000
  SetStyleSize 3, 10
  SetStyleBold 3, True
  SetStyleFore 4, &H800080
  SetStyleItalic 4, True
  SetStyleSize 4, 10
  SetStyleFore 32, vbBlack
  SetStyleBack 32, vbWhite
  SetStyleSize 32, 10
  SetStyleFore 35, vbRed
  SetStyleBold 35, True
  SetStyleFore 34, &H800000
  SetStyleBold 34, True
  SetKeywords 0, "and begin case call class continue do each else elseif end erase error event exit false for function get gosub goto if implement in load loop lset me mid new next not nothing on or property raiseevent rem resume return rset select set stop sub then to true unload until wend while with withevents attribute alias as boolean byref byte byval const compare currency date declare dim double enum explicit friend global integer let lib long module object option optional preserve private property public redim single static string type variant"
  SetLexer 8
End Sub



Private Sub SetHighlightHTML()
  'This is a big One here
  SetStyleFore 0, &H0
  SetStyleBack 0, &HFFFFFF
  SetStyleFore 1, &H800000
  SetStyleBack 1, &HFFFFFF
  SetStyleFont 1, "Courier New"
  SetStyleSize 1, 10
  SetStyleFore 2, &HFF
  SetStyleBack 2, &HFFFFFF
  SetStyleFont 2, "Courier New"
  SetStyleSize 2, 10
  SetStyleFore 3, &H808000
  SetStyleBack 3, &HFFFFFF
  SetStyleSize 3, 10
  SetStyleFore 4, &HFF
  SetStyleBack 4, &HFFFFFF
  SetStyleFont 4, "Courier New"
  SetStyleSize 4, 10
  SetStyleFore 5, &H7F7F00
  SetStyleBack 5, &HFFFFFF
  SetStyleFore 6, &H7F007F
  SetStyleBack 6, &HFFFFFF
  SetStyleFore 7, &H7F007F
  SetStyleBack 7, &HFFFFFF
  SetStyleFore 8, &H800080
  SetStyleBack 8, &HFFFFFF
  SetStyleFore 9, &H8080
  SetStyleBack 9, &HFFFFFF
  SetStyleFont 9, "Verdana"
  SetStyleSize 9, 9
  SetStyleFore 10, &H800080
  SetStyleBack 10, &HFFFFFF
  SetStyleFont 10, "Times New Roman"
  SetStyleSize 10, 11
  SetStyleFore 11, &H800000
  SetStyleBack 11, &HFFFFFF
  SetStyleFore 12, &HFF0000
  SetStyleBack 12, &HFFFFFF
  SetStyleFore 13, &HFF0000
  SetStyleBack 13, &HFFFFFF
  SetStyleFore 14, &H80
  SetStyleBack 14, &HFFFFFF
  SetStyleFore 15, &H20000000
  SetStyleBack 15, &HFFFF
  SetStyleFore 16, &H20000000
  SetStyleBack 16, &HDFFF
  SetStyleFore 17, &H20000000
  SetStyleBack 17, &HDFFF
  SetStyleFore 18, &HFF0000
  SetStyleBack 18, &HBFEFFF
  SetStyleFore 19, &HFF00FF
  SetStyleBack 19, &HFFEFFF
  SetStyleFore 20, &H0
  SetStyleBack 20, &HFFFFFF
  SetStyleFore 21, &H800000
  SetStyleBack 21, &HFFEFEF
  SetStyleBold 22, True
  SetStyleFore 22, &H800000
  SetStyleBack 22, &HFFEFEF
  SetStyleFore 23, &H6600
  SetStyleBack 23, &HFFEFEF
  SetStyleFore 24, &H80
  SetStyleBack 24, &HFFEFEF
  SetStyleFore 25, &H3399
  SetStyleBack 25, &HFFEFEF
  SetStyleFore 26, &H80
  SetStyleBack 26, &H6666FF
  SetStyleFore 27, &HFF6633
  SetStyleBack 27, &HFFEFEF
  SetStyleFore 28, &H333333
  SetStyleBack 28, &HFFEFEF
  SetStyleFore 29, &H8080
  SetStyleBack 29, &HFFEFEF
  SetStyleFore 30, &H0
  SetStyleBack 30, &HFFFFFF
  SetStyleFore 31, &H660000
  SetStyleBack 31, &HE0CCCC
  SetStyleFore 32, &H0
  SetStyleBack 32, &HFFFFFF
  SetStyleFore 33, &HBD7173
  SetStyleBack 33, &HEFEFEF
  SetStyleFont 33, "Courier New"
  SetStyleSize 33, 8
  SetStyleBold 34, True
  SetStyleFore 34, &HFF0000
  SetStyleBack 34, &HFFFFFF
  SetStyleFont 34, "Times New Roman"
  SetStyleSize 34, 11
  SetStyleBold 35, True
  SetStyleFore 35, &HFF
  SetStyleBack 35, &HFFFFFF
  SetStyleFont 35, "Times New Roman"
  SetStyleSize 35, 11
  SetStyleFore 36, &H0
  SetStyleBack 36, &HFFFFFF
  SetStyleFore 37, &H0
  SetStyleBack 37, &HFFFFFF
  SetStyleFore 38, &H0
  SetStyleBack 38, &HFFFFFF
  SetStyleFore 39, &H0
  SetStyleBack 39, &HFFFFFF
  SetStyleFore 40, &H7F7F
  SetStyleBack 40, &HFFFFFF
  SetStyleBold 41, True
  SetStyleFore 41, &H0
  SetStyleBack 41, &HFFF0F0
  SetStyleFont 41, "Comic Sans MS"
  SetStyleSize 41, 9
  SetStyleFore 42, &H7F00
  SetStyleBack 42, &HFFF0F0
  SetStyleFont 42, "Comic Sans MS"
  SetStyleSize 42, 9
  SetStyleFore 43, &H7F00
  SetStyleBack 43, &HFFF0F0
  SetStyleFont 43, "Comic Sans MS"
  SetStyleSize 43, 9
  SetStyleBold 44, True
  SetStyleFore 44, &H3F703F
  SetStyleBack 44, &HFFF0F0
  SetStyleFont 44, "Comic Sans MS"
  SetStyleSize 44, 9
  SetStyleFore 45, &H7F7F00
  SetStyleBack 45, &HFFF0F0
  SetStyleFont 45, "Comic Sans MS"
  SetStyleSize 45, 9
  SetStyleFore 46, &H0
  SetStyleBack 46, &HFFF0F0
  SetStyleFont 46, "Comic Sans MS"
  SetStyleSize 46, 9
  SetStyleBold 47, True
  SetStyleFore 47, &H7F0000
  SetStyleBack 47, &HFFF0F0
  SetStyleFont 47, "Comic Sans MS"
  SetStyleSize 47, 9
  SetStyleBold 48, True
  SetStyleFore 48, &H7F007F
  SetStyleBack 48, &HFFF0F0
  SetStyleFont 48, "Comic Sans MS"
  SetStyleSize 48, 9
  SetStyleBold 49, True
  SetStyleFore 49, &H7F007F
  SetStyleBack 49, &HFFF0F0
  SetStyleFont 49, "Comic Sans MS"
  SetStyleSize 49, 9
  SetStyleBold 50, True
  SetStyleFore 50, &H0
  SetStyleBack 50, &HFFF0F0
  SetStyleFont 50, "Comic Sans MS"
  SetStyleSize 50, 9
  SetStyleFore 51, &H20000000
  SetStyleBack 51, &HB0BBBF
  SetStyleFore 52, &H20000000
  SetStyleBack 52, &HB0BBFF
  SetStyleFore 71, &H0
  SetStyleBack 71, &HFFEFEF
  SetStyleFont 71, "Lucida Sans Unicode"
  SetStyleSize 71, 10
  SetStyleFore 72, &H8000
  SetStyleBack 72, &HFFEFEF
  SetStyleFont 72, "Comic Sans MS"
  SetStyleSize 72, 9
  SetStyleFore 73, &H808000
  SetStyleBack 73, &HFFEFEF
  SetStyleFont 73, "Lucida Sans Unicode"
  SetStyleSize 73, 10
  SetStyleBold 74, True
  SetStyleFore 74, &H800000
  SetStyleBack 74, &HFFEFEF
  SetStyleFont 74, "Lucida Sans Unicode"
  SetStyleSize 74, 10
  SetStyleFore 75, &H800080
  SetStyleBack 75, &HFFEFEF
  SetStyleFont 75, "Lucida Sans Unicode"
  SetStyleSize 75, 10
  SetStyleFore 76, &H800000
  SetStyleBack 76, &HFFEFEF
  SetStyleFont 76, "Lucida Sans Unicode"
  SetStyleSize 76, 10
  SetStyleFore 77, &H800000
  SetStyleBack 77, &HFF7F7F
  SetStyleFont 77, "Lucida Sans Unicode"
  SetStyleSize 77, 10
  SetStyleFore 81, &H0
  SetStyleBack 81, &HCFCFEF
  SetStyleFont 81, "Lucida Sans Unicode"
  SetStyleSize 81, 10
  SetStyleFore 82, &H8000
  SetStyleBack 82, &HCFCFEF
  SetStyleFont 82, "Comic Sans MS"
  SetStyleSize 82, 9
  SetStyleFore 83, &H808000
  SetStyleBack 83, &HCFCFEF
  SetStyleFont 83, "Lucida Sans Unicode"
  SetStyleSize 83, 10
  SetStyleBold 84, True
  SetStyleFore 84, &H800000
  SetStyleBack 84, &HCFCFEF
  SetStyleFont 84, "Lucida Sans Unicode"
  SetStyleSize 84, 10
  SetStyleFore 85, &H800080
  SetStyleBack 85, &HCFCFEF
  SetStyleFont 85, "Lucida Sans Unicode"
  SetStyleSize 85, 10
  SetStyleFore 86, &H800000
  SetStyleBack 86, &HCFCFEF
  SetStyleFont 86, "Lucida Sans Unicode"
  SetStyleSize 86, 10
  SetStyleFore 87, &H800000
  SetStyleBack 87, &HBF7F7F
  SetStyleFont 87, "Lucida Sans Unicode"
  SetStyleSize 87, 10
  SetStyleFore 90, &H808080
  SetStyleBack 90, &HFFFFFF
  SetStyleFore 91, &H808080
  SetStyleBack 91, &HEFFFEF
  SetStyleFore 92, &H7F00
  SetStyleBack 92, &HEFFFEF
  SetStyleFont 92, "Comic Sans MS"
  SetStyleSize 92, 9
  SetStyleFore 93, &H7F7F00
  SetStyleBack 93, &HEFFFEF
  SetStyleFore 94, &H7F007F
  SetStyleBack 94, &HEFFFEF
  SetStyleFont 94, "Courier New"
  SetStyleSize 94, 10
  SetStyleFore 95, &H7F007F
  SetStyleBack 95, &HEFFFEF
  SetStyleFont 95, "Courier New"
  SetStyleSize 95, 10
  SetStyleBold 96, True
  SetStyleFore 96, &H7F0000
  SetStyleBack 96, &HEFFFEF
  SetStyleFore 97, &H7F
  SetStyleBack 97, &HEFFFEF
  SetStyleFore 98, &H7F
  SetStyleBack 98, &HEFFFEF
  SetStyleBold 99, True
  SetStyleFore 99, &HFF0000
  SetStyleBack 99, &HEFFFEF
  SetStyleBold 100, True
  SetStyleFore 100, &H7F7F00
  SetStyleBack 100, &HEFFFEF
  SetStyleBold 101, True
  SetStyleFore 101, &H20000000
  SetStyleBack 101, &HEFFFEF
  SetStyleFore 102, &H20000000
  SetStyleBack 102, &HEFFFEF
  SetStyleFore 105, &H808080
  SetStyleBack 105, &HFFFFFF
  SetStyleFore 106, &H808080
  SetStyleBack 106, &HCFEFCF
  SetStyleFore 107, &H7F00
  SetStyleBack 107, &HCFEFCF
  SetStyleFont 107, "Comic Sans MS"
  SetStyleSize 107, 9
  SetStyleFore 108, &H7F7F00
  SetStyleBack 108, &HCFEFCF
  SetStyleFore 109, &H7F007F
  SetStyleBack 109, &HCFEFCF
  SetStyleFont 109, "Courier New"
  SetStyleSize 109, 10
  SetStyleFore 110, &H7F007F
  SetStyleBack 110, &HCFEFCF
  SetStyleFont 110, "Courier New"
  SetStyleSize 110, 10
  SetStyleBold 111, True
  SetStyleFore 111, &H7F0000
  SetStyleBack 111, &HCFEFCF
  SetStyleFore 112, &H7F
  SetStyleBack 112, &HCFEFCF
  SetStyleFore 113, &H7F
  SetStyleBack 113, &HCFEFCF
  SetStyleBold 114, True
  SetStyleFore 114, &HFF0000
  SetStyleBack 114, &HCFEFCF
  SetStyleBold 115, True
  SetStyleFore 115, &H7F7F00
  SetStyleBack 115, &HEFFFEF
  SetStyleBold 116, True
  SetStyleFore 116, &H20000000
  SetStyleBack 116, &HCFEFCF
  SetStyleFore 117, &H20000000
  SetStyleBack 117, &HCFEFCF
  SetStyleFore 118, &H330000
  SetStyleBack 118, &HF8F8FF
  SetStyleFore 119, &H7F00
  SetStyleBack 119, &HF8F8FF
  SetStyleFore 120, &H9F00
  SetStyleBack 120, &HF8F8FF
  SetStyleItalic 121, True
  SetStyleFore 121, &H7F007F
  SetStyleBack 121, &HF8F8FF
  SetStyleFore 122, &H99CC
  SetStyleBack 122, &HF8F8FF
  SetStyleItalic 123, True
  SetStyleFore 123, &H7F0000
  SetStyleBack 123, &HF8F8FF
  SetStyleFore 124, &H999999
  SetStyleBack 124, &HF8F8FF
  SetStyleFont 124, "Comic Sans MS"
  SetStyleSize 124, 9
  SetStyleFore 125, &H666666
  SetStyleBack 125, &HF8F8FF
  SetStyleFont 125, "Comic Sans MS"
  SetStyleSize 125, 9
  SetStyleItalic 126, True
  SetStyleFore 126, &H7F0000
  SetStyleBack 126, &HF8F8FF
  SetStyleFore 127, &H0
  SetStyleBack 127, &HF8F8FF
  SetKeywords 0, "a abbr acronym address applet area b base basefont bdo big blockquote body br button caption center cite code col colgroup dd del dfn dir div dl dt em fieldset font form frame frameset h1 h2 h3 h4 h5 h6 head hr html i iframe img input ins isindex kbd lab"
  SetKeywords 1, "abstract boolean break byte case catch char class const continue debugger default delete do double else enum export extends final finally float for function goto if implements import in instanceof int interface long native new package private protected p"
  SetKeywords 2, "and begin case call class continue do each else elseif end erase error event exit false for function get gosub goto if implement in load loop lset me mid new next not nothing on or property raiseevent rem resume return rset select set stop sub then to tr"
  SetKeywords 3, "and assert break class continue def del elif else except exec finally for from global if import in is lambda None not or pass print raise return try while yield"
  SetKeywords 4, "and argv as argc break case cfunction class continue declare default do die echo else elseif empty enddeclare endfor endforeach endif endswitch endwhile e_all e_parse e_error e_warning eval exit extends false for foreach function global http_cookie_vars "
  SetKeywords 5, "ELEMENT DOCTYPE ATTLIST ENTITY NOTATION"
  SetLexer 4
  'SendEditor SCI_STYLESETBITS, 7
End Sub


Private Sub SetHighlightCPP()
  SetStyleFore 0, &H808080
  SetStyleBack 0, &HFFFFFF
  SetStyleFore 1, &H7F00
  SetStyleBack 1, &HFFFFFF
  SetStyleFont 1, "Comic Sans MS"
  SetStyleSize 1, 9
  SetStyleFore 2, &H7F00
  SetStyleBack 2, &HFFFFFF
  SetStyleFont 2, "Comic Sans MS"
  SetStyleSize 2, 9
  SetStyleFore 3, &H3F703F
  SetStyleBack 3, &HFFFFFF
  SetStyleFont 3, "Comic Sans MS"
  SetStyleSize 3, 9
  SetStyleFore 4, &H7F7F00
  SetStyleBack 4, &HFFFFFF
  SetStyleBold 5, True
  SetStyleFore 5, &H800000
  SetStyleBack 5, &HFFFFFF
  SetStyleItalic 6, True
  SetStyleFore 6, &H7F007F
  SetStyleBack 6, &HFFFFFF
  SetStyleItalic 7, True
  SetStyleUnderline 7, True
  SetStyleFore 7, &H7F007F
  SetStyleBack 7, &HE0C0E0
  SetStyleFore 8, &H804080
  SetStyleBack 8, &HFFFFFF
  SetStyleFore 9, &H7F7F
  SetStyleBack 9, &HFFFFFF
  SetStyleFore 10, &H7F7F00
  SetStyleBack 10, &HFFFFFF
  SetStyleFore 11, &H20000000
  SetStyleBack 11, &HFFFFFF
  SetStyleFore 12, &H0
  SetStyleBack 12, &HE0C0E0
  SetStyleFont 12, "Courier New"
  SetStyleFore 13, &H7F00
  SetStyleBack 13, &HE0FFE0
  SetStyleFont 13, "Courier New"
  SetStyleFore 14, &H3F7F3F
  SetStyleBack 14, &HE0F0FF
  SetStyleFont 14, "Courier New"
  SetStyleFore 15, &H7F00
  SetStyleBack 15, &HFFFFFF
  SetStyleFont 15, "Comic Sans MS"
  SetStyleSize 15, 9
  SetStyleFore 16, &HB00040
  SetStyleBack 16, &HFFFFFF
  SetStyleFore 17, &H3060A0
  SetStyleBack 17, &HFFFFFF
  SetStyleFont 17, "Comic Sans MS"
  SetStyleFore 18, &H804020
  SetStyleBack 18, &HFFFFFF
  SetStyleFont 18, "Comic Sans MS"
  SetStyleSize 18, 9
  SetStyleFore 32, &H0
  SetStyleBack 32, &HFFFFFF
  SetStyleFore 33, &HBD7173
  SetStyleBack 33, &HEFEFEF
  SetStyleFont 33, "Courier New"
  SetStyleSize 33, 8
  SetStyleBold 34, True
  SetStyleFore 34, &HFF0000
  SetStyleBack 34, &HFFFFFF
  SetStyleBold 35, True
  SetStyleFore 35, &HFF
  SetStyleBack 35, &HFFFFFF
  SetStyleFore 37, &H0
  SetStyleBack 37, &HFFFFFF
  SetKeywords 0, "asm auto bool break case catch char class const const_cast continue default delete do double dynamic_cast else enum explicit export extern false float for friend goto if inline int long mutable namespace new operator private protected public register rei"
  SetKeywords 2, "a addindex addtogroup anchor arg attention author b brief bug c class code date def defgroup deprecated dontinclude e em endcode endhtmlonly endif endlatexonly endlink endverbatim enum example exception f$ f[ f] file fn hideinitializer htmlinclude htmlon"
  SetLexer 3
End Sub

Private Sub SetHighlightPascal()

  SetStyleFore 0, &H0
  SetStyleBack 0, &HFFFFFF
  SetStyleFore 1, &H7F00
  SetStyleBack 1, &HFFFFFF
  SetStyleFont 1, "Comic Sans MS"
  SetStyleSize 1, 9
  SetStyleFore 2, &H7F00
  SetStyleBack 2, &HFFFFFF
  SetStyleFont 2, "Comic Sans MS"
  SetStyleSize 2, 9
  SetStyleFore 3, &H7F7F7F
  SetStyleBack 3, &HFFFFFF
  SetStyleFont 3, "Comic Sans MS"
  SetStyleSize 3, 9
  SetStyleFore 4, &H7F7F00
  SetStyleBack 4, &HFFFFFF
  SetStyleBold 5, True
  SetStyleFore 5, &H7F0000
  SetStyleBack 5, &HFFFFFF
  SetStyleItalic 6, True
  SetStyleFore 6, &H7F7F00
  SetStyleBack 6, &HFFFFFF
  SetStyleItalic 7, True
  SetStyleFore 7, &H7F7F00
  SetStyleBack 7, &HFFFFFF
  SetStyleBold 8, True
  SetStyleFore 8, &H7F7F
  SetStyleBack 8, &HFFFFFF
  SetStyleFore 9, &H7F7F00
  SetStyleBack 9, &HFFFFFF
  SetStyleFore 10, &H7F7F00
  SetStyleBack 10, &HFFFFFF
  SetStyleFore 32, &H0
  SetStyleBack 32, &HFFFFFF
  SetStyleFore 33, &HBD7173
  SetStyleBack 33, &HEFEFEF
  SetStyleFont 33, "Courier New"
  SetStyleSize 33, 8
  SetStyleBold 34, True
  SetStyleFore 34, &HFF0000
  SetStyleBack 34, &HFFFFFF
  SetStyleBold 35, True
  SetStyleFore 35, &HFF
  SetStyleBack 35, &HFFFFFF
  SetStyleFore 37, &H0
  SetStyleBack 37, &HFFFFFF
  SetKeywords 0, "program const type var begin end array set packed record string if then else while for to downto do with repeat until case of goto exit label procedure function nil file and or not xor div mod unit uses implementation interface external asm inline object"
  SetKeywords 1, "write read default public protected private property published stored"
  SetLexer 18
End Sub
Public Sub SetFocus()
  ' Set focus to the Scintilla handle
  SetFocusEx SCI
End Sub

Public Sub SetMarginWidth(enMargin As enMargins, lWidth As Long)
  SendEditor SCI_SETMARGINWIDTHN, enMargin, lWidth
End Sub

Public Sub ClearCommandKeys()
  ' This function is here for situations where you might want to use it
  ' but it removes ALL command keys including keys such as left, up, down, etc,
  ' not just ctrl+ sets.
  SendEditor SCI_CLEARALLCMDKEYS
End Sub

Public Sub SetContextMenu(ByVal bVal As Boolean)
  SendEditor SCI_USEPOPUP, CLng(bVal), 0
End Sub

Public Function GetReadOnly() As Boolean
  GetReadOnly = SendEditor(SCI_GETREADONLY)
End Function

Public Function SetFixedFont(ByVal strFont As String, lSize As Long)
  Dim i As Long
  For i = 0 To 127
    SetStyleFont i, strFont
    SetStyleSize i, lSize
  Next i
End Function

Public Sub SetLineBreak(ByVal eolVal As EOL)
  SendEditor SCI_SETEOLMODE, CLng(eolVal)
End Sub


Private Sub WriteToFile(strFile As String, strData As String)
  On Error GoTo eHandle
  Dim i As Long
  Dim l As Long
  Dim hFile As Long
  Dim bByte() As Byte
  Str2Byte strData, bByte()
  l = UBound(bByte())
  'Convert the string to the proper EOL style to match scintilla
  For i = 0 To l
    If LineBreak = SC_EOL_CR Then
       If Chr(bByte(i)) = vbLf Then
         bByte(i) = 0
       End If
    End If
    If LineBreak = SC_EOL_LF Then
      If Chr(bByte(i)) = vbCr Then
        bByte(i) = 0
      End If
    End If
  Next i
  hFile = CreateFile(strFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, 0, 0&)
  WriteFile hFile, bByte(0), l, 0, ByVal 0&
  CloseHandle hFile
  Exit Sub
eHandle:
  ' Just in case anything happens let's close the handle
  CloseHandle hFile
End Sub

Private Function GetFile(strFilePath As String, Optional bolAsString = True) As String

  Dim arrFileMain() As Byte
  Dim arrFileBuffer() As Byte
  Dim lngAllBytes As Long
  Dim lngSize As Long, lngRet As Long
  Dim lngFileHandle As Long
  Dim ofData As OFSTRUCT
  Const lngMaxSizeForOneStep = 10000000
    'Prepare Arrays ==========================================================
    ReDim arrFileMain(0)
    ReDim arrFileBuffer(lngMaxSizeForOneStep)

    'Open the two files
    lngFileHandle = OpenFile(strFilePath, ofData, OF_READ)

    'Get the file size
    lngSize = GetFileSize(lngFileHandle, 0)
    Do While Not UBound(arrFileMain) = lngSize - 1
        If lngSize = 0 Then Exit Function

        'Redim Array to fit a smaller file
        lngAllBytes = UBound(arrFileMain)
        If lngSize - lngAllBytes < lngMaxSizeForOneStep Then ReDim arrFileBuffer(lngSize - lngAllBytes - 2)

        'Read from the file
        ReadFile lngFileHandle, arrFileBuffer(0), UBound(arrFileBuffer) + 1, lngRet, ByVal 0&

        'Calculate Buffer's position in Main Array
        If lngAllBytes > 0 Then lngAllBytes = lngAllBytes + 1

        'Make place for the Buffer in the Main Array
        ReDim Preserve arrFileMain(lngAllBytes + UBound(arrFileBuffer))

        'Put Buffer at end of Main Array
        MemCopy arrFileMain(lngAllBytes), arrFileBuffer(0), UBound(arrFileBuffer) + 1

        DoEvents

    Loop

    'Close the file
    CloseHandle lngFileHandle
    ReDim arrFileBuffer(0)

    'Convert Main Array to String
    GetFile = StrConv(arrFileMain(), vbUnicode)

End Function

Private Sub Str2Byte(sInput As String, bOutput() As Byte)
  ' This function is used to convert strings to bytes
  ' This comes in handy for saving the file.  It's also
  ' useful when dealing with certain things related to
  ' sending info to Scintilla
  
  Dim i As Long
  ReDim bOutput(Len(sInput) - 1)

  For i = 0 To Len(sInput) - 1
    bOutput(i) = Asc(Mid(sInput, i + 1, 1))
  Next i
End Sub

Private Sub SetWordWrap(ByVal bVal As Boolean)
  Dim lTmp As Long
  If bVal = True Then
    lTmp = 1
  Else
    lTmp = 0
  End If
  SendEditor SCI_SETWRAPMODE, lTmp
End Sub

Public Sub ShowCallTip(strVal As String)
  Dim bByte() As Byte
  Str2Byte strVal, bByte
  Call SendEditor(SCI_CALLTIPSHOW, GetCurPos, VarPtr(bByte(0)))
End Sub

Public Function GetCharAt(lPos As Long) As String
  GetCharAt = Chr(SendMessage(SCI, SCI_GETCHARAT, lPos, CLng(0)))
End Function

Private Function ToLastSpaceCount() As Long
  ' This function will figure out how many characters there are in the currently
  ' selected word.  It gets the line text, finds the position of the caret in
  ' the line text, then converts the line to a byte array to do a faster compare
  ' till it reaches something not interpreted as a letter IE a space or a
  ' line break.  This is kind of overly complex but seems to be faster overall
  
  Dim l As Long, i As Long, current As Long, pos As Long, startWord As Long, iHold As Long
  Dim str As String, bByte() As Byte, strTmp As String
  Dim line As String
  line = GetLine(GetCurLine)
  current = GetCaretInLine
  startWord = current
   
  Str2Byte line, bByte()
  
  iHold = 0
  While (startWord > 0) And InStr(1, CallTipWordCharacters, strTmp) > 0
    startWord = startWord - 1
    iHold = iHold + 1
    If startWord >= 0 Then
      strTmp = Chr(bByte(startWord))
    End If
  Wend
  If strTmp = " " Then iHold = iHold - 1
  ToLastSpaceCount = iHold
End Function

Public Function GetCaretInLine() As Long
  Dim caret As Long, lineStart As Long, line As Long
  caret = GetCurPos
  line = GetCurLine
  lineStart = PositionFromLine(line)
  GetCaretInLine = caret - lineStart
End Function

Public Sub ShowAutoComplete(strVal As String)
  Dim i As Long
  i = ToLastSpaceCount
  SendMessageString SCI, SCI_AUTOCSHOW, i, strVal
End Sub

Public Function Byte2Str(bVal() As Byte) As String
  Dim i As Long
  For i = 0 To UBound(bVal())
    Byte2Str = Byte2Str & Chr(bVal(i))
  Next i
End Function


Public Function GetLine(lLine As Long) As String
  Dim Txt As String
  Dim lLength As Long
  Dim bByte() As Byte
  lLength = SendMessage(SCI, SCI_LINELENGTH, GetCurLine, 0)
  ReDim bByte(0 To lLength)
  SendMessage SCI, SCI_GETLINE, lLine, VarPtr(bByte(0))
  Txt = Byte2Str(bByte())
  GetLine = Txt
End Function

Public Function GetZoom() As Long
  GetZoom = SendEditor(SCI_GETZOOM)
End Function

Public Sub SetZoom(lval As Long)
  SendEditor SCI_SETZOOM, lval
End Sub

Public Sub ZoomIn()
  Dim i As Long
  i = GetZoom
  i = i + 1
  SetZoom i
End Sub

Public Sub ZoomOut()
  Dim i As Long
  i = GetZoom
  i = i - 1
  SetZoom i
End Sub

Public Sub SetScrollWidth(lval As Long)
  SendEditor SCI_SETSCROLLWIDTH, lval
End Sub

Public Function GetScrollWidth() As Long
  GetScrollWidth = SendEditor(SCI_GETSCROLLWIDTH)
End Function

Public Function SetAutoCompleteIgnoreCase(bVal As Boolean)
  Dim i As Long
  If bVal = True Then
    i = 1
  Else
    i = 0
  End If
  SendEditor SCI_AUTOCSETIGNORECASE, CLng(i)
End Function

Public Sub SetViewWhiteSpace(bVal As Boolean)
  Dim i As Long
  If bVal = True Then
    i = 1
  Else
    i = 0
  End If
  SendEditor SCI_SETVIEWWS, i
End Sub

Public Function GetViewWhiteSpace() As Boolean
  GetViewWhiteSpace = SendEditor(SCI_GETVIEWWS)
End Function

Public Function FindText(txttofind As String, Optional FindReverse As Boolean = False, Optional ByVal findinrng As Boolean, Optional WrapDocument As Boolean = True, Optional CaseSensative As Boolean = False, Optional WordStart As Boolean = False, Optional WholeWord As Boolean = False, Optional RegExp As Boolean = False) As Boolean
  Dim lval As Long, Find As Long
  ' Sending a null string to scintilla for the find text willc ause errors!
  If txttofind = "" Then Exit Function
  lval = 0
  If CaseSensative Then
    lval = lval Or SCFIND_MATCHCASE
  End If
  If WordStart Then
    lval = lval Or SCFIND_WORDSTART
  End If
  If WholeWord Then
    lval = lval Or SCFIND_WHOLEWORD
  End If
  If RegExp Then
    lval = lval Or SCFIND_REGEXP
  End If
  Dim targetstart As Long, targetend As Long, pos As Long
    Call SendEditor(SCI_SETSEARCHFLAGS, lval)
    If findinrng Then
        targetstart = SendMessage(SCI, SCI_GETSELECTIONSTART, CLng(0), CLng(0))
        targetend = SendMessage(SCI, SCI_GETSELECTIONEND, CLng(0), CLng(0))
    Else
      If FindReverse = False Then
        targetstart = SendMessage(SCI, SCI_GETSELECTIONEND, 0, 0)
        targetend = Len(Text)
      Else
        targetstart = SendMessage(SCI, SCI_GETSELECTIONSTART, 0, 0)
        targetend = 0
      End If
    End If
    ' Creamos una región de búsqueda (que puede ser el texto completo)
    Call SendEditor(SCI_SETTARGETSTART, targetstart)
    Call SendEditor(SCI_SETTARGETEND, targetend)
    Find = SendMessageString(SCI, SCI_SEARCHINTARGET, Len(txttofind), txttofind)
    ' Seleccionamos lo que se ha encontrado
    If Find > -1 Then

        targetstart = SendMessage(SCI, SCI_GETTARGETSTART, CLng(0), CLng(0))
        targetend = SendMessage(SCI, SCI_GETTARGETEND, CLng(0), CLng(0))
        SetSel targetstart, targetend
    Else
      If WrapDocument Then
        If FindReverse = False Then
          targetstart = 0
          targetend = Len(Text)
        Else
          targetstart = Len(Text)
          targetend = 0
        End If
        Call SendEditor(SCI_SETTARGETSTART, targetstart)
        Call SendEditor(SCI_SETTARGETEND, targetend)
        Find = SendMessageString(SCI, SCI_SEARCHINTARGET, Len(txttofind), txttofind)
        If Find > -1 Then
          targetstart = SendMessage(SCI, SCI_GETTARGETSTART, CLng(0), CLng(0))
          targetend = SendMessage(SCI, SCI_GETTARGETEND, CLng(0), CLng(0))
          SetSel targetstart, targetend
        End If
      End If
    End If
    
  ' A find has been performed so now FindNext will work.
  bFindEvent = True
  If Find > -1 Then
    FindText = True
  Else
    FindText = False
  End If
    
  ' Set the info that we've used so we findnext can send the same thing
  ' out if called.
  
    bWrap = WrapDocument
    bCase = CaseSensative
    bWholeWord = WholeWord
    bRegEx = RegExp
    bWordStart = WordStart
    bFindInRange = findinrng
    bFindReverse = FindReverse
    strFind = txttofind
  
End Function

Public Function FindNext() As Boolean
  'If no find events have occurred exit this sub or it may cause errors.
  If bFindEvent = False Then Exit Function
  FindNext = FindText(strFind, False, bFindInRange, bWrap, bCase, bWordStart, bWholeWord, bRegEx)
End Function

Public Function FindPrev() As Boolean
  If bFindEvent = False Then Exit Function
  FindPrev = FindText(strFind, True, bFindInRange, bWrap, bCase, bWordStart, bWholeWord, bRegEx)
End Function

Public Function GetStyleAt(lPos As Long) As Long
  GetStyleAt = SendMessage(SCI, SCI_GETSTYLEAT, lPos, 0)
End Function

Public Function LoadAPIFile(strFile As String)
  ' This function will load an api file for calltips.
  Dim iFile As Integer, str As String, i As Integer
  iFile = FreeFile
  Erase APIStrings  'Clear the old array
  i = 0
  APIStrings = Split(GetFile(strFile), vbCr)
  For i = 0 To UBound(APIStrings) - 1
    APIStrings(i) = Replace(APIStrings(i), Chr(13), "")
  Next i
  APIStringLoaded = True
'  Open strFile For Input As #iFile
'    Do While Not EOF(iFile)
'      Input #iFile, str
'      ReDim Preserve APIStrings(0 To (UBound(APIStrings) + 1))
'      APIStrings(i) = str
'    Loop
'  Close #iFile
End Function

Public Function GetSelStart() As Long
  GetSelStart = SendEditor(SCI_GETSELECTIONSTART)
End Function

Public Function GetSelEnd() As Long
  GetSelEnd = SendEditor(SCI_GETSELECTIONEND)
End Function

Public Function CallTipActive() As Long
  CallTipActive = SendEditor(SCI_CALLTIPACTIVE)
End Function

Private Sub StartCallTip(ch As Long)
  ' This entire function is a bit of a hack.  It seems to work but it's very
  ' messy.  If anyone cleans it up please send me a new version so I can add
  ' it to this release.  Thanks :)
  Dim line As String, str As String, i As Integer, x As Integer
  Dim newstr As String, iPos As Integer, iPos2 As Integer, iStart As Long, iEnd As Long
  Dim iPos3 As Integer, iPos4 As Integer, str2 As String
  If APIStringLoaded = False Then Exit Sub
  If UBound(APIStrings) = 0 Then Exit Sub
  If CallTipActive = 0 Then
    If ch = Asc("(") Then
      line = GetLine(GetCurLine)
      x = GetCaretInLine
      str = Mid(line, 1, x)
      newstr = ""
      For i = x - 1 To 1 Step -1
        If InStr(1, CallTipWordCharacters, Mid(str, i, 1)) > 0 Then
          newstr = Mid(str, i, 1) & newstr
        Else
          Exit For
        End If
      Next i
      If Len(newstr) = 0 Then
        StopCallTip
        Exit Sub
      End If
      For i = 0 To UBound(APIStrings) - 1
        iPos2 = InStr(1, APIStrings(i), newstr)
        If iPos2 > 0 Then
          'if instr(1, CallTipWordCharacters, mid(line,
          If InStr(1, CallTipWordCharacters, Mid(APIStrings(i), iPos2 - 1, 1)) > 0 Then Exit Sub
          If InStr(1, CallTipWordCharacters, Mid(APIStrings(i), iPos2 + Len(newstr), 1)) > 0 Then Exit Sub
          ActiveCallTip = i
          CallTipPos = 0
          ShowCallTip Mid(APIStrings(i), 2, Len(APIStrings(i)) - 2)
          iPos = InStr(1, APIStrings(i), ",")
          If iPos > 0 Then
            iStart = Len(newstr) + 1
            iEnd = iPos - 1
            SetCallTipHighlight iStart, iEnd
        
            Exit Sub
          End If
        End If
      Next
    End If
  ElseIf ch = Asc(")") Then
    StopCallTip
  Else
    If ch = Asc("(") Then
      StopCallTip
      Exit Sub
    End If
  'ch = Asc(",") Then
    ' First determine where we are at within the ( and )
    Dim UA() As String
    line = GetLine(GetCurLine)
    x = GetCaretInLine
    iPos = InStrRev(line, "(", x)
    str = Mid(line, iPos + 1, x - iPos) 'Get the chunk of the string were in
    UA = Split(str, ",")
    iPos2 = UBound(UA)
    
    iPos3 = InStr(1, APIStrings(ActiveCallTip), "(")
    iPos4 = InStrRev(APIStrings(ActiveCallTip), ")")
    str2 = Mid$(APIStrings(ActiveCallTip), iPos3 + 1, iPos4 - iPos3 - 1)
    
    Erase UA
    UA = Split(str2, ",")
        
    iStart = 0
    iEnd = 0
    x = 0
    iPos3 = iPos3 + 1
    If iPos2 = 0 Then
      iStart = iPos3 - 2
      iEnd = Len(UA(0)) + iStart
    ElseIf iPos2 <= UBound(UA) Then
        For i = 1 To iPos2
          x = x + Len(UA(i))
        Next i
        iStart = x + iPos3 + (iPos2)
        iEnd = iStart + Len(UA(iPos2)) + (iPos2)
    Else
      ' In this case we have more commas than the function needs
      ' So we will just keep highlighting the last part of the
      ' tip.
      x = 0
      For i = 0 To UBound(UA) - 1
        x = x + Len(UA(i))
      Next
      iStart = x + iPos3 + UBound(UA) - 1
      iEnd = x + Len(UA(UBound(UA))) + iPos3 + (UBound(UA) - 1)
    End If
    SetCallTipHighlight iStart, iEnd
  End If
End Sub

Public Sub StopCallTip()
  SendEditor SCI_CALLTIPCANCEL
End Sub

Public Sub SetCallTipHighlight(lStart As Long, lEnd As Long)
  SendEditor SCI_CALLTIPSETHLT, lStart, lEnd
End Sub

Public Function GetCharAtLong(iPos As Long) As Long
  GetCharAtLong = SendMessage(SCI, SCI_GETCHARAT, iPos, 0)
End Function

Public Function ReplaceText(strSearchFor As String, strReplaceWith As String, Optional ReplaceAll As Boolean = False, Optional CaseSensative As Boolean = False, Optional WordStart As Boolean = False, Optional WholeWord As Boolean = False, Optional RegExp As Boolean = False) As Boolean
  bRepLng = True
  If FindText(strSearchFor, False, False, True, CaseSensative, WordStart, WholeWord) = True Then
    ReplaceSel strReplaceWith
    If ReplaceAll Then
      bRepAll = True
      Do Until FindText(strSearchFor, False, False, True, CaseSensative, WordStart, WholeWord) = False
        ReplaceSel strReplaceWith
      Loop
      bRepAll = False
    End If
  End If
  bRepLng = False
End Function

Public Sub DoFind()
  Load frmFind
  With frmFind
    If SelText <> "" Then .cmbFind.Text = SelText
    .Show vbModal, Owner
    If .DoWhat = 0 Then
      Exit Sub
    ElseIf .DoWhat = 1 Then
      FindText .cmbFind.Text, .optUp.value, False, .chkWrap.value, .chkCase.value, False, .chkWhole.value, .chkRegExp.value
    ElseIf .DoWhat = 2 Then
      'MarkAll .cmbFind.Text
      ' This will be in a future release
    End If
    Unload frmFind
  End With
  SetFocus
End Sub

Public Sub MarkerSet(iLine As Long, iMarkerNum As Long)
  SendEditor SCI_MARKERADD, iLine, iMarkerNum
End Sub

Public Function GetDocLen() As Long
  GetDocLen = SendMessage(SCI, SCI_GETLENGTH, 0, 0)
End Function

Public Function ReplaceAll(strSearchFor As String, strReplaceWith As String, Optional CaseSensative As Boolean = False, Optional WordStart As Boolean = False, Optional WholeWord As Boolean = False, Optional RegExp As Boolean = False) As Long
  ReplaceAll = 0
  Dim lval As Long
  Dim lenSearch As Long, lenReplace As Long
  Dim Find As Long
  If strSearchFor = "" Then Exit Function
  lval = 0
  If CaseSensative Then
    lval = lval Or SCFIND_MATCHCASE
  End If
  If WordStart Then
    lval = lval Or SCFIND_WORDSTART
  End If
  If WholeWord Then
    lval = lval Or SCFIND_WHOLEWORD
  End If
  If RegExp Then
    lval = lval Or SCFIND_REGEXP
  End If
  Dim targetstart As Long, targetend As Long, pos As Long, docLen As Long
  targetstart = 0
  docLen = GetLen
  lenSearch = Len(strSearchFor)
  lenReplace = Len(strReplaceWith)
  
  targetend = docLen
  Call SendEditor(SCI_SETSEARCHFLAGS, lval)
  Call SendEditor(SCI_SETTARGETSTART, targetstart)
  Call SendEditor(SCI_SETTARGETEND, targetend)
  Find = SendMessageString(SCI, SCI_SEARCHINTARGET, lenSearch, strSearchFor)
  Do Until Find = -1
    targetstart = SendMessage(SCI, SCI_GETTARGETSTART, CLng(0), CLng(0))
    targetend = SendMessage(SCI, SCI_GETTARGETEND, CLng(0), CLng(0))
    
    ReplaceTarget lenReplace, strReplaceWith
    targetstart = targetstart + lenReplace
    targetend = docLen
    ReplaceAll = ReplaceAll + 1
    Call SendEditor(SCI_SETTARGETSTART, targetstart)
    Call SendEditor(SCI_SETTARGETEND, targetend)
    Find = SendMessageString(SCI, SCI_SEARCHINTARGET, lenSearch, strSearchFor)
  Loop
End Function

Public Sub ReplaceTarget(lLen As Long, strRep As String)
  SendEditor SCI_REPLACETARGET, lLen, strRep
End Sub

Public Sub GotoLine(iLine As Long)
  Call SendEditor(SCI_ENSUREVISIBLEENFORCEPOLICY, iLine)
  SendEditor SCI_GOTOLINE, iLine
End Sub

Public Sub GotoLineColumn(iLine As Long, iCol As Long)
  Dim i As Long
  i = SendEditor(SCI_FINDCOLUMN, iLine, iCol)
  SetSel i, i
  'i = i + iCol
  'SendEditor
End Sub

Public Sub DoReplace()
  Load frmReplace
  With frmReplace
    Set .cScintilla = Me
    If SelText <> "" Then .cmbFind.Text = SelText
    .Show , Owner
  End With
End Sub

Public Sub DoGoto()
  Load frmGoto
  Dim iLine As Long, iCol As Long
  With frmGoto
    .lblCurLine.Caption = "Current Line: " & GetCurLine
    .lblLineCount.Caption = "Last Line: " & GetLastLine
    .lblColumn.Caption = "Column: " & GetColumn
    .Show vbModal, Owner
    If .iWhatToDo = 1 Then
      If .txtLine.Text = "" Then .txtLine.Text = 1
      If .txtColumn.Text = "" Then .txtColumn.Text = 1
      iLine = .txtLine.Text
      iCol = .txtColumn.Text
      GotoLineColumn iLine - 1, iCol - 1
    End If
  End With
  Unload frmGoto
  SetFocus
End Sub

Private Sub RemoveHotKeys()
  ' This just removes some of the common hot keys that
  ' could cause scintilla to interfere with the application
  
  ClearCmdKey Asc("V") + LShift(SCMOD_CTRL, 16)
  ClearCmdKey Asc("C") + LShift(SCMOD_CTRL, 16)
  ClearCmdKey Asc("X") + LShift(SCMOD_CTRL, 16)
  ClearCmdKey Asc("Z") + LShift(SCMOD_CTRL, 16)
  ClearCmdKey Asc("Y") + LShift(SCMOD_CTRL, 16)
End Sub

Public Sub ClearCmdKey(lKeyDef As Long)
  SendEditor SCI_CLEARCMDKEY, lKeyDef, SCI_NULL
End Sub

Private Function ShiftLeft(ByVal value As Long, ByVal times As Long) As Long

    ' we need to create a mask of 1's corresponding to the
    ' times in VALUE that will be retained in the result
    Dim mask As Long, signBit As Long
    
    ' return zero if too many times
    If times >= 32 Then Exit Function
    ' return the value if zero times
    If times = 0 Then ShiftLeft = value: Exit Function
    
    ' this extracts the bit in Value that will become the sign bit
    mask = Power2(31 - times)
    ' this calculates the sign bit of the result
    signBit = CBool(value And mask) And &H80000000
    ' this clears all the most significant times,
    ' that would be lost anyway, and also clears the sign bit
    value = value And (mask - 1)
    ' do the shift to the left, without risking an overflow
    ' and then add the sign bit
    ShiftLeft = (value * Power2(times)) Or signBit
End Function

Private Function Power2(ByVal exponent As Long) As Long
    Static res(0 To 31) As Long
    Dim i As Long
    
    ' rule out errors
    If exponent < 0 Or exponent > 31 Then Err.Raise 5
    
    ' initialize the array at the first call
    If res(0) = 0 Then
        res(0) = 1
        For i = 1 To 30
            res(i) = res(i - 1) * 2
        Next
        ' this is a special case
        res(31) = &H80000000
    End If
    
    ' return the result
    Power2 = res(exponent)
        
End Function


Public Function Shift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long, ByVal lDirectionToShift As dcShiftDirection) As Long

    Const ksCallname As String = "Shift"
    On Error GoTo Procedure_Error
    Dim LShift As Long

    If lDirectionToShift Then 'shift left
        LShift = lValue * (2 ^ lNumberOfBitsToShift)
    Else 'shift right
        LShift = lValue \ (2 ^ lNumberOfBitsToShift)
    End If

    
Procedure_Exit:
    Shift = LShift
    Exit Function
    
Procedure_Error:
    Err.Raise Err.Number, ksCallname, Err.Description, Err.HelpFile, Err.HelpContext
    Resume Procedure_Exit
End Function

Public Function LShift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long) As Long

    Const ksCallname As String = "LShift"
    On Error GoTo Procedure_Error
    LShift = Shift(lValue, lNumberOfBitsToShift, lLeft)
    
Procedure_Exit:
    Exit Function
    
Procedure_Error:
    Err.Raise Err.Number, ksCallname, Err.Description, Err.HelpFile, Err.HelpContext
    Resume Procedure_Exit
End Function

Public Sub ClearUndoBuffer()
  SendEditor SCI_EMPTYUNDOBUFFER
End Sub

Public Sub SetStyleBits(iBit As Long)
  SendEditor SCI_SETSTYLEBITS, iBit
End Sub

Public Sub PrintDoc()
  PrintSCI SCI, GetLen, 1000, 1000, 1000, 1000
End Sub

Public Sub PrintSCI(sciHwnd As Long, txtLen As Long, LeftMarginWidth As Long, _
   TopMarginHeight, RightMarginWidth, BottomMarginHeight)
   Dim LeftOffset As Long, TopOffset As Long
   Dim LeftMargin As Long, TopMargin As Long
   Dim RightMargin As Long, BottomMargin As Long
   Dim fr As FormatRange
   Dim rcDrawTo As Rect
   Dim rcPage As Rect
   Dim TextLength As Long
   Dim NextCharPosition As Long
   Dim r As Long
   Dim PhysWidth As Long, PhysHeight As Long
   Dim PrintWidth As Long, PrintHeight As Long
   Dim ptDPI As POINTAPI, ptPage As POINTAPI
   Dim rectPhysMargins As Rect, rectMargins As Rect, rectSetup As Rect
   Printer.Print Space(1)
   Printer.ScaleMode = vbPixels


   ' Get the offsett to the printable area on the page in twips
   ptDPI.x = GetDeviceCaps(Printer.hdc, LOGPIXELSX)
   ptDPI.Y = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
   ptPage.x = GetDeviceCaps(Printer.hdc, PHYSICALWIDTH)
   ptPage.Y = GetDeviceCaps(Printer.hdc, PHYSICALHEIGHT)
   
   rectPhysMargins.Left = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
   rectPhysMargins.Top = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
   rectPhysMargins.Right = ptPage.x - GetDeviceCaps(Printer.hdc, HORZRES) - rectPhysMargins.Left
   rectPhysMargins.Bottom = ptPage.Y - GetDeviceCaps(Printer.hdc, VERTRES) - rectPhysMargins.Top
   
   rectSetup.Left = MulDiv(LeftMarginWidth, ptDPI.x, 1000)
   rectSetup.Top = MulDiv(TopMarginHeight, ptDPI.Y, 1000)
   rectSetup.Right = MulDiv(RightMarginWidth, ptDPI.x, 1000)
   rectSetup.Bottom = MulDiv(BottomMarginHeight, ptDPI.Y, 1000)
    
   rectMargins.Left = Max(rectPhysMargins.Left, rectSetup.Left)
   rectMargins.Top = Max(rectPhysMargins.Top, rectSetup.Top)
   rectMargins.Right = Max(rectPhysMargins.Right, rectSetup.Right)
   rectMargins.Bottom = Max(rectPhysMargins.Bottom, rectSetup.Bottom)
    
   ' Calculate the Left, Top, Right, and Bottom margins
   'LeftMargin = (LeftMarginWidth - LeftOffset) \ Printer.TwipsPerPixelX
   'TopMargin = (TopMarginHeight - TopOffset) \ Printer.TwipsPerPixelY
   'RightMargin = (((Printer.Width - RightMarginWidth) - LeftOffset) \ Printer.TwipsPerPixelX) + (LeftMargin + LeftOffset)
   'BottomMargin = (((Printer.Height - BottomMarginHeight) - TopOffset) \ Printer.TwipsPerPixelY) + (TopMargin + TopOffset)

   ' Set printable area rect
   'rcPage.Left = 0
   'rcPage.Top = 0
   'rcPage.Right = Printer.ScaleWidth
   'rcPage.Bottom = Printer.ScaleHeight

   ' Set rect in which to print (relative to printable area)
   'rcDrawTo.Left = LeftMargin
   'rcDrawTo.Top = TopMargin
   'rcDrawTo.Right = RightMargin
   'rcDrawTo.Bottom = BottomMargin
   
   'rcPage = rcDrawTo
   ' Set up the print instructions
   
  fr.rc.Left = rectMargins.Left - rectPhysMargins.Left
  fr.rc.Top = rectMargins.Top - rectPhysMargins.Top
  fr.rc.Right = ptPage.x - rectMargins.Right - rectPhysMargins.Left
  fr.rc.Bottom = ptPage.Y - rectMargins.Bottom - rectPhysMargins.Top
  fr.rcPage.Left = 0
  fr.rcPage.Top = 0
  fr.rcPage.Right = ptPage.x - rectPhysMargins.Left - rectPhysMargins.Right - 1
  fr.rcPage.Bottom = ptPage.Y - rectPhysMargins.Top - rectPhysMargins.Bottom - 1

   fr.chrg.cpMin = 0           ' Indicate start of text through
'
   fr.chrg.cpMax = txtLen          ' end of the text

   ' Get length of text in RTF
   TextLength = txtLen
   'NextCharPosition = SendMessage2(SCIHwnd, SCI_FORMATRANGE, True, fr)
   NextCharPosition = 0
   Do
      'Printer.NewPage                  ' Move on to next page
      Printer.Print Space(1) ' Re-initialize hDC
      fr.hdc = Printer.hdc
      fr.hdcTarget = Printer.hdc
      fr.chrg.cpMin = NextCharPosition ' Starting position for next page
      fr.chrg.cpMax = txtLen
      NextCharPosition = SendMessage2(sciHwnd, SCI_FORMATRANGE, True, fr)
      'MsgBox NextCharPosition & "|" & txtLen
      If NextCharPosition >= txtLen Then Exit Do
      Printer.NewPage
      
   Loop
'
'   ' Commit the print job
   Printer.EndDoc
'
'   ' Allow the RTF to free up memory
   r = SendMessage2(sciHwnd, SCI_FORMATRANGE, False, ByVal CLng(0))
End Sub

Private Function Max(a As Long, b As Long) As Long
  If a > b Then
    Max = a
  Else
    Max = b
  End If
End Function
