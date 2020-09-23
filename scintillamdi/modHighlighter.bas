Attribute VB_Name = "modHighlighter"
Option Explicit

'+---------------------------------------------------------------------------+
'| modHighlighter.bas                                                        |
'+---------------------------------------------------------------------------+
'| This is a basic module to provide very basic highlighter loading support. |
'| In reality I wouldn't really recomend using this as a basis for your      |
'| editor but it should give you some idea's.  The biggest reason I did not  |
'| want to bundle the code to read highlighter files into the class itself   |
'| is for performance reasons.  With this setup you can load the files one   |
'| time, and then just set each editor.  For the demo application this is a  |
'| fairly useless feature but if your dealing with a MDI application it's    |
'| going to make a world of difference.  If it was bundled directly into the |
'| class quite litterly every document you create would load every single    |
'| file.  That would be very poor use of system resources :)                 |
'+---------------------------------------------------------------------------+

Private Type Highlighter
  StyleBold(127) As Long
  StyleItalic(127) As Long
  StyleUnderline(127) As Long
  StyleVisible(127) As Long
  StyleEOLFilled(127) As Long
  StyleFore(127) As Long
  StyleBack(127) As Long
  StyleSize(127) As Long
  StyleFont(127) As String
  StyleName(127) As String
  Keywords(7) As String
  strFilter As String
  strComment As String
  strName As String
  iLang As Long
End Type

Private HCount As Integer

Public Highlighters() As Highlighter ' Make it publicly exposed so the app can
                                     ' read off name's for menu's or such
Private CurHigh As Integer

Private sBuffer As String
Private Const ciIncriment As Integer = 15000
Private lOffset As Long

Public Sub ReInit()
sBuffer = ""
lOffset = 0
End Sub

Public Function GetString() As String
GetString = Left$(sBuffer, lOffset)
sBuffer = ""  'reset
End Function

'This function lets you assign a string to the concating buffer.
Public Sub SetString(ByRef Source As String)
sBuffer = Source & String$(ciIncriment, 0)
End Sub

Public Sub SConcat(ByRef Source As String)
Dim lBufferLen As Long
lBufferLen = Len(Source)
'Allocate more space in buffer if needed
If (lOffset + lBufferLen) >= Len(sBuffer) Then
   If lBufferLen > lOffset Then
      sBuffer = sBuffer & String$(lBufferLen, 0)
   Else
      sBuffer = sBuffer & String$(ciIncriment, 0)
   End If
End If
Mid$(sBuffer, lOffset + 1, lBufferLen) = Source
lOffset = lOffset + lBufferLen
End Sub

Private Function FindHighlighter(strLangName As String) As Integer
  Dim i As Integer
   For i = 0 To UBound(Highlighters) - 1
    'MsgBox """" & UCase(strLangName) & """" & "|" & """" & Highlighters(i).strName & """"""
    If UCase(Highlighters(i).strName) = UCase(strLangName) Then
      FindHighlighter = i
      Exit Function
    End If
  Next i
End Function

Public Function SetHighlighter(cScintilla As clsScintilla, strHighlighter As String)
  Dim i As Long, X As Integer
  X = FindHighlighter(strHighlighter)
  cScintilla.StyleClearALL
  cScintilla.StartStyle
  For i = 0 To 127
    cScintilla.SetStyleBold i, Highlighters(X).StyleBold(i)
    cScintilla.SetStyleItalic i, Highlighters(X).StyleItalic(i)
    cScintilla.SetStyleUnderline i, Highlighters(X).StyleUnderline(i)
    cScintilla.SetStyleVisible i, Highlighters(X).StyleVisible(i)
    If Highlighters(X).StyleFont(i) <> "" Then cScintilla.SetStyleFont i, Highlighters(X).StyleFont(i)
    cScintilla.SetStyleFore i, Highlighters(X).StyleFore(i)
    cScintilla.SetStyleBack i, Highlighters(X).StyleBack(i)
    cScintilla.SetStyleSize i, Highlighters(X).StyleSize(i)
    cScintilla.SetStyleEOLFilled i, Highlighters(X).StyleEOLFilled(i)
  Next i
  For i = 0 To 7
    If Highlighters(X).Keywords(i) <> "" Then cScintilla.SetKeywords i, Highlighters(X).Keywords(i)
  Next i
  If LCase(strHighlighter) = "html" Then
    cScintilla.SetStyleBits 7
  Else
    cScintilla.SetStyleBits 5
  End If
  cScintilla.SetLexer Highlighters(X).iLang
  cScintilla.Colourise
  CurHigh = X
End Function

Public Function LoadHighlighter(strFile As String)
  Dim d() As String
  Dim s As String, i As Long, m As Long
  Dim l As Long
  ReDim Preserve Highlighters(0 To HCount + 1)
  For i = 0 To 7
    s = ReadINI("data", "Keywords[" & i & "]", strFile)
    Highlighters(HCount).Keywords(i) = s
  Next i
  
  For i = 0 To 127
    s = ReadINI("data", "style[" & i & "]", strFile)
    Highlighters(HCount).StyleBold(i) = 0
    Highlighters(HCount).StyleUnderline(i) = 0
    Highlighters(HCount).StyleItalic(i) = 0
    Highlighters(HCount).StyleVisible(i) = 0
    Highlighters(HCount).StyleEOLFilled(i) = 0
    Highlighters(HCount).StyleFont(i) = "Courier New"
    Highlighters(HCount).StyleSize(i) = 10
    Highlighters(HCount).StyleFore(i) = vbBlack
    Highlighters(HCount).StyleBack(i) = vbWhite
    Highlighters(HCount).StyleName(i) = ""
    If s <> "" Then
      d = Split(s, ":")
      If UCase(d(0)) = "B" Then Highlighters(HCount).StyleBold(i) = 1
      If UCase(d(1)) = "I" Then Highlighters(HCount).StyleItalic(i) = 1
      If UCase(d(2)) = "U" Then Highlighters(HCount).StyleUnderline(i) = 1
      If UCase(d(3)) = "V" Then Highlighters(HCount).StyleVisible(i) = 1
      If UCase(d(5)) = "E" Then Highlighters(HCount).StyleEOLFilled(i) = 1
      If UCase(d(7)) <> "" Then Highlighters(HCount).StyleFont(i) = d(7)
      If UCase(d(8)) <> "0" Then Highlighters(HCount).StyleSize(i) = Int(d(8))
      If UCase(d(9)) <> "" Then Highlighters(HCount).StyleFore(i) = Int(d(9))
      If UCase(d(10)) <> "" Then Highlighters(HCount).StyleBack(i) = Int(d(10))
      If UCase(d(11)) <> "" Then Highlighters(HCount).StyleName(i) = d(10)
      Erase d
    End If
  Next i
  s = ReadINI("data", "Language", strFile)
  If s <> "" Then Highlighters(HCount).iLang = Int(s)
  s = ReadINI("data", "Filter", strFile)
  If s <> "" And Right(s, 1) <> "|" Then s = s & "|"
  If s <> "" Then Highlighters(HCount).strFilter = s
  s = ReadINI("data", "LangName", strFile)
  If s <> "" Then Highlighters(HCount).strName = s
  s = ReadINI("data", "SingleComment", strFile)
  If s <> "" Then Highlighters(HCount).strComment = s
  HCount = HCount + 1
  
End Function

Public Sub LoadDirectory(strDir As String)
  Dim str As String, i As Integer
  If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
  str = Dir(strDir & "\*.CHL")
  Do Until str = ""
    
    LoadHighlighter strDir & "\" & str
    str = Dir
  Loop
End Sub

Public Function GetExtension(sFileName As String) As String
    Dim lPos As Long
    lPos = InStrRev(sFileName, ".")
    If lPos = 0 Then
        GetExtension = " "
    Else
        GetExtension = LCase$(Right$(sFileName, Len(sFileName) - lPos))
    End If
End Function

Public Function SetHighlighterBasedOnExtension(file As String) As String
  Dim Extension As String, UA() As String, ClrExt As String, X As Long
  Extension = LCase$(Mid$(file, InStrRev(file, ".") + 1, Len(file) - InStrRev(file, ".")))
  For X = 0 To UBound(Highlighters)
    If InStr(1, Highlighters(X).strFilter, Extension) Then
      SetHighlighterBasedOnExtension = Highlighters(X).strName
    End If
  Next X
  Erase UA
End Function

Public Function ExportToHTML(strFile As String, cScintilla As clsScintilla)
  On Error Resume Next
  ' This function will output the source to HTML with the styling
  ' It is far from perfect and frankly it's slower than hell if you ask me
  ' It takes it a solid 7-8 seconds to output this file (modHighlighter.bas)
  ' So if anyone can think of ways to improve it's speed.  At least its
  ' better than what it initially was (about 19 seconds for this file)
  ' thanks to a simple concatation function and comparing the long value's
  ' of the characters in question instead of a string to string comparison.
  ' but otherwise still slow :)
  Dim iLen As Long
  Dim strOutput As String
  Dim strCSS As String
  Dim lPrevStyle As Long
  Dim lStyle As Long
  Dim style(127) As Boolean
  Dim prevStyle As Long
  Dim curStyle As Long
  Dim nextStyle As Long
  Dim i As Long
  Dim strTotal As String
  Dim strStyle As String
  For i = 0 To 127
    style(i) = False
  Next i
  For i = 0 To Len(cScintilla.GetText)
    lStyle = cScintilla.GetStyleAt(i)
    style(lStyle) = True
  Next
  strCSS = ""
  strTotal = "<HTML>" & vbCrLf & "  <HEAD>" & vbCrLf & "    <Meta Generator=" & """" & "CScintilla Class (http://www.ceditmx.com)" & """" & ">" & vbCrLf
  strCSS = "<style type=" & """" & "text/css" & """" & ">" & vbCrLf
  For i = 0 To 127
    If style(i) = True Then
      With Highlighters(CurHigh)
        strCSS = strCSS & ".c" & i & " {" & vbCrLf
        If .StyleFont(i) <> "" Then
          strCSS = strCSS & "font-family: " & "'" & .StyleFont(i) & "'" & ";" & vbCrLf
        End If
        If .StyleFore(i) <> 0 Then
          strCSS = strCSS & "color: " & DectoHex(.StyleFore(i)) & ";" & vbCrLf
        End If
        If .StyleBack(i) <> 0 Then
          strCSS = strCSS & "background: " & DectoHex(.StyleBack(i)) & ";" & vbCrLf
        End If
        If .StyleSize(i) <> 0 Then
          strCSS = strCSS & "font-size: " & .StyleSize(i) & "pt" & ";" & vbCrLf
        End If
        If .StyleBold(i) = 0 Then
          strCSS = strCSS & "font-weight: 400;" & vbCrLf
        Else
          strCSS = strCSS & "font-weight: 700;" & vbCrLf
        End If
        strStyle = ""
        If .StyleItalic(i) <> 0 Then
          strStyle = "text-decoration: italic;"
        End If
        If .StyleUnderline(i) <> 0 Then
          If strStyle = "" Then
            strStyle = "text-decoration: underline;"
          Else
            strStyle = strStyle & ", underline;"
          End If
        End If
        If strStyle <> "" Then
          strCSS = strCSS & strStyle & vbCrLf
        End If
        strCSS = strCSS & "}" & vbCrLf
      End With
    End If
  Next i
  strCSS = strCSS & "</style>" & vbCrLf
  strTotal = strTotal & strCSS
  strTotal = strTotal & "  </HEAD>" & vbCrLf & "  <BODY BGCOLOR=#FFFFFF TEXT=#000000>"
  strOutput = ""
  sBuffer = ""
  iLen = cScintilla.GetLen
  For i = 0 To iLen
    curStyle = cScintilla.GetStyleAt(i)
    If (i + 1) < iLen Then
      nextStyle = cScintilla.GetStyleAt(i + 1)
    End If
    If curStyle <> prevStyle Then
      SConcat "<span class=c" & curStyle & ">"
      'strOutput = strOutput & "<span class=c" & curStyle & ">"
      If cScintilla.GetCharAtLong(i) <> 13 Then
        If cScintilla.GetCharAt(i) = " " Then
          'strOutput = strOutput & "&nbsp;"
          SConcat "&nbsp;"
        Else
          SConcat cScintilla.GetCharAt(i)
          'strOutput = strOutput & cScintilla.GetCharAt(i)
        End If
      Else
        SConcat "<BR>"
        'strOutput = strOutput & "<BR>"
      End If
      If i = iLen Or nextStyle <> curStyle Then
        SConcat "</span>"
        'strOutput = strOutput & "</span>"
      End If
    Else
      If cScintilla.GetCharAtLong(i) <> 13 Then
        If cScintilla.GetCharAt(i) = " " Then
          SConcat "&nbsp;"
          'strOutput = strOutput & "&nbsp;"
        Else
          SConcat cScintilla.GetCharAt(i)
          'strOutput = strOutput & cScintilla.GetCharAt(i)
        End If
      Else
        SConcat "<BR>"
        'strOutput = strOutput & "<BR>"
      End If
      If i = iLen Or nextStyle <> curStyle Then
        SConcat "</span>"
        'strOutput = strOutput & "</span>"
      End If
    End If
    prevStyle = curStyle
  Next i
  strOutput = GetString
  strTotal = strTotal & strOutput
  strTotal = strTotal & "  </BODY>" & vbCrLf & "</HTML>"
  i = FreeFile
  Open strFile For Output As #i
    Print #i, strTotal
  Close #i
End Function

Public Function DectoHex(lngColour As Long) As String

    '     *********
    Dim strColour As String
    'Convert decimal colour to hex
    strColour = Hex(lngColour)
    'Add leading zero's

    Do While Len(strColour) < 6
        strColour = "0" & strColour
    Loop

    'Reverse the bgr string pairs to rgb
    DectoHex = "#" & Right$(strColour, 2) & _
    Mid$(strColour, 3, 2) & _
    Left$(strColour, 2)
End Function

Function AddToString(St As String, ToAdd As String, Optional NumTimes As Long = 1) As String

    Dim LC As Long, StrLoc As Long
    AddToString = String$((Len(ToAdd) * NumTimes) + Len(St), 0) 'For CopyMemory() to work, the string must be padded With nulls to the desired size
    CopyMemory ByVal StrPtr(AddToString), ByVal StrPtr(St), LenB(St) 'Copy the original string to the return code
    StrLoc = StrPtr(AddToString) + LenB(St) 'Memory Location = Location of return code + size of original string
    'We use LenB() because strings are actua
    '     lly twice as long as Len() says when sto
    '     red in memory

    For LC = 1 To NumTimes
        CopyMemory ByVal StrLoc, ByVal StrPtr(ToAdd), LenB(ToAdd) 'Copy the source String to the return code
        StrLoc = StrLoc + LenB(ToAdd) 'Add the size of the String to the pointer


        DoEvents 'Comment this out If you don't plan To use huge repeat values, you'll Get a nice speed boost
        Next LC

 End Function
