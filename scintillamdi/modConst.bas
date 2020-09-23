Attribute VB_Name = "modConst"
Option Explicit

Public Const WM_NOTIFY = &H4E

Public Const PHYSICALWIDTH = 110 '  Physical Width in device units
Public Const PHYSICALHEIGHT = 111 '  Physical Height in device units


Type NMHDR
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type

Public Type SCNotification
    NotifyHeader As NMHDR
    position As Long
    ch As Long
    modifiers As Long
    modificationType As Long
    Text As Long
    length As Long
    linesAdded As Long
    Message As Long
    wParam As Long
    lParam As Long
    line As Long
    foldLevelNow As Long
    foldLevelPrev As Long
    margin As Long
    listType As Long
    X As Long
    Y As Long
End Type

Public Enum EOL
    SC_EOL_CRLF = 0                     ' CR + LF
    SC_EOL_CR = 1                       ' CR
    SC_EOL_LF = 2                       ' LF
End Enum

Public Const CB_FINDSTRING = &H14C

