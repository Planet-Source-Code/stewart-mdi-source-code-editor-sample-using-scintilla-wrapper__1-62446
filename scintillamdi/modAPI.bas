Attribute VB_Name = "modAPI"
Option Explicit
Public Const FILE_BEGIN = 0
Public Const OF_READ = &H0
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_NEW = 1
Public Const OPEN_EXISTING = 3
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OFS_MAXPATHNAME = 128
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const CREATE_ALWAYS = 2

Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

'+--------------------------------+
'| Begin API's                    |
'+--------------------------------+

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Integer
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Public Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, _
   lp As Any) As Long
Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal m As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal flags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function SendMessageStringString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal iMsg As Long, ByVal str1 As String, ByVal str1 As String) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

' The following functions handle reading and writing to a file
' The use the api CreateFile, WriteFile, OpenFile and ReadFile.
' This works more quickly than bothering with VB's built in file
' reading features.

Public Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type CharRange
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Public Type FormatRange
  hdc As Long       ' Actual DC to draw on
  hdcTarget As Long ' Target DC for determining text formatting
  rc As Rect        ' Region of the DC to draw to (in twips)
  rcPage As Rect    ' Region of the entire DC (page size) (in twips)
  chrg As CharRange ' Range of text to draw (see above declaration)
End Type

Public Const WM_USER As Long = &H400
Public Const EM_FORMATRANGE As Long = WM_USER + 57
Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Public Const PHYSICALOFFSETX As Long = 112
Public Const PHYSICALOFFSETY As Long = 113

Public Declare Function GetDeviceCaps Lib "gdi32" ( _
   ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
   (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
   ByVal lpOutput As Long, ByVal lpInitData As Long) As Long

Public Type POINTAPI
        X As Long
        y As Long
End Type
Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Public Declare Function DPtoLP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Const HORZRES = 8            '  Horizontal width in pixels
Public Const VERTRES = 10           '  Vertical width in pixels
Public Const VERTSIZE = 6           '  Vertical size in millimeters
Public Const HORZSIZE = 4           '  Horizontal size in millimeters

