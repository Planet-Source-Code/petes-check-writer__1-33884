Attribute VB_Name = "modWin32Api"
'
' Declarations for miscellaneous Win32 API functions
'
Option Explicit

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal nLen As Long)
Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function IsBadStringPtrp Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long
Declare Function lstrlenp Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_WNDPROC = (-4)

Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10

Public Const WM_CONTEXTMENU = &H7B
Public Const WM_HELP = &H53
Public Const WM_NCDESTROY = &H82
Public Const WM_NOTIFY = &H4E
Public Const WM_TCARD = &H52

Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000

Public Const WS_EX_CONTEXTHELP = &H400

'Type POINTAPI
'    x As Long
'    y As Long
'End Type

Type HELPINFO
    cbSize As Long
    iContextType As Long
    iCtrlId As Long
    hItemHandle As Long
    dwContextId As Long
    MousePos As POINTAPI
End Type

Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

