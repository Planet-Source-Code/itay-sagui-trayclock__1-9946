Attribute VB_Name = "Module1"
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As _
  Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Type POINTAPI
  x As Long
  y As Long
End Type
Declare Function CreatePolygonRgn& Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long)
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
  ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const WM_NCLBUTTONDOWN = &HA1
Global Const WM_LBUTTONCLK = &H202
Global Const WM_LBUTTONDBLCLK = &H203
Global Const WM_RBUTTONUP = &H205
Global Const NIM_ADD = &H0
Global Const NIM_MODIFY = &H1
Global Const NIF_MESSAGE = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const WM_MOUSEMOVE = &H200
Global Const HTCAPTION = 2
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

