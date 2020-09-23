VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picClock 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1740
      Top             =   1350
   End
   Begin VB.PictureBox picSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3030
      Picture         =   "frmMain.frx":014A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   2190
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2745
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   2190
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TrayIcon As NOTIFYICONDATA
Dim Radius
Private Sub Form_Load()
  success% = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

' START OF TRAY
  TrayIcon.cbSize = Len(TrayIcon)
  TrayIcon.hwnd = Me.hwnd
  TrayIcon.uId = vbNull
  TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  TrayIcon.ucallbackMessage = WM_MOUSEMOVE
  TrayIcon.hIcon = picSrc.Picture
  TrayIcon.szTip = Time & Chr$(0)
  Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
  App.TaskVisible = False
' END OF TRAY

' START OF SHAPE
  Dim MyPoly() As POINTAPI
  Dim ZZ As Long
  Radius = 32
  ReDim Preserve MyPoly(0 To 359)
  MyPoly(0).x = Radius: MyPoly(0).y = 0
  picClock.Circle (Radius, Radius), Radius - 1, RGB(255, 127, 0)
  picClock.Circle (Radius - 1, Radius), Radius - 1, RGB(255, 127, 0)
  For i = 1 To 359
    MyPoly(i).x = Sin(i / 180 * 3.14) * Radius + Radius
    MyPoly(i).y = -Cos(i / 180 * 3.14) * Radius + Radius
    picClock.Line (MyPoly(i - 1).x - 1, MyPoly(i - 1).y - 1)-(MyPoly(i).x - 1, MyPoly(i).y - 1), RGB(255, 255, 0)
    picClock.Line (MyPoly(i - 1).x, MyPoly(i - 1).y + 1)-(MyPoly(i).x, MyPoly(i).y + 1), 127
  Next
  picClock.Line (MyPoly(359).x - 1, MyPoly(359).y - 1)-(MyPoly(0).x - 1, MyPoly(0).y - 1), RGB(255, 255, 0)
  picClock.Line (MyPoly(359).x, MyPoly(359).y + 1)-(MyPoly(0).x, MyPoly(0).y + 1), 127
  picClock.Picture = picClock.Image
  Me.Refresh
  i = 359
  ZZ = CreatePolygonRgn(MyPoly(0), i, 2)
  SetWindowRgn hwnd, ZZ, True
' END OF SHAPE
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Static Message As Long
  Static RR As Boolean
  Message = x ' / Screen.TwipsPerPixelX
  If RR = False Then
    RR = True
    Select Case Message
' Left click (This should bring up a dialog box)
      Case WM_LBUTTONCLK
' Left double click (This should bring up a dialog box)
      Case WM_LBUTTONDBLCLK
        TrayIcon.cbSize = Len(TrayIcon)
        TrayIcon.hwnd = Me.hwnd
        TrayIcon.uId = vbNull
        Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
        Me.Show
' Right button up (This should bring up a menu)
      Case WM_RBUTTONUP
'        PopupMenu mnuDropdown
    End Select
    RR = False
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  TrayIcon.cbSize = Len(TrayIcon)
  TrayIcon.hwnd = Me.hwnd
  TrayIcon.uId = vbNull
  Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
End Sub
Private Sub picClock_DblClick()
  Me.Visible = False
' START OF TRAY
  TrayIcon.cbSize = Len(TrayIcon)
  TrayIcon.hwnd = Me.hwnd
  TrayIcon.uId = vbNull
  TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  TrayIcon.ucallbackMessage = WM_MOUSEMOVE
  TrayIcon.hIcon = picSrc.Picture
  TrayIcon.szTip = Time & Chr$(0)
  Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
' END OF TRAY
End Sub
Private Sub picClock_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  ReleaseCapture
  SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub Timer1_Timer()
  picIcon.Picture = picSrc.Picture
  picIcon.Cls
  CurTime = Time
  Ang = Val(Left(CurTime, InStr(CurTime, ":") - 1)) / 12 * 360
  Ang2 = Val(Mid(CurTime, InStr(CurTime, ":") + 1, 2)) / 60 * 360
  Ang3 = Val(Mid(CurTime, InStr(InStr(CurTime, ":") + 1, CurTime, ":") + 1, 2)) / 60 * 360
  x = Sin((Ang + Ang2 / 12) / 180 * 3.14)
  y = -Cos((Ang + Ang2 / 12) / 180 * 3.14)
  X2 = Sin(Ang2 / 180 * 3.14)
  Y2 = -Cos(Ang2 / 180 * 3.14)
  X3 = Sin(Ang3 / 180 * 3.14)
  Y3 = -Cos(Ang3 / 180 * 3.14)
  picIcon.Line (7, 7)-(7 + X3 * 8, 7 + Y3 * 8), 0
  picIcon.Line (7, 7)-(7 + X2 * 7, 7 + Y2 * 7), 128
  picIcon.Line (7, 8)-(7 + X2 * 7, 8 + Y2 * 7), 128
  picIcon.Line (8, 8)-(8 + X2 * 7, 8 + Y2 * 7), 128
  picIcon.Line (7, 7)-(7 + x * 6, 7 + y * 6), 0
  picIcon.Line (8, 7)-(8 + x * 6, 7 + y * 6), 0
  picIcon.Line (7, 8)-(7 + x * 6, 8 + y * 6), 0
  picIcon.Line (8, 8)-(8 + x * 6, 8 + y * 6), 0
  picClock.Cls
  picClock.Line (Radius, Radius)-(Radius + x * Radius / 3 * 2, Radius + y * Radius / 3 * 2), 0
  picClock.Line (Radius, Radius)-(Radius + X2 * Radius / 4 * 3, Radius + Y2 * Radius / 4 * 3), 128
  picClock.Line (Radius, Radius)-(Radius + X3 * (Radius - 2), Radius + Y3 * (Radius - 2)), 0
  picClock.Circle (Radius, Radius), 1, RGB(255, 255, 255)
  picClock.PSet (Radius, Radius), RGB(255, 255, 255)
  picIcon.Refresh
  picClock.Refresh
  If Dir(App.Path + "\Tray.ico") > "" Then Kill App.Path + "\Tray.ico"
  Open App.Path + "\TrayBack.ico" For Binary Access Read As #1
  a$ = Space(LOF(1))
  Get #1, 1, a$
  Close #1
  Open App.Path + "\Tray.ico" For Binary As #1
  Put #1, 1, a$
  n = 126
  a$ = " "
  For j = 15 To 0 Step -1
    For i = 0 To 15 Step 2
      n = n + 1
      Get #1, Int(n), a$
      If picIcon.Point(i, j) <= 128 Then
        a$ = Chr(Asc(a$) Mod 16 + picIcon.Point(i, j) / 128 * 1 * 16)
      End If
      If picIcon.Point(i + 1, j) <= 128 Then
        a$ = Chr(Int(Asc(a$) / 16) * 16 + picIcon.Point(i + 1, j) / 128 * 1)
      End If
      Put #1, Int(n), a$
    Next
  Next
  Close #1
  ToolTipText = Time & " " & Date
  picClock.ToolTipText = ToolTipText
  picIcon.Picture = LoadPicture(App.Path + "\tray.ico")
  TrayIcon.hIcon = picIcon.Picture
  TrayIcon.szTip = ToolTipText & Chr$(0)
  Call Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
End Sub
