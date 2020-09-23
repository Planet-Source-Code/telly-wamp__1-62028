Attribute VB_Name = "Tray_Mod"
Option Explicit

'Systray
  Public Const WM_LBUTTONDOWN = &H201      'Button down
  Public Const WM_LBUTTONUP = &H202        'Button up
  Public Const WM_LBUTTONDBLCLK = &H203    'Double-click
  Public Const WM_RBUTTONDOWN = &H204      'Button down
  Public Const WM_RBUTTONUP = &H205        'Button up
  Public Const WM_RBUTTONDBLCLK = &H206    'Double-click
  Public Const WM_MOUSEMOVE = &H200
  Public Const NIM_ADD = &H0
  Public Const NIM_MODIFY = &H1
  Public Const NIM_DELETE = &H2
  Public Const NIF_MESSAGE = &H1
  Public Const NIF_ICON = &H2
  Public Const NIF_TIP = &H4

Public Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private SysTray As NOTIFYICONDATA

'Stay On top
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Sub ToSysTray(ByVal frm As Form)
frm.Show
frm.Refresh
SysTray.cbSize = Len(SysTray)
SysTray.hwnd = frm.hwnd
SysTray.uID = vbNull
SysTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
SysTray.uCallbackMessage = WM_MOUSEMOVE
SysTray.hIcon = frm.Icon
SysTray.szTip = frm.Caption & vbNullChar
Call Shell_NotifyIcon(NIM_ADD, SysTray)
App.TaskVisible = False
frm.Hide
End Sub

Sub KillTray(ByVal frm As Form)
'Kill Icon
  SysTray.cbSize = Len(SysTray)
  SysTray.hwnd = frm.hwnd
  SysTray.uID = vbNull
Shell_NotifyIcon NIM_DELETE, SysTray
End Sub

Sub StayOnTop(ByVal Ontop As Boolean, frm As Form)
If Ontop Then
    SetWindowPos frm.hwnd, -1, 0, 0, 0, 0, TOPMOST_FLAGS
Else
    SetWindowPos frm.hwnd, -2, 0, 0, 0, 0, TOPMOST_FLAGS
End If
End Sub
