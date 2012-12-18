Attribute VB_Name = "systrayModul"
Option Explicit
Public brewingpic As Picture
Public hotpic As Picture
Public coldpic As Picture

' NotifyIconMessage-Konstanten
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
'--------------------------------------
Public Const WM_MOUSEMOVE = &H200
'--------------------------------------
Public Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type
Public tbIcon As NOTIFYICONDATA
'--------------------------------------
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, lpdata As NOTIFYICONDATA) As Long

Public Declare Function SetForegroundWindow& Lib "user32" (ByVal hwnd&)

Sub Main()

Set coldpic = LoadResPicture(101, 1)
Set brewingpic = LoadResPicture(102, 1)
Set hotpic = LoadResPicture(103, 1)
frmSystray.img1.Picture = coldpic

With tbIcon
  .cbSize = Len(tbIcon)
  .hwnd = frmSystray.txtTooltip.hwnd   '# txtTooltip ist die "Empfänger-Textbox"
  .uID = 2&
  .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
  .uCallbackMessage = WM_MOUSEMOVE
  .hIcon = frmSystray.img1.Picture
  .szTip = LoadResString(107) & vbNullChar
  
End With

Call Shell_NotifyIcon(NIM_ADD, tbIcon)  '# Symbol in Taskbar einrichten

Load frmHowLong
Load frmOptionFrame
End Sub

