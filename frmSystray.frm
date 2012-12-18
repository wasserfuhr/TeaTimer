VERSION 5.00
Begin VB.Form frmSystray 
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7410
   Icon            =   "frmSystray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.TextBox txtTooltip 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   3015
   End
   Begin VB.Image img1 
      Height          =   480
      Left            =   480
      Picture         =   "frmSystray.frx":0152
      Top             =   240
      Width           =   480
   End
   Begin VB.Menu taskmenu 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnShowWindow 
         Caption         =   "Show Window"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    mnExit.Caption = LoadResString(108)
    mnShowWindow.Caption = LoadResString(109)
End Sub

Private Sub mnExit_Click()
  Call Shell_NotifyIcon(NIM_DELETE, tbIcon)
  End
End Sub



Private Sub mnShowWindow_Click()
        frmHowLong.Show
    Call SetForegroundWindow(frmHowLong.hwnd)
    frmHowLong.SetFocus

End Sub

Private Sub txtTooltip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'# Die Textbox "txtTooltip" empfängt die Notifikationsmeldungen vom Symbol
Dim dl

dl = Hex(X)
'Debug.Print dl

'If (dl = "1E2D") Or (dl = "1824") Then            '# Doppelklick links
'    Call SetForegroundWindow(Me.hwnd)
'    DoEvents
 '   frmHowLong.makeTea (frmOptionFrame.txtDefaultTime.Text)
          
If dl = "1E0F" Then                    '# Linke Maustaste wurde gedrückt
    'Call SetForegroundWindow(Me.hwnd)
    'DoEvents
    
    If tbIcon.hIcon = hotpic Then
        With tbIcon
         .hIcon = coldpic
         .szTip = LoadResString(107) & vbNullChar
        End With
        
        Call Shell_NotifyIcon(NIM_MODIFY, tbIcon)
    
    Else
        frmHowLong.Show
        frmHowLong.txtSeconds.SelLength = 3
        Call SetForegroundWindow(frmHowLong.hwnd)
        frmHowLong.SetFocus
        'DoEvents
    End If
    
  
 ElseIf (dl = "1E3C") Or (dl = "1830") Then            '# Rechte Maustaste wurde gedrückt
    'Call SetForegroundWindow(Me.hwnd)
    DoEvents
    PopupMenu taskmenu, 2
  
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Shell_NotifyIcon(NIM_DELETE, tbIcon)
End Sub

