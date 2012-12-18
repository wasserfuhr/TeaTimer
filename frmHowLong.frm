VERSION 5.00
Begin VB.Form frmHowLong 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TeaTimer - Start"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   ControlBox      =   0   'False
   Icon            =   "frmHowLong.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Timer timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "=LoadResString(103)"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdMakeTea 
      Caption         =   "LoadResString(102)"
      Default         =   -1  'True
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtSeconds 
      Alignment       =   1  'Rechts
      Height          =   285
      HideSelection   =   0   'False
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblHowLong 
      Caption         =   "LoadResString(101)"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmHowLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim timeleft As Long
Dim timerIsSet As Boolean

Private Sub cmdCancel_Click()
Hide
End Sub

Private Sub cmdMakeTea_Click()
makeTea (txtSeconds.Text)
Hide
End Sub

Sub makeTea(ByVal l As Long)
    timeleft = l
    timerIsSet = True
    With tbIcon
        .hIcon = brewingpic
    End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, tbIcon)
End Sub

Private Sub Form_GotFocus()
txtSeconds.SetFocus
End Sub

Private Sub Form_Load()
timerIsSet = False
txtSeconds.Text = frmOptionFrame.txtDefaultTime.Text
lblHowLong.Caption = LoadResString(101)
cmdCancel.Caption = LoadResString(103)
cmdMakeTea.Caption = LoadResString(102)


End Sub


Private Sub timer1_Timer()

If timerIsSet Then

If timeleft > 0 Then
    timeleft = timeleft - 1
    With tbIcon
      .szTip = LoadResString(105) & Str$(timeleft) & LoadResString(104) & vbNullChar
    End With
    Call Shell_NotifyIcon(NIM_MODIFY, tbIcon)
    
Else
    timerIsSet = False

    With tbIcon
        .hIcon = hotpic
        .szTip = LoadResString(106) & vbNullChar
    End With
    Call Shell_NotifyIcon(NIM_MODIFY, tbIcon)
    Dim s As String
    s = MsgBox(LoadResString(106), vbInformation, "TeaTimer - It's teatime!")
     
End If
End If
    
End Sub
