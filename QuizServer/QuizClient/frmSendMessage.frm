VERSION 5.00
Begin VB.Form frmSendMessage 
   Caption         =   "Send Message"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   3000
      Top             =   4800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   4800
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2280
      Top             =   4800
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000080&
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5280
      Width           =   4695
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000080&
      ForeColor       =   &H80000005&
      Height          =   4350
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Type Your Message Here:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   2415
   End
End
Attribute VB_Name = "frmSendMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSend_Click()
If FrmQuizClient.Winsock1.State = sckConnected Then
      'List1.AddItem Text1.Text
      FrmQuizClient.Winsock1.SendData "GroupMessage~" & " " & GlobalUserName & "-" & Text1.Text
      Text1.Text = ""
      Text1.SetFocus
      cmdSend.Enabled = False
   Else
       MsgBox ("You are not connected to the server!")
End If

End Sub

Private Sub Form_Load()
GlobalMessageFlag = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
GlobalMessageFlag = 0
Timer2.Enabled = False
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
  cmdSend.Enabled = True
 Else
  cmdSend.Enabled = False
End If
IdleTime = 55
End Sub

Private Sub Timer1_Timer()
If globalSendList1 <> "" Then
List1.AddItem globalSendList1
End If
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Unload Me
End Sub

