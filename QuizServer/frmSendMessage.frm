VERSION 5.00
Begin VB.Form frmSendMessage 
   Caption         =   "Send Message To The Clients"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   5520
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000080&
      ForeColor       =   &H80000005&
      Height          =   4545
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000080&
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5160
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
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   2415
   End
End
Attribute VB_Name = "frmSendMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
Dim DClientIndex As Integer
On Error Resume Next
For DClientIndex = 1 To 25
   frmQuizServerRun.Socket(DClientIndex).SendData "SendMessage~" & Text1.Text
Next DClientIndex
List1.AddItem Text1.Text
Text1.Text = ""
Text1.SetFocus
cmdSend.Enabled = False
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
   cmdSend.Enabled = True
  Else
   cmdSend.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
If globalsendList1 <> "" Then
List1.AddItem globalsendList1
End If
Timer1.Enabled = False
End Sub
