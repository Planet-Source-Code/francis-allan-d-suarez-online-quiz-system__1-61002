VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About "
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1455
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmAbout.frx":0000
      Top             =   120
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   1440
      Picture         =   "frmAbout.frx":01D6
      ScaleHeight     =   2115
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Line Line5 
      X1              =   3120
      X2              =   4560
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   3120
      X2              =   4560
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   1680
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   1680
      X2              =   120
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label6 
      Caption         =   "Ziggurat01@yahoo.com"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Email Address:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Francis Allan D. Suarez"
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
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Click me!"
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Analyst/Programmer:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "&Ok" Then
   Unload Me
Else
   Command1.Caption = "&Ok"
   Label2.Visible = True
   Label3.Visible = True
   Label5.Visible = True
   Label6.Visible = True
   Text1.Visible = True
   Picture1.Visible = False
End If
IdleTime = 51
End Sub

Private Sub Form_Load()
Picture1.Visible = False
IdleTime = 52
End Sub

Private Sub Label3_Click()
   Label2.Visible = False
   Label3.Visible = False
   Label5.Visible = False
   Label6.Visible = False
Text1.Visible = False
Command1.Caption = "&Back"
Picture1.Visible = True
IdleTime = 22
End Sub

Private Sub Label4_Click()

End Sub
