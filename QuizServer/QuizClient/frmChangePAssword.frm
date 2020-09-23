VERSION 5.00
Begin VB.Form frmChangePAssword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "&Change Password"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtVerifyNewPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtNewPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   4320
      X2              =   120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   2280
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2280
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label3 
      Caption         =   "Verify New Password:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "This is your first login session. You are required to change your account password. Please change it now. Thank you."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmChangePAssword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChangePassword_Click()
IdleTime = 25
  If Trim(txtNewPassword.Text) <> Trim(txtVerifyNewPassword.Text) Then
      MsgBox ("The password could not be verified. Please type it again.")
      txtNewPassword.Text = ""
      txtVerifyNewPassword.Text = ""
      txtNewPassword.SetFocus
  Else
      If UCase(Trim(txtNewPassword.Text)) <> "PASS" Then
         GlobalStudentPassword = UCase(txtNewPassword.Text)
         Unload Me
         MsgBox ("Password has been changed!")
       Else
         MsgBox ("You have just entered our universal password. Please try another one.")
         txtNewPassword.Text = ""
         txtVerifyNewPassword.Text = ""
         txtNewPassword.SetFocus
       End If
   End If
End Sub

Private Sub txtNewPassword_Change()
IdleTime = 23
End Sub

Private Sub txtVerifyNewPassword_Change()
IdleTime = 24
End Sub


