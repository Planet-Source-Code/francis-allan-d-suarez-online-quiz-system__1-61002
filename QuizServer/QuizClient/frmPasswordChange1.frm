VERSION 5.00
Begin VB.Form frmPasswordChange1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Old Password"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "&Change Password"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtOldPassword 
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
      Top             =   360
      Width           =   1935
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   4560
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line5 
      X1              =   4560
      X2              =   120
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   4560
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
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Type New Password:"
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
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Type Old Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmPasswordChange1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChangePassword_Click()
IdleTime = 28
If UCase(Trim(GlobalPassword)) = UCase(Trim(txtOldPassword.Text)) Then
  If Trim(txtNewPassword.Text) <> Trim(txtVerifyNewPassword.Text) Then
      MsgBox ("The password could not be verified. Please type it again.")
      txtNewPassword.Text = ""
      txtVerifyNewPassword.Text = ""
      txtNewPassword.SetFocus
  Else
      If UCase(Trim(txtNewPassword.Text)) <> "PASS" Then
         GlobalStudentPassword = UCase(txtNewPassword.Text)
         FrmQuizClient.Winsock1.SendData "CLogin" & GlobalUserName & "~" & _
                 UCase(Trim(txtOldPassword.Text)) & "~" & UCase(Trim(txtNewPassword.Text))
         Unload Me
         MsgBox ("Password has been changed!")
       Else
         MsgBox ("You have just entered our universal password. Please try another one.")
         txtNewPassword.Text = ""
         txtVerifyNewPassword.Text = ""
         txtNewPassword.SetFocus
       End If
   End If
End If
End Sub

Private Sub txtNewPassword_Change()
If txtOldPassword.Text <> "" And txtNewPassword.Text <> "" And txtVerifyNewPassword.Text <> "" Then
  cmdChangePassword.Enabled = True
End If
IdleTime = 26
End Sub

Private Sub txtOldPassword_Change()
If txtOldPassword.Text <> "" And txtNewPassword.Text <> "" And txtVerifyNewPassword.Text <> "" Then
  cmdChangePassword.Enabled = True
End If
IdleTime = 25
End Sub

Private Sub txtVerifyNewPassword_Change()
If txtOldPassword.Text <> "" And txtNewPassword.Text <> "" And txtVerifyNewPassword.Text <> "" Then
  cmdChangePassword.Enabled = True
End If
IdleTime = 27
End Sub



