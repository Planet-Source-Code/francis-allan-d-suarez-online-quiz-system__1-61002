VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "First Visit Change Password"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtverifynewPassword 
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
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtnewpassword 
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
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Line Line23 
      X1              =   4560
      X2              =   2520
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line22 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4560
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line21 
      X1              =   120
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line20 
      X1              =   120
      X2              =   2040
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line19 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line18 
      X1              =   4320
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line17 
      BorderColor     =   &H80000005&
      X1              =   4320
      X2              =   4560
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line16 
      X1              =   4320
      X2              =   4560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line15 
      X1              =   120
      X2              =   2040
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line13 
      X1              =   4320
      X2              =   4560
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000005&
      X1              =   4320
      X2              =   4560
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line10 
      X1              =   4320
      X2              =   4560
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   2040
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   2040
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   3360
      X2              =   4560
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line6 
      X1              =   3360
      X2              =   4560
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   1200
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   1200
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   3360
      X2              =   4560
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   1200
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   3120
      Y2              =   3120
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
      Left            =   120
      TabIndex        =   5
      Top             =   2040
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
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "This is your first visit, please change the password initially set on this user account. Thank you."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Trim(txtNewPassword.Text) <> Trim(txtVerifyNewPassword.Text) Then
    MsgBox ("The password could not be verified. Please type it again.")
    txtNewPassword.Text = ""
    txtVerifyNewPassword.Text = ""
    txtNewPassword.SetFocus
Else
    facultyPassword = UCase(txtNewPassword.Text)
    Unload Me
    MsgBox ("Password has been changed!")
End If
End Sub
 
