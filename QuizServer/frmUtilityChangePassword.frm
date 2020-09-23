VERSION 5.00
Begin VB.Form frmUtilityChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change of Password"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
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
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
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
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2175
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
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4560
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line8 
      X1              =   4560
      X2              =   120
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line7 
      X1              =   3720
      X2              =   4560
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   3720
      X2              =   4560
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      X1              =   3720
      X2              =   4560
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   4560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   960
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   960
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   960
      Y1              =   2400
      Y2              =   2400
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
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
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
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Old Password:"
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
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmUtilityChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
strCn = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1 = New ADODB.Connection
strCn1.Open strCn
Call Initialization
With rsFaculty
  .MoveFirst
  .Find "username= '" & UCase(Trim(Form1.txtuserName.Text)) & "'"
  If Not .EOF Then
      If !facultyPassword = UCase(txtOldPassword.Text) Then
          If Trim(txtNewPassword.Text) <> Trim(txtVerifyNewPassword.Text) Then
               MsgBox ("The password could not be verified. Please type it again.")
               txtNewPassword.Text = ""
               txtVerifyNewPassword.Text = ""
               txtNewPassword.SetFocus
            Else
               !facultyPassword = UCase(Trim(txtNewPassword.Text))
               Unload Me
               MsgBox ("Password has been changed!")
           End If
       End If
       .Update
      Else
       MsgBox "Please check your Old Password", vbExclamation
      End If
End With
End Sub

Private Sub Command2_Click()
End
End Sub
