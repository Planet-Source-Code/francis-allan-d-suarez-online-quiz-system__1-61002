VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmQuizClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Client 2.12.10"
   ClientHeight    =   5550
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   4920
   End
   Begin VB.TextBox txtOptionAnswer 
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   0
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton cmdQuizFormRequest 
      Caption         =   "&Quiz Form Request"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3615
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Top             =   4800
      Width           =   1095
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Option5"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Start"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox txtIpAddress 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "192.168.10.1"
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "E."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "D."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "C."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "B."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "A."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   3120
      Width           =   255
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000009&
      X1              =   7320
      X2              =   5040
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line11 
      X1              =   7320
      X2              =   5040
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      X1              =   4800
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line9 
      X1              =   3840
      X2              =   4800
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   3840
      X2              =   4800
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line7 
      X1              =   3840
      X2              =   4800
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line6 
      X1              =   6960
      X2              =   7320
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      X1              =   6960
      X2              =   7320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   7320
      X2              =   3840
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line3 
      X1              =   7320
      X2              =   3840
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Online Quiz "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3720
      TabIndex        =   8
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Online Quiz "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   3840
      TabIndex        =   24
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label lblInstruction 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   5040
      Width           =   4935
   End
   Begin VB.Label lblTimeCounter 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   41.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   5040
      TabIndex        =   22
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   1215
      Left            =   3480
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Shape ShpStop 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Connection Status:"
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
      Left            =   5040
      TabIndex        =   18
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Shape ShpGo 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Quiz Form:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblQuestion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   10
      Top             =   2040
      Width           =   6735
   End
   Begin VB.Label lblItemNumber 
      Caption         =   "1."
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
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   7320
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7320
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Server IP Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      ToolTipText     =   "Click me to change the Server IP number!"
      Top             =   0
      Width           =   2415
   End
   Begin VB.Menu mnuChangePassword 
      Caption         =   "Change &Password"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuSpokenEnglishCredits 
      Caption         =   "Sp&oken English Credits"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuSendMessage 
      Caption         =   "Send &Message"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "A&bout"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "FrmQuizClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GlobalData As String
Public LengthofArray As Integer
Dim QuestionsArray() As String
Dim TimeFlag As Integer
Dim varTimer As Integer
Dim TimerCtr As Integer
Dim Logtime As Integer
Dim Logtime1 As Integer

Private Sub cmdClose_Click()

Winsock1.Close
End
End Sub
Sub Disconnecting_Idle()
GlobalUserName = ""
GlobalPassword = ""

Flag = 0
MyValue = 0
VarQuizCompositeKey = ""
VarStudentUsername = ""
VarQuizScore = 0
VarQuestionSequence = ""
StudentQuizArrangement = 0
GlobalStudentPassword = ""
GlobalItemNumber = 0
Combo1.Clear
Combo1.Text = ""
Option1.Caption = ""
Option2.Caption = ""
Option3.Caption = ""
Option4.Caption = ""
Option5.Caption = ""
lblQuestion.Caption = ""
lblItemNumber.Caption = "1"
'cmdNext.Enabled = True
   Combo1.Clear
   Winsock1.Close
   ShpGo.Visible = False
   ShpStop.Visible = True
   mnuChangePassword.Enabled = False
   mnuSpokenEnglishCredits.Enabled = False
   mnuSendMessage.Enabled = False
   cmdLogin.Enabled = False
   cmdQuizFormRequest.Enabled = False
   cmdSend.Enabled = False
   Combo1.Enabled = False
   cmdNext.Enabled = False
   cmdConnect.Caption = "&Connect"
   lblInstruction.Caption = "Disconnected - No activity for the last two (2) minutes."
   TimeFlag = 0



End Sub

Sub ExecuteAnswer()
Dim SequenceFlag As Integer
Dim Flagger As String
Dim Adder As Integer
Dim VerifyValue As Integer
SequenceFlag = 0
IdleTime = 1
Dim LengthVarQuestionSequence As Integer

Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False

If Trim(txtOptionAnswer.Text) Like GlobalArray(Val(lblItemNumber.Caption), 7) Then
  VarQuizScore = 1
 Else
  VarQuizScore = 0
End If
If Val(lblItemNumber.Caption) <= (LengthofArray + 1) Then
    lblItemNumber.Caption = Val(lblItemNumber.Caption) + 1
End If
If lblQuestion.Caption <> "" And Option1.Caption <> "" And Option2.Caption <> "" _
   And Option3.Caption <> "" And Option4.Caption <> "" And Option5.Caption <> "" Then
   
Winsock1.SendData "QuizAnswer" & VarQuizCompositeKey _
                  & "~" & VarStudentUsername & "~" & _
                  VarQuizScore & "~" & VarQuestionSequence & "~" & StudentQuizArrangement
End If

If Val(lblItemNumber.Caption) <= (LengthofArray) Then
            If Len(GlobalArray(Val(lblItemNumber.Caption), 8)) = 1 Then
               VarQuestionSequence = VarQuestionSequence & "0" & GlobalArray(Val(lblItemNumber.Caption), 8)
            Else
               VarQuestionSequence = VarQuestionSequence & GlobalArray(Val(lblItemNumber.Caption), 8)
            End If
            lblQuestion.Caption = GlobalArray(Val(lblItemNumber.Caption), 1)
            Option1.Caption = GlobalArray(Val(lblItemNumber.Caption), 2)
            Option2.Caption = GlobalArray(Val(lblItemNumber.Caption), 3)
            Option3.Caption = GlobalArray(Val(lblItemNumber.Caption), 4)
            Option4.Caption = GlobalArray(Val(lblItemNumber.Caption), 5)
            Option5.Caption = GlobalArray(Val(lblItemNumber.Caption), 6)
End If

If Val(lblItemNumber.Caption) >= (LengthofArray + 1) Then
  lblItemNumber.Caption = lblItemNumber.Caption - 1
  cmdNext.Enabled = False
   GlobalGoTimeCounter = 0
End If
lblTimeCounter.Caption = varTimer
End Sub

Private Sub cmdConnect_Click()
GlobalUserName = ""
GlobalPassword = ""

Flag = 0
MyValue = 0
VarQuizCompositeKey = ""
VarStudentUsername = ""
VarQuizScore = 0
VarQuestionSequence = ""
StudentQuizArrangement = 0
GlobalStudentPassword = ""
GlobalItemNumber = 0
Combo1.Clear
Combo1.Text = ""
Option1.Caption = ""
Option2.Caption = ""
Option3.Caption = ""
Option4.Caption = ""
Option5.Caption = ""
lblQuestion.Caption = ""
lblItemNumber.Caption = "1"
'cmdNext.Enabled = True
On Error GoTo ConnectError
If cmdConnect.Caption = "&Connect" Then
   If Winsock1.State <> sckConnected Then
       Winsock1.RemoteHost = Trim(txtIpAddress.Text)
       Winsock1.RemotePort = 1007
       Winsock1.Connect
       ShpGo.Visible = True
       ShpStop.Visible = False
       cmdConnect.Caption = "&Disconnect"
       cmdLogin.Enabled = True
       lblInstruction.Caption = "Connection is already established. You may click the LOGIN button."

       Exit Sub
    Else
      MsgBox "Already Connected at " & Winsock1.RemoteHost
   
      Exit Sub
   End If
  
 ElseIf cmdConnect.Caption = "&Disconnect" Then
   Combo1.Clear
   Winsock1.Close
   ShpGo.Visible = False
   ShpStop.Visible = True
   mnuChangePassword.Enabled = False
   mnuSpokenEnglishCredits.Enabled = False
   mnuSendMessage.Enabled = False
   cmdLogin.Enabled = False
   cmdQuizFormRequest.Enabled = False
   cmdSend.Enabled = False
   Combo1.Enabled = False
   cmdNext.Enabled = False
   cmdConnect.Caption = "&Connect"
   Label7.Visible = False
   Label8.Visible = False
   Label9.Visible = False
   Label10.Visible = False
   Label11.Visible = False
   Option1.Visible = False
   Option2.Visible = False
   Option3.Visible = False
   Option4.Visible = False
   Option5.Visible = False
   lblItemNumber.Visible = False
      
         
   lblInstruction.Caption = "You have been disconnected from the server. You may try to connect again."
   TimeFlag = 0
   
End If
ConnectError:
  MsgBox "Server not running!"
  
End Sub

Private Sub cmdLogin_Click()
If Winsock1.State = sckConnected Then
       frmLogin.Show (1)
       Winsock1.SendData "Login" & GlobalUserName & "~" & GlobalPassword
   Else
       MsgBox ("You are not connected to the server!")
       cmdConnect.Caption = "&Connect"
       ShpGo.Visible = False
       ShpStop.Visible = True
       cmdLogin.Enabled = False
End If
IdleTime = 54
End Sub

Private Sub cmdNext_Click()
Dim SequenceFlag As Integer
Dim Flagger As String
Dim Adder As Integer
Dim VerifyValue As Integer
SequenceFlag = 0
IdleTime = 2
Dim LengthVarQuestionSequence As Integer

Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False

If Trim(txtOptionAnswer.Text) Like GlobalArray(Val(lblItemNumber.Caption), 7) Then
  VarQuizScore = 1
 Else
  VarQuizScore = 0
End If
If Val(lblItemNumber.Caption) <= (LengthofArray + 1) Then
    lblItemNumber.Caption = Val(lblItemNumber.Caption) + 1
End If
If lblQuestion.Caption <> "" And Option1.Caption <> "" And Option2.Caption <> "" _
   And Option3.Caption <> "" And Option4.Caption <> "" And Option5.Caption <> "" Then
   
Winsock1.SendData "QuizAnswer" & VarQuizCompositeKey _
                  & "~" & VarStudentUsername & "~" & _
                  VarQuizScore & "~" & VarQuestionSequence & "~" & StudentQuizArrangement
End If

If Val(lblItemNumber.Caption) <= (LengthofArray) Then
            If Len(GlobalArray(Val(lblItemNumber.Caption), 8)) = 1 Then
               VarQuestionSequence = VarQuestionSequence & "0" & GlobalArray(Val(lblItemNumber.Caption), 8)
            Else
               VarQuestionSequence = VarQuestionSequence & GlobalArray(Val(lblItemNumber.Caption), 8)
            End If
            lblQuestion.Caption = GlobalArray(Val(lblItemNumber.Caption), 1)
            Option1.Caption = GlobalArray(Val(lblItemNumber.Caption), 2)
            Option2.Caption = GlobalArray(Val(lblItemNumber.Caption), 3)
            Option3.Caption = GlobalArray(Val(lblItemNumber.Caption), 4)
            Option4.Caption = GlobalArray(Val(lblItemNumber.Caption), 5)
            Option5.Caption = GlobalArray(Val(lblItemNumber.Caption), 6)
End If

If Val(lblItemNumber.Caption) >= (LengthofArray + 1) Then
  cmdNext.Enabled = False
  lblItemNumber.Caption = lblItemNumber.Caption - 1
   GlobalGoTimeCounter = 0
End If
lblTimeCounter.Caption = varTimer
cmdNext.Enabled = False
End Sub

Private Sub cmdQuizFormRequest_Click()
 If Winsock1.State = sckConnected Then
       Combo1.Enabled = True
       Winsock1.SendData "Quizforms"
       cmdSend.Enabled = True
       lblInstruction.Caption = "You are requesting for the quizforms."
 End If
 Combo1.Clear
 IdleTime = 3
End Sub

Private Sub cmdSend_Click()
If Combo1.Text = "" Then
 MsgBox "Please select a quiz form.", vbOKOnly
 Combo1.SetFocus
 cmdSend.Enabled = True
Else
If Winsock1.State = sckConnected Then
       Winsock1.SendData "GoQuiz" & Combo1.Text
'       cmdNext.Enabled = True
       cmdLogin.Enabled = False
       cmdSend.Enabled = False
       lblInstruction.Caption = "You are about to start with the quiz."
       cmdQuizFormRequest.Enabled = False

       
End If
End If

IdleTime = 4

End Sub

Private Sub Combo1_Change()
IdleTime = 28
End Sub

Private Sub Form_Click()
IdleTime = 35
End Sub

Private Sub Form_Load()
lblInstruction.Caption = "Please click CONNECT button to start. Please click the SERVER IP ADDRESS label to change IP number."
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IdleTime = 30
End Sub

Private Sub Label1_Click()
txtIpAddress.Locked = False
lblInstruction.Caption = "You are about to change your server IP."
txtIpAddress.Text = InputBox("Please enter the new IP Address of the server!")
If txtIpAddress.Text = "" Then
   txtIpAddress.Text = "192.168.10.1"
End If
txtIpAddress.Locked = True
IdleTime = 6
End Sub

Private Sub Label2_Click()
IdleTime = 29
End Sub

Private Sub Label3_Click()
IdleTime = 33
End Sub

Private Sub lblInstruction_Click()
IdleTime = 31
End Sub

Private Sub lblItemNumber_Click()
IdleTime = 32
End Sub

Private Sub lblQuestion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IdleTime = 35
End Sub

Private Sub lblTimeCounter_Click()
IdleTime = 20
End Sub

Private Sub lblTimeCounter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IdleTime = 39
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show (1)
IdleTime = 5
End Sub

Private Sub mnuChangePassword_Click()
lblInstruction.Caption = "You are planning to change you password."
frmPasswordChange1.Show (1)
IdleTime = 7
End Sub

Private Sub mnuExit_Click()
Winsock1.Close
End
End Sub

Private Sub mnuSendMessage_Click()
If glabalmessageflag = 0 Then
  GlobalMessageFlag = 1
  frmSendMessage.Show 1
End If

End Sub

Private Sub mnuSpokenEnglishCredits_Click()
  lblInstruction.Caption = "You are requesting for your Spoken English Credits status."
  Winsock1.SendData "EnglishCredits" & GlobalUserName
  IdleTime = 8
End Sub

Private Sub Option1_Click()
txtOptionAnswer.Text = Option1.Caption
IdleTime = 9
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IdleTime = 36
End Sub

Private Sub Option2_Click()
txtOptionAnswer.Text = Option2.Caption
IdleTime = 10
End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IdleTime = 37
End Sub

Private Sub Option3_Click()
txtOptionAnswer.Text = Option3.Caption
IdleTime = 11
End Sub

Private Sub Option4_Click()
txtOptionAnswer.Text = Option4.Caption
IdleTime = 12
End Sub

Private Sub Option5_Click()
txtOptionAnswer.Text = Option5.Caption
IdleTime = 13
End Sub

Private Sub Option5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IdleTime = 38
End Sub

Private Sub Timer1_Timer()
If GlobalGoTimeCounter = 1 Then
  If lblTimeCounter.Caption >= 0 And TimeFlag = 1 Then
     lblTimeCounter.Caption = lblTimeCounter.Caption - 1
  End If
  If TimeFlag = 0 Then
    lblTimeCounter.Caption = varTimer
  End If
End If
If lblTimeCounter = 0 Then
   Call ExecuteAnswer
End If
If lblTimeCounter.Caption = (varTimer - 2) Then
  cmdNext.Enabled = True
End If

If TimerCtr = 0 Then
    Logtime = IdleTime
End If
TimerCtr = TimerCtr + 1
If TimerCtr >= 1 And TimerCtr < 120 Then
    If Logtime <> IdleTime Then
       TimerCtr = 0
       Logtime = IdleTime
    End If
End If
If TimerCtr >= 120 Then
  Call Disconnecting_Idle
  TimerCtr = 0
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
Dim SearchChar As String
Dim TargetChar As String
Dim VarLength As Integer
Dim Item As String
Dim IPNumber As Integer
Dim Addition As Integer
Dim DivideArray As Integer
Dim varUnfinishedSequence As Integer
Dim varRoute As Integer
Dim QuestionString As String
Dim CreditsLEft As Integer
SearchChar = "~"
Winsock1.GetData sData, vbString
Label5.Caption = sData
VarLength = Len(sData)
GlobalData = sData

If Left(sData, 16) = "DisconnectClient" Then
   frmSendMessage.Timer2.Enabled = True
   MsgBox "You are going to be disconnected by your instructor!", vbOKOnly
   Combo1.Clear
   Winsock1.Close
   ShpGo.Visible = False
   ShpStop.Visible = True
   mnuChangePassword.Enabled = False
   mnuSpokenEnglishCredits.Enabled = False
   mnuSendMessage.Enabled = False
   cmdLogin.Enabled = False
   cmdQuizFormRequest.Enabled = False
   cmdSend.Enabled = False
   Combo1.Enabled = False
   cmdNext.Enabled = False
   cmdConnect.Caption = "&Connect"
   lblInstruction.Caption = "You have been disconnected from the server. You may try to connect again."
   TimeFlag = 0
End If
If Left(sData, 12) = "SendMessage~" Then
   If GlobalMessageFlag = 0 Then
      globalSendList1 = Right(sData, Len(sData) - 12)
      GlobalMessageFlag = 1
      frmSendMessage.Show 1
   End If
   frmSendMessage.List1.AddItem Right(sData, Len(sData) - 12)
   
End If

If Left(sData, 13) = "GroupMessage~" Then
  If GlobalMessageFlag = 0 Then
      GlobalMessageFlag = 1
      frmSendMessage.Show 1
      
  End If
  frmSendMessage.List1.AddItem Right(sData, Len(sData) - 13)
 
  
End If
' ******************Quizforms
If Left(sData, 9) = "Quizforms" Then
   
   IdleTime = 14
   sData = Right(sData, Len(sData) - 9)
   For i = 2 To VarLength
       If Mid(sData, i, 1) <> SearchChar Then
           Item = Item & Mid(sData, i, 1)
       Else
           Combo1.AddItem Item
           Item = ""
       End If
   Next i
End If
' /******************Quizforms

' ******************Login
If Left(sData, 5) = "Login" Then
   sData = Right(sData, Len(sData) - 5)
   If Left(sData, 2) = "Go" Then
      MsgBox "Welcome " & GlobalUserName & "!"
      cmdLogin.Enabled = False
      
'     cmdQuizFormRequest.SetFocus
'      cmdLogin.Enabled = False
      TimerCtr = 0
      If Right(sData, 1) = 0 Then
        frmChangePAssword.Show (1)
        Winsock1.SendData "ChangePassword" & GlobalUserName & "~" & GlobalStudentPassword
      End If
      cmdQuizFormRequest.Enabled = True
      cmdQuizFormRequest.SetFocus
      mnuSpokenEnglishCredits.Enabled = True
      mnuSendMessage.Enabled = True
      mnuChangePassword.Enabled = True
      lblInstruction.Caption = "You have just logged in. Please click QUIZ FORM REQUEST button."
     Else
      MsgBox "Please check your Username and Password!", vbCritical
      
   End If
End If
' /******************Login

' ******************Finished
If Left(sData, 8) = "Finished" Then
   IdleTime = 15
   sData = Right(sData, Len(sData) - 8)
   MsgBox "Finished! Your score is " & Val(sData), vbOKOnly
    lblQuestion.Visible = False
    cmdNext.Enabled = False
    Option1.Visible = False
    Option2.Visible = False
    Option3.Visible = False
    Option4.Visible = False
    Option5.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    lblItemNumber.Visible = False
    lblInstruction.Caption = "You have just finished answering all the items."
End If
' /******************Finished

' ******************Expired
If sData = "Expired" Then
   IdleTime = 16
   MsgBox "The quizform has expired. You can no longer take the quiz!", vbOKOnly
   cmdNext.Enabled = False
   
   Combo1.Clear
   cmdQuizFormRequest.Enabled = True
   cmdSend.Enabled = False
   cmdQuizFormRequest.SetFocus
   lblInstruction.Caption = "The quiz form has expired already. Try another quiz form."
End If
' /******************Expired

' ******************EnglishCredits
If Left(sData, 14) = "EnglishCredits" Then
   IdleTime = 17
   CreditsLEft = Right(sData, Len(sData) - 14)
   MsgBox "You only have " & CreditsLEft & " credits left!", vbOKOnly
   lblInstruction.Caption = "You have requested for your SEE Program credits."
End If
' /******************EnglishCredits


' ******************CompletedAlready
If Trim(sData) = "CompletedAlready" Then
   IdleTime = 14
   MsgBox "Sorry, but you have already completed taking this quiz!", vbOKOnly
   
   cmdNext.Enabled = False
   Combo1.Clear
   cmdQuizFormRequest.Enabled = True
   cmdSend.Enabled = False
   cmdQuizFormRequest.SetFocus
   lblInstruction.Caption = "You have taken this quiz already. Try another quiz form."
  ' Combo1.SetFocus
End If
' /******************CompletedAlready

' ******************Questions
If Left(sData, 9) = "Questions" Then
   IdleTime = 19
   sData = Right(sData, Len(sData) - 9)
   If sData = "0" Then
     Combo1.Clear
     cmdSend.Enabled = False
     cmdQuizFormRequest.Enabled = True
     MsgBox "Quiz Form does not exist!", vbCritical
     cmdQuizFormRequest.SetFocus
     Exit Sub
   End If
   varTimer = Left(sData, 2)
   sData = Right(sData, Len(sData) - 2)
   mypos = InStr(1, sData, "~", 1)
   LengthofArray = Left(sData, mypos - 1)
   QuestionString = Right(sData, Len(sData) - 2)
   ReDim QuestionsArray(LengthofArray, 8)
   ReDim GlobalArray(LengthofArray, 8)
   txtNumber.Text = LengthofArray
   Flag = 1
   For i = 1 To LengthofArray
     For j = 1 To 7
         mypos = InStr(1, QuestionString, "~", 1)
         If mypos = 1 And j = 1 Then
            QuestionString = Right(QuestionString, Len(QuestionString) - 1)
         End If
         If mypos > 0 Then
            mypos = InStr(1, QuestionString, "~", 1)
            QuestionsArray(i, j) = Left(QuestionString, mypos - 1)
            GlobalArray(i, j) = QuestionsArray(i, j)
            GlobalArray(i, 8) = i
            QuestionsArray(i, 8) = i
         Else
            QuestionsArray(i, j) = QuestionString
            GlobalArray(i, j) = QuestionString
            GlobalArray(i, 8) = i
            QuestionsArray(i, 8) = i
         End If
         
         QuestionString = Right(QuestionString, Len(QuestionString) - mypos)
      Next j
      
   Next i
   
   If Flag = 1 Then
     If MsgBox("The quiz is good for " & txtNumber.Text & " items. Proceed?", vbYesNo) = vbYes Then
        lblInstruction.Caption = "The quiz is on-going..."
        cmdNext.Enabled = True
        lblTimeCounter.Caption = varTimer
        lblQuestion.Visible = True
        Option1.Visible = True
        Option2.Visible = True
        Option3.Visible = True
        Option4.Visible = True
        Option5.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        Label11.Visible = True
        lblItemNumber.Visible = True
        GlobalGoTimeCounter = 1
        TimeFlag = 1
        IPNumber = Right(Winsock1.LocalIP, 1)
        If IPNumber = 0 Then
          IPNumber = 10
        End If
     ' -----------
        If ((IPNumber Mod 2) = 0) Then
            ' /111111111111111111
            If IPNumber <= 5 And ((Int(Second(Time())) Mod 2) = 0) Then
                DivideArray = Int(LengthofArray / 2)
                
                For i = 1 To DivideArray
                    For j = 1 To 8
                       GlobalArray((LengthofArray - DivideArray) + i, j) = QuestionsArray(i, j)
                    Next j
                Next i
                For i = 1 To DivideArray
                    For j = 1 To 8
                       GlobalArray(i, j) = QuestionsArray((LengthofArray - DivideArray) + i, j)
                    Next j
                Next i
                Label5.Caption = "1"
                StudentQuizArrangement = 1
            End If
            ' /111111111111111111111
            ' 2222222222222222222222
            If IPNumber <= 5 And ((Int(Second(Time())) Mod 2) = 1) Then
                For i = 1 To LengthofArray
                    For j = 1 To 8
                       GlobalArray(i, j) = QuestionsArray(i, j)
                    Next j
                Next i
                Label5.Caption = "2"
                StudentQuizArrangement = 2
            End If
            ' /2222222222222222222222
            ' 33333333333333333333333
            If IPNumber > 5 And ((Int(Second(Time())) Mod 2) = 0) Then
              For i = 1 To LengthofArray
                  For j = 1 To 8
                    GlobalArray(LengthofArray - (i - 1), j) = QuestionsArray(i, j)
                 Next j
              Next i
             Label5.Caption = "3"
             StudentQuizArrangement = 3
            End If
           ' /3333333333333333333333333
        
           ' 44444444444444444444444444
            If IPNumber > 5 And ((Int(Second(Time())) Mod 2) = 1) Then
                DivideArray = Int(LengthofArray / 2)
                
                For i = 1 To DivideArray
                     For j = 1 To 8
                       GlobalArray(LengthofArray - (i - 1), j) = QuestionsArray((LengthofArray - DivideArray) + i, j)
                     Next j
                Next i
                For i = 1 To DivideArray
                     For j = 1 To 8
                       GlobalArray((DivideArray) - (i - 1), j) = QuestionsArray(i, j)
                     Next j
                Next i
                Label5.Caption = "4"
                StudentQuizArrangement = 4
             End If
            ' /444444444444444444444444
       End If
' ---------------------- Rearranging the Questions
           
       If ((IPNumber Mod 2) = 1) Then
             ' 55555555555555555555555
             If IPNumber <= 5 And ((Int(Second(Time())) Mod 2) = 0) Then
                 DivideArray = Int(LengthofArray / 2)
                 If LengthofArray Mod 2 = 0 Then
                   Addition = 0
                  Else
                   Addition = 1
                 End If
                 For i = 1 To (DivideArray + Addition)
                    For j = 1 To 8
                       GlobalArray((DivideArray) + i, j) = QuestionsArray(i, j)
                    Next j
                 Next i
                 For i = 1 To DivideArray
                    For j = 1 To 8
                       GlobalArray(i, j) = QuestionsArray(DivideArray + i + Addition, j)
                    Next j
                 Next i
                 Label5.Caption = "5"
                 StudentQuizArrangement = 5
              End If
            ' /555555555555555555555555
            ' 6666666666666666666666666
             If IPNumber <= 5 And ((Int(Second(Time())) Mod 2) = 1) Then
                 For i = 1 To LengthofArray
                   For j = 1 To 8
                    GlobalArray(LengthofArray - (i - 1), j) = QuestionsArray(i, j)
                   Next j
                Next i
              Label5.Caption = "6"
              StudentQuizArrangement = 6
             End If
            ' /666666666666666666666666
            ' 7777777777777777777777777
           If IPNumber > 5 And ((Int(Second(Time())) Mod 2) = 0) Then
              DivideArray = Int(LengthofArray / 2)
              If LengthofArray Mod 2 = 0 Then
                   Addition = 0
                  Else
                   Addition = 1
              End If
                 
              For i = 1 To DivideArray
                 For j = 1 To 8
                    GlobalArray(LengthofArray - (i - 1), j) = QuestionsArray((LengthofArray - DivideArray) + i, j)
                 Next j
              Next i
              For i = 1 To (DivideArray + Addition)
                 For j = 1 To 8
                    GlobalArray((DivideArray + Addition) - (i - 1), j) = QuestionsArray(i, j)
                 Next j
              Next i
              Label5.Caption = "7"
              StudentQuizArrangement = 7
            End If
            ' /7777777777777777777777777
            ' 88888888888888888888888888
           If IPNumber > 5 And ((Int(Second(Time())) Mod 2) = 1) Then
              For i = 1 To LengthofArray
                   For j = 1 To 8
                    GlobalArray(LengthofArray - (i - 1), j) = QuestionsArray(i, j)
                   Next j
                Next i
              Label5.Caption = "8"
              StudentQuizArrangement = 8
           End If
           ' /88888888888888888888888888
        End If
'---------------------------- Rearranging the questions
        If Len(GlobalArray(1, 8)) = 1 Then
           VarQuestionSequence = VarQuestionSequence & "0" & GlobalArray(1, 8)
        Else
           VarQuestionSequence = VarQuestionSequence & GlobalArray(1, 8)
        End If
        lblQuestion.Caption = GlobalArray(1, 1)
        Option1.Caption = GlobalArray(1, 2)
        Option2.Caption = GlobalArray(1, 3)
        Option3.Caption = GlobalArray(1, 4)
        Option4.Caption = GlobalArray(1, 5)
        Option5.Caption = GlobalArray(1, 6)
        'cmdSend.Enabled = True
        VarQuizCompositeKey = Combo1.Text
        VarStudentUsername = GlobalUserName
        'VarQuestionSequence = MyValue
     End If
   End If
End If
' /******************Questions

' ******************Continue
' ******************The student user can continue with the quiz if interrupted
' ******************due to uncontrollable situations like power outage. The
' ******************items taken won't be repeated.
If Left(sData, 8) = "Continue" Then
   IdleTime = 14
   cmdNext.Enabled = True
   varTimer = Mid(sData, 9, 2)
   varUnfinishedSequence = Mid(sData, 11, 2)
   varRoute = Mid(sData, 13, 1)
   sData = Right(sData, Len(sData) - 13)
   mypos = InStr(1, sData, "~", 1)
   LengthofArray = Left(sData, mypos - 1)
   MsgBox "You are about to continue with the quiz! From item/s " & (varUnfinishedSequence + 1) _
          & " to " & LengthofArray, vbOKOnly
    lblTimeCounter.Caption = varTimer
    lblQuestion.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    Label9.Visible = True
    Label10.Visible = True
    Label11.Visible = True
    
    Option1.Visible = True
    Option2.Visible = True
    Option3.Visible = True
    Option4.Visible = True
    Option5.Visible = True
    lblItemNumber.Visible = True
   GlobalGoTimeCounter = 1
   TimeFlag = 1
   QuestionString = Right(sData, Len(sData) - 2)
   ReDim QuestionsArray(LengthofArray, 8)
   ReDim GlobalArray(LengthofArray, 8)
   txtNumber.Text = LengthofArray
   Flag = 1
   For i = 1 To LengthofArray
     For j = 1 To 7
         mypos = InStr(1, QuestionString, "~", 1)
         If mypos = 1 And j = 1 Then
            QuestionString = Right(QuestionString, Len(QuestionString) - 1)
         End If
         If mypos > 0 Then
            mypos = InStr(1, QuestionString, "~", 1)
            QuestionsArray(i, j) = Left(QuestionString, mypos - 1)
            GlobalArray(i, j) = QuestionsArray(i, j)
            GlobalArray(i, 8) = i
            QuestionsArray(i, 8) = i
         Else
            QuestionsArray(i, j) = QuestionString
            GlobalArray(i, j) = QuestionString
            GlobalArray(i, 8) = i
            QuestionsArray(i, 8) = i
         End If
         
         QuestionString = Right(QuestionString, Len(QuestionString) - mypos)
      Next j
      
   Next i
   
   If Flag = 1 Then
     'If MsgBox("The quiz is good for " & txtNumber.Text & " items. Proceed?", vbYesNo) = vbYes Then
   
            ' /111111111111111111
            If varRoute = 1 Then
                DivideArray = Int(LengthofArray / 2)
                
                For i = 1 To DivideArray
                    For j = 1 To 8
                       GlobalArray((LengthofArray - DivideArray) + i, j) = QuestionsArray(i, j)
                    Next j
                Next i
                For i = 1 To DivideArray
                    For j = 1 To 8
                       GlobalArray(i, j) = QuestionsArray((LengthofArray - DivideArray) + i, j)
                    Next j
                Next i
                Label5.Caption = "1"
                StudentQuizArrangement = 1
            End If
            ' /111111111111111111111
            ' 2222222222222222222222
            If varRoute = 2 Then
                For i = 1 To LengthofArray
                    For j = 1 To 8
                       GlobalArray(i, j) = QuestionsArray(i, j)
                    Next j
                Next i
                Label5.Caption = "2"
                StudentQuizArrangement = 2
            End If
            ' /2222222222222222222222
            ' 33333333333333333333333
            If varRoute = 3 Then
              For i = 1 To LengthofArray
                  For j = 1 To 8
                    GlobalArray(LengthofArray - (i - 1), j) = QuestionsArray(i, j)
                 Next j
              Next i
             Label5.Caption = "3"
             StudentQuizArrangement = 3
            End If
           ' /3333333333333333333333333
        
           ' 44444444444444444444444444
            If varRoute = 4 Then
                DivideArray = Int(LengthofArray / 2)
                
                For i = 1 To DivideArray
                     For j = 1 To 8
                       GlobalArray(LengthofArray - (i - 1), j) = QuestionsArray((LengthofArray - DivideArray) + i, j)
                     Next j
                Next i
                For i = 1 To DivideArray
                     For j = 1 To 8
                       GlobalArray((DivideArray) - (i - 1), j) = QuestionsArray(i, j)
                     Next j
                Next i
                Label5.Caption = "4"
                StudentQuizArrangement = 4
             End If
            ' /444444444444444444444444
      
' ---------------------- Rearranging the Questions
           
      
             ' 55555555555555555555555
             If varRoute = 5 Then
                 DivideArray = Int(LengthofArray / 2)
                 If LengthofArray Mod 2 = 0 Then
                   Addition = 0
                  Else
                   Addition = 1
                 End If
                 For i = 1 To (DivideArray + Addition)
                    For j = 1 To 8
                       GlobalArray((DivideArray) + i, j) = QuestionsArray(i, j)
                    Next j
                 Next i
                 For i = 1 To DivideArray
                    For j = 1 To 8
                       GlobalArray(i, j) = QuestionsArray(DivideArray + i + Addition, j)
                    Next j
                 Next i
                 Label5.Caption = "5"
                 StudentQuizArrangement = 5
              End If
            ' /555555555555555555555555
            ' 6666666666666666666666666
             If varRoute = 6 Then
                 For i = 1 To LengthofArray
                   For j = 1 To 8
                    GlobalArray(LengthofArray - (i - 1), j) = QuestionsArray(i, j)
                   Next j
                Next i
              Label5.Caption = "6"
              StudentQuizArrangement = 6
             End If
            ' /666666666666666666666666
            ' 7777777777777777777777777
           If varRoute = 7 Then
              DivideArray = Int(LengthofArray / 2)
              If LengthofArray Mod 2 = 0 Then
                   Addition = 0
                  Else
                   Addition = 1
              End If
                 
              For i = 1 To DivideArray
                 For j = 1 To 8
                    GlobalArray(LengthofArray - (i - 1), j) = QuestionsArray((LengthofArray - DivideArray) + i, j)
                 Next j
              Next i
              For i = 1 To (DivideArray + Addition)
                 For j = 1 To 8
                    GlobalArray((DivideArray + Addition) - (i - 1), j) = QuestionsArray(i, j)
                 Next j
              Next i
              Label5.Caption = "7"
              StudentQuizArrangement = 7
            End If
            ' /7777777777777777777777777
            ' 88888888888888888888888888
           If varRoute = 8 Then
              For i = 1 To LengthofArray
                   For j = 1 To 8
                    GlobalArray(LengthofArray - (i - 1), j) = QuestionsArray(i, j)
                   Next j
                Next i
              Label5.Caption = "8" 'hidden
              StudentQuizArrangement = 8
           End If
           ' /88888888888888888888888888
      
'---------------------------- Rearranging the questions
       
        For i = 1 To (Int(varUnfinishedSequence) + 1)
           If Len(GlobalArray(i, 8)) = 1 Then
               VarQuestionSequence = VarQuestionSequence & "0" & GlobalArray(i, 8)
             Else
               VarQuestionSequence = VarQuestionSequence & GlobalArray(i, 8)
           End If
        Next i
        GlobalItemNumber = varUnfinishedSequence
        lblItemNumber.Caption = GlobalItemNumber + 1
        lblQuestion.Caption = GlobalArray(lblItemNumber.Caption, 1)
        Option1.Caption = GlobalArray(lblItemNumber.Caption, 2)
        Option2.Caption = GlobalArray(lblItemNumber.Caption, 3)
        Option3.Caption = GlobalArray(lblItemNumber.Caption, 4)
        Option4.Caption = GlobalArray(lblItemNumber.Caption, 5)
        Option5.Caption = GlobalArray(lblItemNumber.Caption, 6)
        VarQuizCompositeKey = Combo1.Text
        VarStudentUsername = GlobalUserName
     End If
   End If
' /******************Continue
End Sub



