VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faculty Server 2.12.10"
   ClientHeight    =   6465
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10275
   ControlBox      =   0   'False
   Icon            =   "QuizServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   6000
   End
   Begin VB.TextBox txtuserName 
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
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login:"
      Height          =   1335
      Left            =   1200
      TabIndex        =   4
      Top             =   5040
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "&Enter"
         Enabled         =   0   'False
         Height          =   855
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   6000
      Width           =   3735
   End
   Begin VB.Line Line60 
      X1              =   120
      X2              =   10200
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line59 
      X1              =   120
      X2              =   10200
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line58 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   10200
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line57 
      X1              =   7560
      X2              =   10200
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line56 
      BorderColor     =   &H80000005&
      X1              =   6840
      X2              =   10200
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line55 
      X1              =   8880
      X2              =   10200
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line54 
      BorderColor     =   &H80000005&
      X1              =   8880
      X2              =   10200
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line53 
      X1              =   8640
      X2              =   10200
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line52 
      BorderColor     =   &H80000005&
      X1              =   8640
      X2              =   10200
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line51 
      X1              =   8400
      X2              =   10200
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line50 
      BorderColor     =   &H80000005&
      X1              =   8520
      X2              =   10200
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line49 
      BorderColor     =   &H80000005&
      X1              =   2160
      X2              =   2760
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line48 
      X1              =   2040
      X2              =   2640
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line47 
      X1              =   1920
      X2              =   2520
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line46 
      BorderColor     =   &H80000005&
      X1              =   1920
      X2              =   2520
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line45 
      X1              =   2760
      X2              =   6360
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line44 
      BorderColor     =   &H80000005&
      X1              =   2760
      X2              =   6480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line43 
      X1              =   120
      X2              =   10200
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line42 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   10200
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line41 
      X1              =   120
      X2              =   1680
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line40 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   1800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line39 
      X1              =   120
      X2              =   1560
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line38 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   1680
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line37 
      X1              =   120
      X2              =   1560
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line36 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   1560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line35 
      X1              =   120
      X2              =   1440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line34 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   1440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line33 
      X1              =   8160
      X2              =   10200
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line32 
      X1              =   120
      X2              =   7680
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line31 
      BorderColor     =   &H80000005&
      X1              =   8040
      X2              =   10200
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line30 
      X1              =   120
      X2              =   10200
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line29 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   7560
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line28 
      X1              =   8880
      X2              =   10200
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line27 
      BorderColor     =   &H80000005&
      X1              =   3480
      X2              =   10200
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line26 
      X1              =   3360
      X2              =   10200
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line25 
      X1              =   120
      X2              =   2400
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line24 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   10200
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line23 
      X1              =   120
      X2              =   2160
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line22 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   2160
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line21 
      X1              =   120
      X2              =   2280
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line20 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   2160
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line19 
      X1              =   120
      X2              =   2040
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line18 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   2040
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line17 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   2040
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line16 
      X1              =   8400
      X2              =   10200
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line15 
      BorderColor     =   &H80000005&
      X1              =   8880
      X2              =   10200
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line14 
      X1              =   8400
      X2              =   10200
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000005&
      X1              =   8400
      X2              =   10200
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line12 
      X1              =   120
      X2              =   10200
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000005&
      X1              =   8280
      X2              =   10200
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000005&
      X1              =   4440
      X2              =   6360
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line9 
      X1              =   4440
      X2              =   6360
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   4440
      X2              =   10200
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line7 
      X1              =   4440
      X2              =   6360
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   4440
      X2              =   10200
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line5 
      X1              =   4440
      X2              =   10200
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   4440
      X2              =   10200
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line3 
      X1              =   4440
      X2              =   10200
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   10200
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10200
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Faculty Server"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   90
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4335
      Left            =   840
      TabIndex        =   6
      Top             =   480
      Width           =   9135
   End
   Begin VB.Label Label1 
      Caption         =   "User Name:"
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
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
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
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Menu mnuRunQuizServer 
      Caption         =   "&Run Quiz Server"
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Begin VB.Menu mnuFaculty 
         Caption         =   "&Faculty"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCourse 
         Caption         =   "&Course"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStudent 
         Caption         =   "&Student"
      End
      Begin VB.Menu mnuQuiz 
         Caption         =   "&Quiz Form"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuDuplicate 
      Caption         =   "&Duplicate "
      Enabled         =   0   'False
      Begin VB.Menu mnuDuplicateSQuestions 
         Caption         =   "&Select Questions Only"
      End
      Begin VB.Menu mnuDuplicateQuizForm 
         Caption         =   "&Quiz Form"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Enabled         =   0   'False
      Begin VB.Menu mnuViewFaculty 
         Caption         =   "All &Faculty Members"
      End
      Begin VB.Menu mnuViewCourse 
         Caption         =   "All &Courses"
      End
      Begin VB.Menu mnuViewStudent 
         Caption         =   "All &Students"
      End
      Begin VB.Menu mnuQuizForm 
         Caption         =   "All &Quiz Forms"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuViewQuizResults 
         Caption         =   "Quiz &Results"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuViewFacultyEnglishViolators 
         Caption         =   "Faculty E&nglish Violators"
      End
      Begin VB.Menu mnuViewStudentEnglishViolators 
         Caption         =   "S&tudent English Violators"
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilities"
      Enabled         =   0   'False
      Begin VB.Menu mnuUtilitiescreateBAckup 
         Caption         =   "Create Bac&kup"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUtilitiesUseBackup 
         Caption         =   "Use &Backup"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPassword 
         Caption         =   "Change &Password"
      End
   End
   Begin VB.Menu mnuEnglishViolators 
      Caption         =   "&English Violators"
      Enabled         =   0   'False
      Begin VB.Menu mnuViolatorsFaculty 
         Caption         =   "&Faculty"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViolatorsStudents 
         Caption         =   "St&udents"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "A&bout"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim First_Name As String
Private Sub Command1_Click()
strCn = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1 = New ADODB.Connection
strCn1.Open strCn
Call Initialization
With rsFaculty
  .MoveFirst
  .Find "username= '" & UCase(Trim(Form1.txtuserName.Text)) & "'"
  If Not .EOF Then
      If !facultyPassword = UCase(Form1.txtPassword.Text) Then
           If !firstvisit = 0 Then
              frmChangePassword.Show (1)
              !firstvisit = 1
              !facultyPassword = Trim(facultyPassword)
           End If
           If !Level = 3 Then
              mnuUtilitiescreateBAckup.Enabled = True
              mnuUtilitiesUseBackup.Enabled = True
              mnuFaculty.Enabled = True
              mnuCourse.Enabled = True
              mnuViolatorsFaculty.Enabled = True
              mnuViolatorsStudents.Enabled = True
           End If
           If !Level = 2 Then
              mnuUtilitiescreateBAckup.Enabled = True
              mnuUtilitiesUseBackup.Enabled = True
              mnuFaculty.Enabled = True
              mnuCourse.Enabled = True
              mnuViolatorsFaculty.Enabled = True
              mnuViolatorsStudents.Enabled = True
           End If
           If !Level = 1 Then
              mnuViolatorsFaculty.Enabled = True
              mnuViolatorsStudents.Enabled = True
           End If
           mnuRunQuizServer.Enabled = True
           GlobalFacultyUserRights = !Level
           mnuDuplicate.Enabled = True
           mnuEnglishViolators.Enabled = True
           First_Name = !facultyfname
           mnuView.Enabled = True
           mnuAdd.Enabled = True
           mnuUtilities.Enabled = True
           txtuserName.Enabled = False
           txtPassword.Enabled = False
           Command1.Enabled = False
           .Update
           MsgBox "Welcome to the Faculty Server, " & First_Name & "!"
       Else
           MsgBox "Please check your Password", vbExclamation
           Form1.txtPassword.Text = ""
           Form1.txtPassword.SetFocus
       End If
   Else
       MsgBox "Username does not exist.", vbExclamation
       Form1.txtuserName.Text = ""
       Form1.txtPassword.Text = ""
       Form1.Command1.Enabled = False
       Form1.txtuserName.SetFocus
   End If
End With
End Sub

'Private Sub Label4_Change()
'Dim Si As Integer
'Dim Getlast As Integer
'Si = 1
'Getlast = Right(Second(Time), 1)
'Shape1(Getlast).Visible = True
'If Getlast Mod 2 = 1 Then
 ' Si = 10000
 ' Do While Si >= 1
 ' Si = Si + 50
 '   Shape1(Getlast).Left = Si
 ' Loop
'End If
'If Getlast Mod 2 = 0 Then
 ' Do While Si <= 10000
 '   Si = Si + 50
 '   Shape1(Getlast).Left = Si
 ' Loop
'End If
'Shape1(Getlast).Visible = False
'Shape1(Getlast).Left = 1
'End Sub

Private Sub mnuAbout_Click()
frmAbout.Show (1)
End Sub

Private Sub mnuCourse_Click()
frmCourse.Show (1)
End Sub

Private Sub mnuDuplicateSQuestions_Click()
frmDuplicateSQuestions.Show 1
End Sub

Private Sub mnuExit_Click()
End
End Sub
Private Sub mnuFaculty_Click()
frmFaculty.Show (1)
End Sub
Private Sub mnuPassword_Click()
frmUtilityChangePassword.Show (1)
End Sub
Private Sub mnuQuiz_Click()
frmQuizForm.Show (1)
End Sub
Private Sub mnuQuizForm_Click()
frmAllQuizForms.Show (1)
End Sub

Private Sub mnuDuplicateQuizForm_Click()
FrmCopyQuizForm.Show (1)
End Sub

Private Sub mnuRunQuizServer_Click()
frmQuizServerRun.Show (1)
End Sub

Private Sub mnuStudent_Click()
frmStudent.Show (1)
End Sub

Private Sub mnuUtilitiescreateBAckup_Click()
MyAppID = Shell("D:\Academic Head\quizserver\backup.bat")
AppActivate MyAppID
MsgBox "Backup has just been created!"

End Sub

Private Sub mnuUtilitiesUseBackup_Click()
MyAppID = Shell("D:\Academic Head\quizserver\usebak.bat")
AppActivate MyAppID
MsgBox "Backup has just been used! But please check the DOS window first."
End Sub

Private Sub mnuViewCourse_Click()
frmAllCourses.Show (1)
End Sub
Private Sub mnuViewFaculty_Click()
frmAllFacultyMembers.Show (1)
End Sub

Private Sub mnuViewQuestions_Click()
strCnQu = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Qu = New ADODB.Connection
strCn1Qu.Open strCnQu
Call Initialization_Question
frmViewQuizForm.Show (1)
End Sub

Private Sub mnuViewFacultyEnglishViolators_Click()
frmViewFAcultyViolators.Show (1)
End Sub

Private Sub mnuViewQuizResults_Click()
strCnQ = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Q = New ADODB.Connection
strCn1Q.Open strCnQ
Call Initialization_Quiz

strCnLimit = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Limit = New ADODB.Connection
strCn1Limit.Open strCnLimit
Call Initialization_Limit

frmQuizResults.Show (1)
End Sub

Private Sub mnuViewStudent_Click()
frmAllStudents.Show (1)
End Sub

Private Sub mnuViewStudentEnglishViolators_Click()
frmViewStudentViolators.Show (1)
End Sub

Private Sub mnuViolatorsFaculty_Click()
frmFacultyViolations.Show (1)
End Sub

Private Sub mnuViolatorsStudents_Click()
frmStudentViolations.Show (1)
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Date & "    " & Time

End Sub

Private Sub txtPassword_Change()
If txtuserName.Text <> "" And txtPassword.Text <> "" Then
   Command1.Enabled = True
End If
End Sub

Private Sub txtUserName_Change()
If txtuserName.Text <> "" And txtPassword.Text <> "" Then
   Command1.Enabled = True
End If
End Sub
