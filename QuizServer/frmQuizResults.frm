VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQuizResults 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Quiz Results"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtFiremode 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4800
      Width           =   3855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   720
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   3855
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Name"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Fire Mode:"
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
      TabIndex        =   7
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Line Line15 
      BorderColor     =   &H80000009&
      X1              =   5040
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line14 
      X1              =   4200
      X2              =   5040
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000009&
      X1              =   4200
      X2              =   5040
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line12 
      X1              =   4200
      X2              =   5040
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      X1              =   4200
      X2              =   5040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line10 
      X1              =   4200
      X2              =   5040
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   1440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   1440
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   1440
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   1440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line5 
      X1              =   4440
      X2              =   5040
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   4440
      X2              =   5040
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   600
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   600
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5040
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label1 
      Caption         =   "Pick Quiz Form:"
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
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmQuizResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Varlimit As Integer

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
With rsQuiz
  .MoveFirst
  .Find "quizcompositekey = '" & Trim(Combo1.Text) & "'"
  While Not .EOF
      If !studentusername = Trim(Left(txtFiremode.Text, 15)) Then
          If MsgBox("Do you really want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
              .Delete
              Call Populate_List
              txtFiremode.Text = ""
              Exit Sub
            Else
              Exit Sub
          End If
         Else
          .MoveNext
          .Find "quizcompositekey = '" & Trim(Combo1.Text) & "'"
       End If
  Wend
End With
End Sub

Private Sub cmdPrint_Click()
Dim I As Integer
Dim a, b, c, d, e As String
Dim BeginPage, EndPage, NumCopies, j
Dim NumSpaces As Integer
Dim SpaceChar As String
Dim x As Integer
Dim OneChar As String
'Set Cancel to True
CommonDialog1.CancelError = True
'On Error GoTo errhandler
'Display the Prin dialog box
CommonDialog1.ShowPrinter
'Get user-selected values from the dialog box
BeginPage = CommonDialog1.FromPage
EndPage = CommonDialog1.ToPage
NumCopies = CommonDialog1.Copies

strCnQuery = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Query = New ADODB.Connection
strCn1Query.Open strCnQuery
Call Initialization_QuizQuery



With rsLimit
   .MoveFirst
   .Find "compositenumquestion = '" & Trim(Combo1.Text) & "'"
   If Not .EOF Then
       If Len(!completed) = 5 Then
            Varlimit = Right(!completed, 2)
       End If
       If Len(!completed) = 3 Then
            Varlimit = Right(!completed, 1)
       End If
   End If
End With

For j = 1 To NumCopies
I = 1
Printer.FontSize = 20
Printer.Print " "
Printer.Print " "
Printer.Print Tab(18); "STI QUIZ RESULT"
Printer.FontSize = 12
Printer.Print "     "
Printer.Print "     "
Printer.Print Tab(10); "Quiz Form: "; Combo1.Text; "          Total No. of Items: "; Varlimit;
Printer.Print "     "
Printer.Print "     "
Printer.Print Tab(10); "Name of Student "; "          Quiz Score"; "    Signature of Student";
Printer.Print "     "
Printer.FontName = "courier new"
With rsQuizQuery
   .MoveFirst
End With
With rsQuizQuery
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
    SpaceChar = ""
   If !quizcompositekey = Trim(Combo1.Text) Then
       NumSpaces = 25 - Len(!studentusername)
       For x = 1 To NumSpaces
          SpaceChar = SpaceChar & " "
       Next x
       If Len(!quizscore) = 1 Then
         OneChar = " "
        Else
         OneChar = ""
       End If
       Printer.Print Tab(10); " "; !studentusername; SpaceChar; " | "; OneChar; !quizscore; "         ____________________";
   End If
   .MoveNext
Loop
End With
Printer.Print "     "
Printer.Print "     "
Printer.Print "     "
Printer.Print "     "
Printer.Print "     "
Printer.Print Tab(25); "                     _______________"; Date;
Printer.Print Tab(25); "                        Instructor ";

Next j


Printer.EndDoc
 Exit Sub
errhandler:
  Exit Sub
End Sub

Private Sub cmdView_Click()
List1.Clear
Call Populate_List
End Sub

Private Sub Form_Load()
Dim LocalRemainingString As String
Dim LocalMypos As Integer
Dim LocalSearchChar As String

strCnQ = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Q = New ADODB.Connection
strCn1Q.Open strCnQ
Call Initialization_Quiz

LocalSearchChar = "|"
Combo1.Clear
With rsLimit
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
   LocalRemainingString = !compositenumquestion
   LocalMypos = InStr(1, LocalRemainingString, LocalSearchChar)
   
   If UCase(Form1.txtUserName.Text) = Left(LocalRemainingString, LocalMypos - 1) Then
      Combo1.AddItem !compositenumquestion
   End If
   .MoveNext
Loop
End With
End Sub


Sub Populate_List()
List1.Clear
With rsQuiz
   .MoveFirst
End With
With rsQuiz
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
   If !quizcompositekey = Trim(Combo1.Text) Then
       List1.AddItem !studentusername & Space(25 - Len(!studentusername)) & " | " & !quizscore
   End If
   .MoveNext
Loop
End With
End Sub

Private Sub List1_Click()
txtFiremode.Text = List1.Text
End Sub
