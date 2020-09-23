VERSION 5.00
Begin VB.Form FrmCopyQuizForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy Quiz Form"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdViewQs 
      Caption         =   "&View Qs"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopyQuiz 
      Caption         =   "&Duplicate"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtFireMode 
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
      TabIndex        =   5
      Top             =   4560
      Width           =   5535
   End
   Begin VB.ListBox List_AllQuizForms 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   6615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Timer"
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No. of Items"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Expiry Date"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quiz Form"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   2775
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   6720
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   6720
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line6 
      X1              =   5640
      X2              =   6720
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   5640
      X2              =   6720
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   1440
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   1440
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   6720
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6720
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
   End
End
Attribute VB_Name = "FrmCopyQuizForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VarFireMode As String
Dim LastItem As Integer
Private Sub cmdClose_Click()
Unload Me
End Sub
Sub Populate_List()
List_AllQuizForms.Clear
With rsLimit
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
   List_AllQuizForms.AddItem !compositenumquestion & Space(25 - Len(!compositenumquestion)) & " | " & !dateend & Space(15 - Len(!dateend)) & " | " & !completed & Space(5 - Len(!completed)) & " | " _
       & !ftimers & " secs"
   .MoveNext
Loop
End With
End Sub

Private Sub cmdcopyquiz_Click()
Dim SearchString As String
Dim SearchChar As String
Dim VardateEnd As String
Dim VarCompleted As String
Dim PositionSearchChar As Integer
Dim LengthofRest As Integer
Dim JoinString As String
Dim LasItem As Integer
Dim VarItemNumber As Integer
Dim VarQuestion As String
Dim VarOption1 As String
Dim VarOption2 As String
Dim VarOption3 As String
Dim VarOption4 As String
Dim VarOption5 As String
Dim VarCorrectAnswer As String
Dim VarTimer As Integer
Dim Item As Integer
Item = 1
SearchChar = "|"
'MsgBox (Left(txtFireMode.Text, 20))
If txtFireMode.Text = "" Then
   MsgBox ("No quiz form selected! Please click one.")
   List_AllQuizForms.SetFocus
   Exit Sub
End If

VarFireMode = Left(txtFireMode.Text, 20)
With rsLimit
  .MoveFirst
  .Find "compositenumquestion= '" & Left(txtFireMode.Text, 20) & "'"
  If Not .EOF Then
      With rsLimit
         .MoveFirst
          SearchString = Left(txtFireMode.Text, 20)
          PositionSearchChar = InStr(1, SearchString, SearchChar)
          .Find "compositenumquestion = '" & UCase(Trim(Form1.txtUserName.Text)) & "|" & Trim(Right(SearchString, Len(SearchString) - PositionSearchChar)) & "'"
          If Not .EOF Then
             MsgBox ("You already have created a quiz form for the course! If you want to create a new set of questions, please delete first the previously created quiz form.")
             Exit Sub
          End If
      End With
  End If
End With
With rsLimit
  .MoveFirst
  .Find "compositenumquestion= '" & Left(txtFireMode.Text, 20) & "'"
  If Not .EOF Then
         If Len(!completed) = 5 Then
            LastItem = Right(!completed, 2)
         End If
         If Len(!completed) = 3 Then
            LastItem = Right(!completed, 1)
         End If
        SearchString = Left(txtFireMode.Text, 20)
        PositionSearchChar = InStr(1, SearchString, SearchChar)
        SearchString = Left(txtFireMode.Text, PositionSearchChar - 1)
        VarTimer = !ftimers
        VardateEnd = !dateend
        VarCompleted = !completed
        JoinString = UCase(Trim(Form1.txtUserName.Text) & (Right(Left(txtFireMode.Text, 20), 20 - Len(SearchString))))
        Call Populate_List

   End If
End With
With rsLimit
  If UCase(Form1.txtUserName.Text) <> SearchString Then
    .AddNew
    !compositenumquestion = Trim(JoinString)
    !dateend = VardateEnd
    !completed = VarCompleted
    !ftimers = VarTimer
    .Update
    MsgBox ("You have just duplicated a quiz form!")
    Call Populate_List
   Else
    MsgBox "You cannot duplicate your own quiz form!", vbInformation
    Exit Sub
  End If
End With
With rsQuestion
   If Not .EOF Then
     .MoveFirst
   End If
End With
With rsQuestion
  .MoveFirst
  Do While Item <= LastItem
     If Not .EOF Then
        If (Trim(!compositenumquestion) Like Trim(VarFireMode)) And _
            !itemnumber = Item Then
           VarItemNumber = !itemnumber
           VarQuestion = !question
           VarOption1 = !Option1
           VarOption2 = !Option2
           VarOption3 = !Option3
           VarOption4 = !Option4
           VarOption5 = !Option5
           VarCorrectAnswer = !correctanswer
           .AddNew
           !compositenumquestion = Trim(JoinString)
           !itemnumber = VarItemNumber
           !question = VarQuestion
           !Option1 = VarOption1
           !Option2 = VarOption2
           !Option3 = VarOption3
           !Option4 = VarOption4
           !Option5 = VarOption5
           !correctanswer = VarCorrectAnswer
           .Update
           .MoveFirst
           Item = Item + 1
        End If
     End If
    .MoveNext
 Loop
End With
End Sub

Private Sub cmdViewQs_Click()
If txtFireMode.Text <> "" Then
   frmViewQuestions.Show 1
 Else
   MsgBox ("No quiz form selected! Please click one.")
   List_AllQuizForms.SetFocus
End If
End Sub

Private Sub Form_Load()
strCnLimit = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Limit = New ADODB.Connection
strCn1Limit.Open strCnLimit
Call Initialization_Limit
Call Populate_List

strCnQu = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Qu = New ADODB.Connection
strCn1Qu.Open strCnQu
Call Initialization_Question
End Sub
Private Sub List_allquizforms_Click()
txtFireMode.Text = List_AllQuizForms.Text
End Sub



