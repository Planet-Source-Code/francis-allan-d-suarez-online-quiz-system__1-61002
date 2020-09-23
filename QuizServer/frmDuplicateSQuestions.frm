VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmDuplicateSQuestions 
   Caption         =   "Copy Selected Questions"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboTimer1 
      Height          =   315
      ItemData        =   "frmDuplicateSQuestions.frx":0000
      Left            =   6000
      List            =   "frmDuplicateSQuestions.frx":002B
      TabIndex        =   19
      Text            =   "30"
      Top             =   6480
      Width           =   615
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2175
      Left            =   0
      TabIndex        =   17
      Top             =   5640
      Width           =   4455
      _Version        =   524288
      _ExtentX        =   7858
      _ExtentY        =   3836
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2004
      Month           =   12
      Day             =   10
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   3360
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "C&reate Quiz Form"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   7200
      Width           =   2775
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   8280
      TabIndex        =   15
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   240
      Width           =   2055
   End
   Begin VB.ComboBox ComboTotalItems 
      Height          =   315
      ItemData        =   "frmDuplicateSQuestions.frx":0063
      Left            =   2400
      List            =   "frmDuplicateSQuestions.frx":0082
      TabIndex        =   12
      Text            =   "10"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtFireMode2 
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5880
      Width           =   3975
   End
   Begin VB.TextBox txtFireMode1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3360
      Width           =   3975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy "
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   3960
      Width           =   9255
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
      Height          =   1530
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   9255
   End
   Begin VB.CommandButton cmdDisplayQuestions 
      Caption         =   "&Display Questions"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox ComboQuizForm 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Line Line31 
      BorderColor     =   &H80000005&
      X1              =   1680
      X2              =   120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line30 
      X1              =   1680
      X2              =   120
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line29 
      BorderColor     =   &H80000005&
      X1              =   5280
      X2              =   120
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line28 
      X1              =   5280
      X2              =   4440
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line27 
      BorderColor     =   &H80000005&
      X1              =   4440
      X2              =   9360
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line26 
      X1              =   4440
      X2              =   5400
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line25 
      BorderColor     =   &H80000005&
      X1              =   4440
      X2              =   5400
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line24 
      X1              =   6600
      X2              =   9360
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line23 
      BorderColor     =   &H80000005&
      X1              =   6600
      X2              =   9360
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line22 
      X1              =   4440
      X2              =   5880
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line21 
      BorderColor     =   &H80000005&
      X1              =   4440
      X2              =   5880
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line20 
      X1              =   6720
      X2              =   9360
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line19 
      BorderColor     =   &H80000005&
      X1              =   6720
      X2              =   9360
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line18 
      X1              =   4440
      X2              =   5880
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line17 
      BorderColor     =   &H80000005&
      X1              =   4440
      X2              =   5280
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line16 
      BorderColor     =   &H80000005&
      X1              =   4440
      X2              =   5280
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line15 
      X1              =   9360
      X2              =   4440
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line14 
      X1              =   9360
      X2              =   6600
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000005&
      X1              =   9360
      X2              =   6600
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line12 
      X1              =   120
      X2              =   4080
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4080
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line10 
      X1              =   120
      X2              =   4080
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000005&
      X1              =   6600
      X2              =   9360
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   5280
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   9360
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line6 
      X1              =   3120
      X2              =   9360
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   3120
      X2              =   9360
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line4 
      X1              =   4440
      X2              =   9360
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   7800
      X2              =   9360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      X1              =   7800
      X2              =   9360
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   9360
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label7 
      Caption         =   "secs."
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
      Left            =   6000
      TabIndex        =   20
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Time:"
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
      Left            =   5400
      TabIndex        =   18
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Total Number of Items:"
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
      TabIndex        =   11
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Fire Mode 2:"
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
      Left            =   5400
      TabIndex        =   9
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Fire Mode 1:"
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
      Left            =   5400
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Select Quiz Form:"
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
Attribute VB_Name = "frmDuplicateSQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VarListCtr As Integer
Dim VarRow As Integer
Dim ArrQuestions(50, 8)
Dim RecordCtr As Integer



Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCopy_Click()

Dim TargetNum As String
Dim AddZero As String
cmdCreate.Enabled = True
'ReDim ArrQuestions(Val(ComboTotalItems.Text), 8)
RecordCtr = RecordCtr + 1
TargetNum = Left(txtFireMode1.Text, 2)
With rsQuestion
   .MoveFirst
   .Find "compositenumquestion= '" & Trim(ComboQuizForm.Text) & "'"
  While Not .EOF
    If Left(TargetNum, 1) = 0 Then
       TargetNum = Right(TargetNum, 1)
    End If
    If !itemnumber = TargetNum Then
'       VarRow = VarRow + 1
       VarListCtr = VarListCtr + 1
       If Len(Trim(VarListCtr)) = 1 Then
          AddZero = "0"
         Else
          AddZero = ""
       End If
      If VarRow < Val(ComboTotalItems.Text) Then
       VarRow = VarRow + 1
       List2.AddItem AddZero & VarListCtr & Space(3 - Len(VarListCtr)) & " | " & _
      !question & Space(250 - Len(!question)) & " | " & _
      !Option1 & Space(50 - Len(!Option1)) & " | " & _
      !Option2 & Space(50 - Len(!Option2)) & " | " & _
      !Option3 & Space(50 - Len(!Option3)) & " | " & _
      !Option4 & Space(50 - Len(!Option4)) & " | " & _
      !Option5 & Space(50 - Len(!Option5)) & " | " & _
      !correctanswer & Space(50 - Len(!correctanswer))
       ArrQuestions(VarRow, 1) = VarListCtr
       ArrQuestions(VarRow, 2) = !question
       ArrQuestions(VarRow, 3) = !Option1
       ArrQuestions(VarRow, 4) = !Option2
       ArrQuestions(VarRow, 5) = !Option3
       ArrQuestions(VarRow, 6) = !Option4
       ArrQuestions(VarRow, 7) = !Option5
       ArrQuestions(VarRow, 8) = !correctanswer
       
        Else
          MsgBox "You have specified only " & Val(ComboTotalItems.Text) & " questions for the quizform!", vbExclamation
      End If
        
        
    End If
     .MoveNext
     .Find "compositenumquestion= '" & Trim(ComboQuizForm.Text) & "'"
  Wend
End With
List1.SetFocus
cmdCopy.Enabled = False
End Sub

Private Sub cmdCreate_Click()
Dim CreateI As Integer
Dim TargetCourse As String
'strCnLimit = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
'Set strCn1Limit = New ADODB.Connection
'strCn1Limit.Open strCnLimit
'Call Initialization_Limit
Dim Mypos As Integer
Dim SearchChar As String
SearchChar = "|"

Mypos = InStr(1, ComboQuizForm.Text, SearchChar, 1)
TargetCourse = Right(ComboQuizForm.Text, Len(ComboQuizForm.Text) - Mypos)
With rsLimit
  .MoveFirst
  .AddNew
  !compositenumquestion = Trim(txtuserName.Text) & "|" & TargetCourse
  !dateend = Calendar1.Value - 1
  !completed = Trim(RecordCtr) & "|" & Trim(ComboTotalItems.Text)
  !ftimers = ComboTimer1.Text
  .Update
End With
With rsQuestion
   For CreateI = 1 To RecordCtr
      .AddNew
      !compositenumquestion = Trim(txtuserName.Text) & "|" & TargetCourse
      !itemnumber = ArrQuestions(CreateI, 1)
      !question = ArrQuestions(CreateI, 2)
      !Option1 = ArrQuestions(CreateI, 3)
      !Option2 = ArrQuestions(CreateI, 4)
      !Option3 = ArrQuestions(CreateI, 5)
      !Option4 = ArrQuestions(CreateI, 6)
      !Option5 = ArrQuestions(CreateI, 7)
      !correctanswer = ArrQuestions(CreateI, 8)
      .Update
   Next CreateI
End With
If RecordCtr < Val(ComboTotalItems.Text) Then
   MsgBox "You have just created a quiz form, but incomplete!", vbInformation
  Else
   MsgBox "You have just created a complete quiz form!", vbInformation
End If
Timer1.Enabled = False
cmdCreate.Enabled = False
cmdCopy.Enabled = False
cmdDisplayQuestions.Enabled = False
End Sub

Private Sub cmdDisplayQuestions_Click()
Dim TargetCourse As String
strCnLimit = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Limit = New ADODB.Connection
strCn1Limit.Open strCnLimit
Call Initialization_Limit
Dim Mypos As Integer
Dim SearchChar As String
Timer1.Enabled = True
SearchChar = "|"
RecordCtr = 0
Mypos = InStr(1, ComboQuizForm.Text, SearchChar, 1)
TargetCourse = Right(ComboQuizForm.Text, Len(ComboQuizForm.Text) - Mypos)
With rsLimit
  .MoveFirst
  .Find "compositenumquestion= '" & Trim(txtuserName.Text) & "|" & TargetCourse & "'"
  If Not .EOF Then
     MsgBox "You have already tried creating or created this quizform!", vbExclamation
     cmdCopy.Enabled = False
     ComboQuizForm.SetFocus
     Timer1.Enabled = False
     Exit Sub
  End If
End With
Call Populate_List
'CmdGo.Enabled = True
List1.SetFocus
cmdCopy.Enabled = False
End Sub



Private Sub Form_Load()
strCnLimit = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Limit = New ADODB.Connection
strCn1Limit.Open strCnLimit
Call Initialization_Limit
'Call Populate_List

strCnQu = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Qu = New ADODB.Connection
strCn1Qu.Open strCnQu
Call Initialization_Question
Call Populate_ComboQuizForm

txtuserName.Text = Trim(UCase(Form1.txtuserName.Text))
End Sub
Sub Populate_ComboQuizForm()
With rsLimit
   .MoveFirst
   While Not .EOF
      ComboQuizForm.AddItem !compositenumquestion
      .MoveNext
   Wend
End With
End Sub
Sub Populate_List()
Dim AddZero As String
Dim AddSpace As String
Dim ValueString As String
ValueString = Trim(ComboQuizForm.Text)
List1.Clear
With rsQuestion
  .MoveFirst
  '.Find "compositenumquestion= '" & ValueString & "'"
  If Not .EOF Then
     .MoveFirst
  End If
  Do Until .EOF
   If !compositenumquestion = ValueString Then
      If Len(!itemnumber) = 1 Then
         AddZero = "0"
         AddSpace = ""
        Else
         AddSpace = " "
         AddZero = ""
      End If
      List1.AddItem AddZero & !itemnumber & Space(3 - Len(!itemnumber)) & AddSpace & " | " & _
      !question & Space(250 - Len(!question)) & " | " & _
      !Option1 & Space(50 - Len(!Option1)) & " | " & _
      !Option2 & Space(50 - Len(!Option2)) & " | " & _
      !Option3 & Space(50 - Len(!Option3)) & " | " & _
      !Option4 & Space(50 - Len(!Option4)) & " | " & _
      !Option5 & Space(50 - Len(!Option5)) & " | " & _
      !correctanswer & Space(50 - Len(!correctanswer))
      
   End If
   .MoveNext
Loop
End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Do you want to continue with the incomplete quiz form?", vbYesNo + vbQuestion) = vbYes Then
  frmQuizForm.Show (1)
End If
End Sub

Private Sub List1_Click()
txtFireMode1.Text = List1.Text
End Sub

Private Sub List2_Click()
txtFireMode2.Text = List2.Text
End Sub
'Private Sub cmdRemove_Click()
'Dim TargetIndex As String
'Dim VarI As Integer
'Dim SubVarI As Integer
'Dim AddZero As String
'ReDim ArrQuestions(ComboTotalItems.Text, 8)
'TargetIndex = Trim(Left(List2.Text, 2))

'For VarI = 1 To Val(ComboTotalItems.Text)
'   If VarI = TargetIndex Then
'      'SubVarI = VarI
'      For SubVarI = 1 To (VarI - 1)
'         If (SubVarI = 1 Or SubVarI = 2 Or SubVarI = 3 Or SubVarI = 4 _
'            Or SubVarI = 5 Or SubVarI = 6 Or SubVarI = 8 Or SubVarI = 9 _
'            Or SubVarI = 7) And Len(SubVarI) < 2 Then
'            AddZero = "0"
'           'Else
'            'AddZero = ""
'          End If
'          ArrQuestions(SubVarI, 1) = AddZero & SubVarI
'      Next SubVarI
'      For SubVarI = VarI To (Val(ComboTotalItems.Text) - 1)
'          If SubVarI = 1 Or SubVarI = 2 Or SubVarI = 3 Or SubVarI = 4 _
'            Or SubVarI = 5 Or SubVarI = 6 Or SubVarI = 8 Or SubVarI = 9 _
'            Or SubVarI = 7 And Len(SubVarI) < 2 Then
'            AddZero = "0"
'           'Else
'           ' AddZero = ""
'          End If
'          ArrQuestions(SubVarI, 1) = AddZero & ArrQuestions(SubVarI, 1)
'          ArrQuestions(SubVarI, 2) = ArrQuestions(SubVarI + 1, 2)
'          ArrQuestions(SubVarI, 3) = ArrQuestions(SubVarI + 1, 3)
'          ArrQuestions(SubVarI, 4) = ArrQuestions(SubVarI + 1, 4)
'          ArrQuestions(SubVarI, 5) = ArrQuestions(SubVarI + 1, 5)
'          ArrQuestions(SubVarI, 6) = ArrQuestions(SubVarI + 1, 6)
 '         ArrQuestions(SubVarI, 7) = ArrQuestions(SubVarI + 1, 7)
 '         ArrQuestions(SubVarI, 8) = ArrQuestions(SubVarI + 1, 8)
'      Next SubVarI
'
 '  End If
'Next VarI
'Call Populate_List2
'End Sub
'Sub Populate_List2()
'Dim I As Integer
'List2.Clear
'For I = 1 To Val(ComboTotalItems.Text)
'  List2.AddItem ArrQuestions(I, 1) & Space(3 - Len(ArrQuestions(I, 1))) & " | " & _
'      ArrQuestions(I, 2) & Space(250 - Len(ArrQuestions(I, 2))) & " | " & _
'      ArrQuestions(I, 3) & Space(50 - Len(ArrQuestions(I, 3))) & " | " & _
'      ArrQuestions(I, 4) & Space(50 - Len(ArrQuestions(I, 4))) & " | " & _
'      ArrQuestions(I, 5) & Space(50 - Len(ArrQuestions(I, 5))) & " | " & _
'      ArrQuestions(I, 6) & Space(50 - Len(ArrQuestions(I, 6))) & " | " & _
'      ArrQuestions(I, 7) & Space(50 - Len(ArrQuestions(I, 7))) & " | " & _
'      ArrQuestions(I, 8) & Space(50 - Len(ArrQuestions(I, 1)))
'Next I
'End Sub

Private Sub Timer1_Timer()
cmdCopy.Enabled = True
End Sub

