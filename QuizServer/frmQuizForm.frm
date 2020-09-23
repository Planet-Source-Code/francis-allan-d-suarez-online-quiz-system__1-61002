VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmQuizForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quiz Form"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmQuizForm.frx":0000
      Left            =   3480
      List            =   "frmQuizForm.frx":002B
      TabIndex        =   3
      Text            =   "30"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdNewQuizform 
      Caption         =   "New &Quiz Form"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   32
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   2640
      Width           =   1575
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   1935
      Left            =   4080
      TabIndex        =   29
      Top             =   120
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   3413
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2004
      Month           =   11
      Day             =   3
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtCorrectAnswer 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   28
      Top             =   3600
      Width           =   2535
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Option5"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   5040
      Width           =   255
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   4560
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   4080
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   3600
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   3120
      Width           =   255
   End
   Begin VB.TextBox txtOption5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox txtOption4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox txtOption3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txtOption2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtOption1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox txtQuestion 
      Enabled         =   0   'False
      Height          =   735
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   5895
   End
   Begin VB.CommandButton cmdActivate 
      Caption         =   "&Activate"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmQuizForm.frx":0063
      Left            =   1560
      List            =   "frmQuizForm.frx":00B5
      TabIndex        =   2
      Text            =   "5"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   240
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label9 
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
      Left            =   3480
      TabIndex        =   34
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label8 
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
      Left            =   2880
      TabIndex        =   33
      Top             =   1080
      Width           =   615
   End
   Begin VB.Line Line16 
      BorderColor     =   &H80000009&
      X1              =   8160
      X2              =   4200
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line15 
      X1              =   8160
      X2              =   4200
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000009&
      X1              =   8160
      X2              =   6000
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line13 
      X1              =   7200
      X2              =   8160
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000009&
      X1              =   7200
      X2              =   8160
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line11 
      X1              =   7200
      X2              =   8160
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      X1              =   8160
      X2              =   4200
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line9 
      X1              =   8160
      X2              =   4200
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   8160
      X2              =   6840
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line7 
      X1              =   8160
      X2              =   6840
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   8160
      X2              =   6840
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line5 
      X1              =   8160
      X2              =   6840
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   8160
      X2              =   6840
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line3 
      X1              =   4200
      X2              =   8160
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label7 
      Caption         =   "Quiz Form Expires On:"
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
      Left            =   6720
      TabIndex        =   31
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Correct Answer:"
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
      Left            =   4560
      TabIndex        =   27
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Option 5:"
      Height          =   375
      Left            =   600
      TabIndex        =   26
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Option 4:"
      Height          =   375
      Left            =   600
      TabIndex        =   25
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label lblOptn3 
      Caption         =   "Option 3:"
      Height          =   375
      Left            =   600
      TabIndex        =   24
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label LblOptn2 
      Caption         =   "Option 2:"
      Height          =   375
      Left            =   600
      TabIndex        =   23
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label LblOptn1 
      Caption         =   "Option 1:"
      Height          =   495
      Left            =   600
      TabIndex        =   22
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label LblNo 
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
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6720
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6720
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label3 
      Caption         =   "No. of Items:"
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
      TabIndex        =   20
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Course Code:"
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
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmQuizForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Calendar1_Click()
Text1.Text = Calendar1.Value
End Sub
Private Sub cmdAction_Click()
cmdClose.Enabled = False
If txtOption1.Text = "" Then
  txtOption1.SetFocus
End If
If txtOption2.Text = "" Then
  txtOption2.SetFocus
End If
If txtOption3.Text = "" Then
  txtOption3.SetFocus
End If
If txtOption4.Text = "" Then
  txtOption4.SetFocus
End If
If txtOption5.Text = "" Then
  txtOption5.SetFocus
End If
If txtOption1.Text <> "" And txtOption2.Text <> "" And _
txtOption3.Text <> "" And txtOption4.Text <> "" And _
txtOption5.Text <> "" Then
If Qctr <= Val(Combo2.Text) Then
   With rsQuestion
      .AddNew
      !compositenumquestion = txtUserName.Text & "|" & Combo1.Text
      !itemnumber = Qctr
      !question = txtQuestion.Text
      !Option1 = txtOption1.Text
      !Option2 = txtOption2.Text
      !Option3 = txtOption3.Text
      !Option4 = txtOption4.Text
      !Option5 = txtOption5.Text
      !correctanswer = txtCorrectAnswer.Text
      .Update
    End With
   If Qctr = 1 Then
     With rsLimit
        .AddNew
        !compositenumquestion = txtUserName.Text & "|" & Combo1.Text
        !dateend = Text1.Text
        !ftimers = Combo3.Text
        .Update
     End With
   End If
   With rsLimit
      .MoveFirst
      .Find "compositenumquestion= '" & Trim(txtUserName.Text) & "|" & Trim(Combo1.Text) & "'"
      If Not .EOF Then
        !completed = Qctr & "|" & Trim(Combo2.Text)
        .Update
      End If
   End With
txtQuestion.Text = ""
txtOption1.Text = ""
txtOption2.Text = ""
txtOption3.Text = ""
txtOption4.Text = ""
txtOption5.Text = ""
txtOption1.Visible = False
txtOption2.Visible = False
txtOption3.Visible = False
txtOption4.Visible = False
txtOption5.Visible = False
txtOption1.Visible = True
txtOption2.Visible = True
txtOption3.Visible = True
txtOption4.Visible = True
txtOption5.Visible = True
Option1.Refresh
Option2.Refresh
Option3.Refresh
Option4.Refresh
Option5.Refresh
Qctr = Qctr + 1
LblNo.Caption = Qctr & "."
txtQuestion.SetFocus
End If
End If
If Qctr > Val(Combo2.Text) Then
  cmdAction.Caption = "End"
  cmdNewQuizform.Enabled = True
  cmdAction.Enabled = False
  cmdClose.Enabled = True
  frmViewQuizForm.Show (1)
End If
End Sub
Private Sub cmdActivate_Click()
cmdActivate.Enabled = False
txtOption1.Enabled = True
txtOption2.Enabled = True
txtOption3.Enabled = True
txtOption4.Enabled = True
txtOption5.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
txtCorrectAnswer.Enabled = True
cmdAction.Enabled = True
txtQuestion.Enabled = True
txtQuestion.SetFocus
LblNo.Caption = Qctr
strCnQu = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Qu = New ADODB.Connection
strCn1Qu.Open strCnQu
Call Initialization_Question
strCnLimit = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Limit = New ADODB.Connection
strCn1Limit.Open strCnLimit
Call Initialization_Limit
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNewQuizform_Click()
Qctr = 1
Combo1.SetFocus

End Sub

Private Sub Combo1_GotFocus()
cmdActivate.Enabled = False
End Sub
Private Sub Combo1_LostFocus()
strCnLimit = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Limit = New ADODB.Connection
strCn1Limit.Open strCnLimit
Call Initialization_Limit
With rsLimit
  .MoveFirst
  .Find "compositenumquestion= '" & Trim(txtUserName.Text) & "|" & Trim(Combo1.Text) & "'"
  If Not .EOF Then
    Combo3.Text = !ftimers
    If Len(!completed) > 3 Then
       Combo2.Text = Right(!completed, 2)
     Else
       Combo2.Text = Right(!completed, 1)
    End If
    If (Len(!completed) Mod 2) = 1 Then
        If (Len(!completed) = 5 Or Len(!completed) = 3) And (Left(!completed, 2) = Right(!completed, 2) Or Left(!completed, 1) = Right(!completed, 1)) Then
            MsgBox ("You have already completed the quiz form for this subject.")
            Combo1.SetFocus
        Else
            If Len(!completed) = 5 Then
               Qctr = Left(!completed, 2) + 1
            End If
            If Len(!completed) = 3 Then
               Qctr = Left(!completed, 1) + 1
            End If
            cmdActivate.Enabled = True
        End If
    Else
        If Len(!completed) = 5 Then
             Qctr = Left(!completed, 2) + 1
        End If
        If Len(!completed) = 4 Then
            Qctr = Left(!completed, 1) + 1
        End If
        If Len(!completed) = 3 Then
             Qctr = Left(!completed, 1) + 1
        End If
        cmdActivate.Enabled = True
    End If
  Else
    Qctr = 1
    cmdActivate.Enabled = True
   End If
End With
End Sub

Private Sub Form_Load()
Calendar1.Value = Date
Text1.Text = Calendar1.Value
Qctr = 1
txtUserName.Text = UCase(Form1.txtUserName.Text)
strCnC = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1C = New ADODB.Connection
strCn1C.Open strCnC
Call Initialization_Course
With rsCourse
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
   Combo1.AddItem !coursecode
   .MoveNext
Loop
End With
End Sub
Private Sub Option1_Click()
txtCorrectAnswer.Text = txtOption1.Text
End Sub
Private Sub Option2_Click()
txtCorrectAnswer.Text = txtOption2.Text
End Sub
Private Sub Option3_Click()
txtCorrectAnswer.Text = txtOption3.Text
End Sub
Private Sub Option4_Click()
txtCorrectAnswer.Text = txtOption4.Text
End Sub
Private Sub Option5_Click()
txtCorrectAnswer.Text = txtOption5.Text
End Sub

