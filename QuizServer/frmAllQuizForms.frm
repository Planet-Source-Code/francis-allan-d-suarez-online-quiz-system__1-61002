VERSION 5.00
Begin VB.Form frmAllQuizForms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View All Quiz Forms"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTime 
      Caption         =   "Change &Time Limit"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAllow 
      Caption         =   "&Allow To Take Quiz Today"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisallow 
      Caption         =   "Disallow From Taking The &Quiz"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdChangeExpiry 
      Caption         =   "Change E&xpiry Date"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
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
      TabIndex        =   8
      Top             =   4320
      Width           =   5655
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
      Height          =   2790
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   6855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Timer"
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No. of Items"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Expiry Date"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quiz Form"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   2895
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   6840
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   6840
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6840
      Y1              =   4200
      Y2              =   4200
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
      Top             =   4320
      Width           =   1455
   End
End
Attribute VB_Name = "frmAllQuizForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VarFireMode As String
Dim LastItem As Integer
Private Sub btnClose_Click()
Unload Me
End Sub
Sub Populate_List()
List_AllQuizForms.Clear
With rsLimit
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
   List_AllQuizForms.AddItem !compositenumquestion & Space(25 - Len(!compositenumquestion)) & " | " & !dateend & Space(15 - Len(!dateend)) & " | " & !completed & Space(5 - Len(!completed)) & " | " & !ftimers & " secs."
   .MoveNext
Loop
End With
End Sub

Private Sub cmdAllow_Click()
'MsgBox (Left(txtFireMode.Text, 20))
VarFireMode = Left(txtFireMode.Text, 20)
SearchChar = "|"
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
         RemainingString = Left(txtFireMode.Text, 20)
         Mypos = InStr(1, RemainingString, SearchChar)
         If UCase(Trim(Form1.txtUserName.Text)) <> Trim(Left(RemainingString, Mypos - 1)) Then
            MsgBox ("You cannot manipulate a quiz form owned by your co-teacher!")
            Exit Sub
         End If
        !dateend = Date
        .Update
        Call Populate_List
        txtFireMode.Text = ""
              
  Else
    MsgBox "No record to modify.", vbInformation
   End If
End With
End Sub

Private Sub cmdChangeExpiry_Click()
'MsgBox (Left(txtFireMode.Text, 20))
VarFireMode = Left(txtFireMode.Text, 20)
SearchChar = "|"
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
     RemainingString = Left(txtFireMode.Text, 20)
     Mypos = InStr(1, RemainingString, SearchChar)
     If UCase(Trim(Form1.txtUserName.Text)) <> Trim(Left(RemainingString, Mypos - 1)) Then
         MsgBox ("You cannot manipulate a quiz form owned by your co-teacher!")
         Exit Sub
     End If
     If MsgBox("You can now change the expiry date of the quizform. Continue?", vbYesNo + vbQuestion) = vbYes Then
        !dateend = InputBox("Enter the expiry date in mm/dd/yy format. Thank you.")
        If !dateend = "" Then
            !dateend = Date
        End If
        .Update
        Call Populate_List
        txtFireMode.Text = ""
       Else
         Exit Sub
     End If
  Else
    MsgBox "No record to modify.", vbInformation
   End If
End With
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdDelete_Click()
'MsgBox (Left(txtFireMode.Text, 20))
VarFireMode = Left(txtFireMode.Text, 20)
SearchChar = "|"
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
     RemainingString = Left(txtFireMode.Text, 20)
     Mypos = InStr(1, RemainingString, SearchChar)
     If UCase(Trim(Form1.txtUserName.Text)) <> Trim(Left(RemainingString, Mypos - 1)) Then
         MsgBox ("You cannot delete a quiz form owned by your co-teacher!")
         Exit Sub
     End If
     If MsgBox("Do you really want to delete this record? Please be informed that once you have deleted this quiz form, " & _
         "the record of the students who have taken this quiz will also be deleted." _
        , vbYesNo + vbQuestion) = vbYes Then
        .Delete
        Call Populate_List
        txtFireMode.Text = ""
       Else
         Exit Sub
     End If
  Else
    MsgBox "No record to delete.", vbInformation
   End If
End With

With rsQuestion
  For I = 1 To LastItem
  .MoveFirst
  .Find "compositenumquestion= '" & Trim(VarFireMode) & "'"
  If Not .EOF Then
        .Delete
      Else
         Exit Sub
  End If
  Next I
End With

With rsQuiz
  If Not .EOF Then
   .MoveFirst
  End If
  Do While Not .EOF
     .Find "quizcompositekey= '" & VarFireMode & "'"
     If Not .EOF Then
          .Delete
          .MoveNext
     Else
         Exit Sub
     End If
  Loop
End With

End Sub

Private Sub cmdDisallow_Click()
VarFireMode = Left(txtFireMode.Text, 20)
SearchChar = "|"
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
     RemainingString = Left(txtFireMode.Text, 20)
     Mypos = InStr(1, RemainingString, SearchChar)
     If UCase(Trim(Form1.txtUserName.Text)) <> Trim(Left(RemainingString, Mypos - 1)) Then
         MsgBox ("You cannot manipulate a quiz form owned by your co-teacher!")
         Exit Sub
     End If
        !dateend = Date - 1
        .Update
        Call Populate_List
        txtFireMode.Text = ""
  Else
    MsgBox "No record to modify.", vbInformation
   End If
End With
End Sub

Private Sub CmdTime_Click()
Dim VarTimeLimit As String
VarFireMode = Left(txtFireMode.Text, 20)
SearchChar = "|"
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
         RemainingString = Left(txtFireMode.Text, 20)
         Mypos = InStr(1, RemainingString, SearchChar)
         If UCase(Trim(Form1.txtUserName.Text)) <> Trim(Left(RemainingString, Mypos - 1)) Then
            MsgBox ("You cannot manipulate a quiz form owned by your co-teacher!")
            Exit Sub
         End If
         VarTimeLimit = InputBox("Enter figure in seconds (ex. 35,40,50).")
         If Not IsNumeric(VarTimeLimit) Then
            MsgBox "The value should be a number!", vbExclamation
            Exit Sub
         End If
        !ftimers = VarTimeLimit
        .Update
        Call Populate_List
        txtFireMode.Text = ""
              
  Else
    MsgBox "No record to modify.", vbInformation
   End If
End With
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


strCnQ = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Q = New ADODB.Connection
strCn1Q.Open strCnQ
Call Initialization_Quiz
End Sub
Private Sub List_allquizforms_Click()
txtFireMode.Text = List_AllQuizForms.Text
End Sub


