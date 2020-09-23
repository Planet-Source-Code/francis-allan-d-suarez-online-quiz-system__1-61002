VERSION 5.00
Begin VB.Form frmStudent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEditUserName 
      Caption         =   "Edit &User Name"
      Height          =   495
      Left            =   3840
      TabIndex        =   14
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   2775
      Left            =   5280
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&l"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4560
      Width           =   5175
   End
   Begin VB.TextBox txtStudentFName 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox List_Student 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox txtStudentID 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtStudentLName 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Name"
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ID NUmber"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      X1              =   6480
      X2              =   3720
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line9 
      X1              =   6480
      X2              =   3720
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   6480
      X2              =   3720
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line7 
      X1              =   6240
      X2              =   6480
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   6240
      X2              =   6480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line5 
      X1              =   6240
      X2              =   6480
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   1560
      X2              =   120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   1560
      X2              =   120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   6480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label4 
      Caption         =   "First Name:"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name:"
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
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Student ID:"
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
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
Unload Me
End Sub
Private Sub cmdAdd_Click()
txtStudentID.Enabled = True
txtStudentLName.Enabled = True
txtStudentFName.Enabled = True

If cmdAdd.Caption = "&Add" Then
   cmdAdd.Caption = "&Save"
   cmdEdit.Enabled = False
   cmdDelete.Enabled = False
   txtStudentID.Text = ""
   txtStudentLName.Text = ""
   txtStudentFName.Text = ""
   txtStudentID.SetFocus
ElseIf cmdAdd.Caption = "&Save" Then
   cmdAdd.Caption = "&Add"
   With rsStudent
      .MoveFirst
      .Find "studentid= '" & Trim(txtStudentID.Text) & "'"
      If Not .EOF Then
         MsgBox "The Student ID already exists!", vbCritical
      Else
         .AddNew
         !studentid = txtStudentID.Text
         !studentlName = txtStudentLName.Text
         !studentfname = txtStudentFName.Text
         !UserName = UCase(Trim(txtStudentLName.Text) & Left(txtStudentFName.Text, 1))
         !Password = UCase("pass")
         !studentviolation = "10"
         .Update
        txtStudentID.Enabled = False
        txtStudentLName.Enabled = False
        txtStudentFName.Enabled = False
      End If
    End With
Call Populate_List
txtStudentID.Text = ""
txtStudentFName.Text = ""
txtStudentLName.Text = ""
cmdEdit.Enabled = True
cmdDelete.Enabled = True
End If
End Sub
Sub Populate_List()
List_Student.Clear
With rsStudent
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
   List_Student.AddItem !studentid & Space(13 - Len(!studentid)) & " | " & !UserName
   .MoveNext
Loop
End With
End Sub
Private Sub cmdDelete_Click()
With rsStudent
  .MoveFirst
  .Find "studentid= '" & Left(txtFireMode.Text, 12) & "'"
  If Not .EOF Then
     If MsgBox("Do you really want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
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
End Sub
Private Sub cmdEdit_Click()
If cmdEdit.Caption = "&Edit" Then
  txtStudentID.Enabled = True
  txtStudentLName.Enabled = True
  txtStudentFName.Enabled = True
  txtStudentID.SetFocus
  cmdAdd.Enabled = False
  cmdEdit.Caption = "&Save"
  cmdDelete.Enabled = False
With rsStudent
  .MoveFirst
  .Find "studentid= '" & Left(txtFireMode.Text, 12) & "'"
  If Not .EOF Then
      txtStudentID.Text = !studentid
      txtStudentLName.Text = !studentlName
      txtStudentFName.Text = !studentfname
    Else
      MsgBox "No Record to Edit", vbInformation
   End If
End With
Else
  With rsStudent
    !studentid = txtStudentID.Text
    !studentlName = txtStudentLName.Text
    !studentfname = txtStudentFName.Text
    .Update
  End With
  cmdAdd.Enabled = True
  cmdDelete.Enabled = True
  Call Populate_List
  cmdEdit.Caption = "&Edit"
  txtStudentID.Enabled = False
  txtStudentLName.Enabled = False
  txtStudentFName.Enabled = False
End If
End Sub
Private Sub cmdCancel_Click()
With rsStudent
   .CancelUpdate
End With
txtStudentID.Text = ""
txtStudentLName.Text = ""
txtStudentFName.Text = ""
txtStudentID.Enabled = False
txtStudentLName.Enabled = False
txtStudentFName.Enabled = False
cmdAdd.Enabled = True
cmdEdit.Enabled = True
cmdDelete.Enabled = True
cmdAdd.Caption = "&Add"
cmdEdit.Caption = "&Edit"
End Sub

Private Sub cmdEditUserName_Click()
If cmdEditUserName.Caption = "Edit &User Name" Then
  txtStudentID.Enabled = True
  txtStudentLName.Enabled = True
  txtStudentFName.Enabled = True
  txtStudentID.SetFocus
  cmdAdd.Enabled = False
  cmdEditUserName.Caption = "&Save"
  cmdDelete.Enabled = False
With rsStudent
  .MoveFirst
  .Find "studentid= '" & Left(txtFireMode.Text, 12) & "'"
  If Not .EOF Then
      txtStudentID.Text = !studentid
      txtStudentLName.Text = !studentlName
      txtStudentFName.Text = !studentfname
    Else
      MsgBox "No Record to Edit", vbInformation
   End If
End With
Else
  With rsStudent
    !studentid = txtStudentID.Text
    !studentlName = txtStudentLName.Text
    !studentfname = txtStudentFName.Text
    !UserName = UCase(txtStudentLName.Text & Left(txtStudentFName.Text, 1) & Day(Date))
    .Update
  End With
  cmdAdd.Enabled = True
  cmdDelete.Enabled = True
  Call Populate_List
  cmdEditUserName.Caption = "Edit &User Name"
  txtStudentID.Enabled = False
  txtStudentLName.Enabled = False
  txtStudentFName.Enabled = False
End If
End Sub

Private Sub Form_Load()
strCnS = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1S = New ADODB.Connection
strCn1S.Open strCnS
Call Initialization_Student
Call Populate_List
End Sub
Private Sub List_student_Click()
txtFireMode.Text = List_Student.Text
End Sub
Private Sub txtStudentFName_LostFocus()
cmdAdd.SetFocus
End Sub
