VERSION 5.00
Begin VB.Form frmStudentViolations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Student Violations "
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Add1 
      Caption         =   "&Add 1 Point"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
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
      TabIndex        =   9
      Top             =   3720
      Width           =   6255
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Record Violation"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
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
      Height          =   1950
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   5535
   End
   Begin VB.TextBox txtFName 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtLName 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "First Name"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last Name"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   7440
      X2              =   120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   7440
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000009&
      X1              =   7440
      X2              =   4080
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line11 
      X1              =   7440
      X2              =   4080
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      X1              =   7440
      X2              =   4800
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line9 
      X1              =   7440
      X2              =   4800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   7440
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line7 
      X1              =   5760
      X2              =   7440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   5760
      X2              =   7440
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line5 
      X1              =   5760
      X2              =   7440
      Y1              =   1560
      Y2              =   1560
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
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   2520
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmStudentViolations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Add1_Click()
Dim AddOne  As Integer
If txtFireMode.Text = "" Then
   MsgBox ("No student has been selected! Please click one.")
   List_Student.SetFocus
   Exit Sub
End If

AddOne = 0
txtLName.Enabled = True
txtFName.Enabled = True
With rsStudent
    .MoveFirst
    .Find "username= '" & Trim(UCase(Left(txtFireMode.Text, 15))) & "'"
  If Not .EOF Then
    AddOne = Val(!studentviolation)
    AddOne = AddOne + 1
    !studentviolation = AddOne
    .Update
  End If
End With
  Call Populate_List
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub
Sub Populate_List()
List_Student.Clear
With rsStudent
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
   List_Student.AddItem !UserName & Space(15 - Len(!UserName)) & " | " & !studentlName & Space(20 - Len(!studentlName)) & " | " & !studentfname
   .MoveNext
Loop
End With
End Sub
Private Sub cmdAdd_Click()
Dim ReduceNumber  As Integer

If txtFireMode.Text = "" Then
   MsgBox ("No student has been selected! Please click one.")
   List_Student.SetFocus
   Exit Sub
End If


ReduceNumber = 0
txtLName.Enabled = True
txtFName.Enabled = True
With rsStudent
    .MoveFirst
    .Find "username= '" & Trim(UCase(Left(txtFireMode.Text, 15))) & "'"
  If Not .EOF Then
    ReduceNumber = Val(!studentviolation)
    ReduceNumber = ReduceNumber - 1
    !studentviolation = ReduceNumber
    .Update
  End If
End With
  Call Populate_List
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
txtLName.Text = Mid(txtFireMode.Text, 19, 20)

End Sub
Private Sub txtLName_Change()
With rsStudent
  .MoveFirst
  .Find "username= '" & Left(txtFireMode.Text, 15) & "'"
  If Not .EOF Then
      txtFName.Text = !studentfname
      txtLName.Refresh
         End If
End With
End Sub
