VERSION 5.00
Begin VB.Form frmFacultyViolations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Faculty Violations"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMerit 
      Caption         =   "&Merit"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Demerit"
      Height          =   495
      Left            =   4320
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
      Width           =   4815
   End
   Begin VB.ListBox List_Faculty 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox txtFname 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtLname 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "First Name"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last Name"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000005&
      X1              =   6000
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line9 
      X1              =   6000
      X2              =   120
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   6000
      X2              =   3840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line7 
      X1              =   4920
      X2              =   6000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   4920
      X2              =   6000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line5 
      X1              =   4920
      X2              =   6000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   4320
      X2              =   6000
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line3 
      X1              =   4320
      X2              =   6000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   4320
      X2              =   6000
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   6000
      Y1              =   1680
      Y2              =   1680
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
      Width           =   1575
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
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Width           =   1335
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
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmFacultyViolations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClose_Click()
Unload Me
End Sub
Sub Populate_List()
List_Faculty.Clear
With rsFaculty
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
   List_Faculty.AddItem !facultylName & Space(20 - Len(!facultylName)) & " | " & !facultyfname
   .MoveNext
Loop
End With
End Sub


Private Sub cmdAdd_Click()
Dim ReduceNumber  As Integer
If txtFireMode.Text = "" Then
   MsgBox ("No faculty member has been selected! Please click one.")
   List_Faculty.SetFocus
   Exit Sub
End If
reducenumer = 0
txtLname.Enabled = True
txtFname.Enabled = True
With rsFaculty
    .MoveFirst
    .Find "username= '" & Trim(UCase(txtLname.Text)) & UCase(Left(txtFname.Text, 1)) & "'"
  If Not .EOF Then
    ReduceNumber = Val(!facultyviolations)
    ReduceNumber = ReduceNumber - 1
    !facultyviolations = ReduceNumber
    .Update
  End If
End With
  Call Populate_List
End Sub

Private Sub cmdMerit_Click()
Dim IncreaseNumber  As Integer
If txtFireMode.Text = "" Then
   MsgBox ("No faculty member has been selected! Please click one.")
   List_Faculty.SetFocus
   Exit Sub
End If

ReduceNumber = 0
txtLname.Enabled = True
txtFname.Enabled = True
With rsFaculty
    .MoveFirst
    .Find "username= '" & Trim(UCase(txtLname.Text)) & UCase(Left(txtFname.Text, 1)) & "'"
  If Not .EOF Then
    IncreaseNumber = Val(!facultyviolations)
    IncreaseNumber = IncreaseNumber + 1
    !facultyviolations = IncreaseNumber
    .Update
  End If
End With
  Call Populate_List
End Sub

Private Sub Form_Load()
strCn = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1 = New ADODB.Connection
strCn1.Open strCn
Call Initialization
Call Populate_List
End Sub
Private Sub List_faculty_Click()
txtFireMode.Text = List_Faculty.Text
txtLname.Text = Left(txtFireMode.Text, 20)
End Sub
Private Sub txtLName_Change()
With rsFaculty
  .MoveFirst
  .Find "facultylname= '" & txtLname.Text & "'"
  If Not .EOF Then
      txtFname.Text = !facultyfname
      txtLname.Refresh
         End If
End With
End Sub

