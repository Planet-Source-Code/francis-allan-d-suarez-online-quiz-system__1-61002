VERSION 5.00
Begin VB.Form frmCourse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5280
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cance&l"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton CmdAdd 
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
      TabIndex        =   11
      Top             =   4080
      Width           =   5175
   End
   Begin VB.ListBox ListCourse 
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
      ItemData        =   "frmCourse.frx":0000
      Left            =   120
      List            =   "frmCourse.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox txtCourseDescription 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtCourseId 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Course Description"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Course Code"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   6480
      X2              =   3240
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line7 
      X1              =   6480
      X2              =   3240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   6480
      X2              =   3240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line5 
      X1              =   6480
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   4680
      X2              =   6480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   4680
      X2              =   6480
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
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
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Course Description:"
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
      TabIndex        =   8
      Top             =   720
      Width           =   1695
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
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
Unload Me
End Sub
Private Sub cmdAdd_Click()
txtCourseId.Enabled = True
txtCourseDescription.Enabled = True
If cmdAdd.Caption = "&Add" Then
   cmdAdd.Caption = "&Save"
   cmdEdit.Enabled = False
   cmdDelete.Enabled = False
   txtCourseId.SetFocus
ElseIf cmdAdd.Caption = "&Save" Then
   cmdAdd.Caption = "&Add"
   With rsCourse
      .MoveFirst
      .Find "coursecode= '" & Trim(txtCourseId.Text) & "'"
      If Not .EOF Then
         MsgBox "The Course Code already exists!", vbCritical
      Else
         .AddNew
         !coursecode = UCase(txtCourseId.Text)
         !coursedescription = txtCourseDescription.Text
         .Update
      End If
    End With
Call Populate_List
txtCourseId.Text = ""
txtCourseDescription.Text = ""
cmdEdit.Enabled = True
cmdDelete.Enabled = True
End If
End Sub
Sub Populate_List()
ListCourse.Clear
With rsCourse
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
   ListCourse.AddItem !coursecode & Space(10 - Len(!coursecode)) & " | " & !coursedescription
   .MoveNext
Loop
End With
End Sub
Private Sub cmdDelete_Click()
With rsCourse
  .MoveFirst
  .Find "coursecode = '" & Left(txtFireMode.Text, 10) & "'"
  If Not .EOF Then
     If MsgBox("Do you really want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
        .Delete
        Call Populate_List
        txtFireMode.Text = ""
        cmdAdd.SetFocus
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
  txtCourseId.Enabled = True
  txtCourseDescription.Enabled = True
  txtCourseId.SetFocus
  cmdAdd.Enabled = False
  cmdEdit.Caption = "&Save"
With rsCourse
  .MoveFirst
  .Find "coursecode= '" & Left(txtFireMode.Text, 10) & "'"
  If Not .EOF Then
      txtCourseId.Text = !coursecode
      txtCourseDescription.Text = !coursedescription
    Else
      MsgBox "No Record to Edit", vbInformation
   End If
End With
Else
  With rsCourse
    !coursecode = txtCourseId.Text
    !coursedescription = txtCourseDescription.Text
    .Update
  End With
  cmdAdd.Enabled = True
  Call Populate_List
  cmdEdit.Caption = "&Edit"
  txtCourseId.Enabled = False
  txtCourseDescription.Enabled = False
End If
End Sub
Private Sub cmdView_Click()
ListCourse.ListIndex = -1
ListCourse.Clear
Call Populate_List
End Sub
Private Sub cmdCancel_Click()
With rsCourse
   .CancelUpdate
End With
txtCourseId.Text = ""
txtCourseDescription.Text = ""
txtCourseId.Enabled = False
txtCourseDescription.Enabled = False
cmdAdd.Enabled = True
cmdEdit.Enabled = True
cmdAdd.Caption = "&Add"
cmdEdit.Caption = "&Edit"
End Sub
Private Sub Form_Load()
strCnC = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1C = New ADODB.Connection
strCn1C.Open strCnC
Call Initialization_Course
Call Populate_List
ViewFlag = 1
End Sub
Private Sub Listcourse_Click()
txtFireMode.Text = ListCourse.Text
End Sub
Private Sub txtCourseDescription_LostFocus()
cmdAdd.SetFocus
End Sub

