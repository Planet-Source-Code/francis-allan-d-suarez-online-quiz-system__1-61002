VERSION 5.00
Begin VB.Form frmAllFacultyMembers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View All Faculty Members"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3960
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox List_AllFacultyMembers 
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
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last Name"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "First Name"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Name"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ID Number"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   5640
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   5640
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   3600
      X2              =   5640
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line5 
      X1              =   3600
      X2              =   5640
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   2160
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2160
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   5640
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5640
      Y1              =   3840
      Y2              =   3840
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmAllFacultyMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
strCn = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1 = New ADODB.Connection
strCn1.Open strCn
Call Initialization
Call Populate_List
End Sub
Private Sub List_allfacultymembers_Click()
txtFireMode.Text = List_AllFacultyMembers.Text
End Sub
Sub Populate_List()
List_AllFacultyMembers.Clear
With rsFaculty
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
   List_AllFacultyMembers.AddItem !facultyid & Space(12 - Len(!facultyid)) & " | " & !UserName & Space(12 - Len(!UserName)) & " | " & !facultyfname & Space(12 - Len(!facultyfname)) & " | " & !facultylName & Space(15 - Len(!facultylName))
   .MoveNext
Loop
End With
End Sub
