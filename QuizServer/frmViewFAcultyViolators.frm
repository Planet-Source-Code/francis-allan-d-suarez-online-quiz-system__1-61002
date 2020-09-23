VERSION 5.00
Begin VB.Form frmViewFAcultyViolators 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Faculty Violators"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClose 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
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
      Height          =   2790
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   4200
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line6 
      X1              =   1560
      X2              =   2760
      Y1              =   1680
      Y2              =   2160
   End
   Begin VB.Line Line5 
      X1              =   2760
      X2              =   4200
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   2760
      X2              =   4200
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   1440
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   1440
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4200
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "frmViewFAcultyViolators"
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
   List_Faculty.AddItem !UserName & Space(25 - Len(!UserName)) & " | " & !facultyviolations
   .MoveNext
Loop
End With
End Sub
Private Sub Form_Load()
strCn = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1 = New ADODB.Connection
strCn1.Open strCn
Call Initialization
Call Populate_List
End Sub

