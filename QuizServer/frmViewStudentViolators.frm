VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmViewStudentViolators 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Student Violators"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print All"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
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
      Caption         =   "Credits Left"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   4200
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line5 
      X1              =   3360
      X2              =   4200
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   3360
      X2              =   4200
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   840
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   840
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
Attribute VB_Name = "frmViewStudentViolators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
   List_Student.AddItem !UserName & Space(25 - Len(!UserName)) & " | " & !studentviolation
   .MoveNext
Loop
End With
End Sub

Private Sub cmdPrint_Click()
Dim FUllName As String
Dim StringSpace As String
Dim I As Integer
Dim a, b, c, d, e As String
Dim BeginPage, EndPage, NumCopies, j
Dim NumSpaces As Integer
Dim SpaceChar As String
Dim x As Integer
Dim OneChar As String
Dim RecordCtr As Integer
'Set Cancel to True
CommonDialog1.CancelError = True
'On Error GoTo errhandler
'Display the Prin dialog box
CommonDialog1.ShowPrinter
'Get user-selected values from the dialog box
BeginPage = CommonDialog1.FromPage
EndPage = CommonDialog1.ToPage
NumCopies = CommonDialog1.Copies

strCnViolation = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Violation = New ADODB.Connection
strCn1Violation.Open strCnViolation
Call Initialization_QueryViolation


'With rsLimit
'   .MoveFirst
'   .Find "compositenumquestion = '" & Trim(Combo1.Text) & "'"
'   If Not .EOF Then
'       If Len(!completed) = 5 Then
'            Varlimit = Right(!completed, 2)
'       End If
'       If Len(!completed) = 3 Then
'            Varlimit = Right(!completed, 1)
'       End If
'   End If
'End With

For j = 1 To NumCopies
I = 1
Printer.FontSize = 13
'Printer.Print " "
Printer.Print "STI Operation: Spoken English Environment Program"
Printer.FontSize = 12
Printer.Print "     "
Printer.Print Tab(10); "Student Name:   "; "                    Credits Left";
Printer.Print "     "
Printer.FontName = "courier new"
With rsViolation
   .MoveFirst
End With
With rsViolation
  If Not .EOF Then
     .MoveFirst
  End If
RecordCtr = 0
Do Until .EOF
    RecordCtr = RecordCtr + 1
    FUllName = UCase(Left(!studentlName, 1)) & Right(!studentlName, Len(!studentlName) - 1) & ", " _
               & UCase(Left(!studentfname, 1)) & Right(!studentfname, Len(!studentfname) - 1)
    StringSpace = Space(40 - Len(FUllName))
    If Trim(!studentviolation) < 10 Then
       SpaceChar = " "
      Else
       SpaceChar = ""
    End If
    Printer.Print Tab(10); " "; FUllName; StringSpace; SpaceChar; !studentviolation;
   .MoveNext
   If RecordCtr >= 55 Then
      RecordCtr = 0
      Printer.NewPage
      
    End If
Loop
End With
Printer.Print "     "
Printer.FontSize = 8
Printer.Print "Note: If there are problems, please do not hesitate to approach the undersigned. Thank you."
Printer.Print "     "
Printer.Print "     "
Printer.FontSize = 12
Printer.Print "     "
Printer.Print Tab(25); "                     ________________________"; Date;
Printer.Print Tab(25); "                        Academic Supervisor ";

Next j


Printer.EndDoc
 Exit Sub
errhandler:
  Exit Sub
End Sub


Private Sub Form_Load()
strCnS = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1S = New ADODB.Connection
strCn1S.Open strCnS
Call Initialization_Student
Call Populate_List
End Sub


