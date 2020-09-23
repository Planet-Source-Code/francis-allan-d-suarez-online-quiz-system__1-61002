VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmViewQuestions 
   Caption         =   "View Questions"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   4560
      Width           =   1455
   End
   Begin VB.ListBox List_ViewQuizForm 
      Height          =   4155
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   5040
      X2              =   7320
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line6 
      X1              =   5040
      X2              =   7320
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   2160
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   2160
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   7440
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   7320
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7320
      Y1              =   4440
      Y2              =   4440
   End
End
Attribute VB_Name = "frmViewQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
Dim I As Integer
Dim a, b, c, d, e As String
Dim BeginPage, EndPage, NumCopies, j
Dim NumSpaces As Integer
Dim SpaceChar As String
Dim x As Integer
Dim OneChar As String
'Set Cancel to True
CommonDialog1.CancelError = True
On Error GoTo errhandler
'Display the Prin dialog box
CommonDialog1.ShowPrinter
'Get user-selected values from the dialog box
BeginPage = CommonDialog1.FromPage
EndPage = CommonDialog1.ToPage
NumCopies = CommonDialog1.Copies



With rsLimit
   .MoveFirst
   .Find "compositenumquestion = '" & Trim(Left(FrmCopyQuizForm.txtFireMode.Text, 20)) & "'"
   If Not .EOF Then
       If Len(!completed) = 5 Then
            Varlimit = Right(!completed, 2)
       End If
       If Len(!completed) = 3 Then
            Varlimit = Right(!completed, 1)
       End If
   End If
End With

For j = 1 To NumCopies
I = 1
Printer.FontSize = 20
Printer.Print " "
Printer.Print " "
Printer.Print Tab(15); "STI QUIZ QUESTIONS"
Printer.FontSize = 12
Printer.Print "     "
'Printer.Print "     "
Printer.Print Tab(10); "Quiz Form: "; Left(FrmCopyQuizForm.txtFireMode.Text, 20); "          Total No. of Items: "; Varlimit;
Printer.Print "     "
Printer.Print "     "
Printer.Print Tab(1); "Questions ";
Printer.Print "     "
'Printer.FontName = "courier new"
With rsQuestion
   .MoveFirst
End With
With rsQuestion
  If Not .EOF Then
     .MoveFirst
  End If
For I = 1 To Varlimit
    .Find "compositenumquestion = '" & Trim(Left(FrmCopyQuizForm.txtFireMode.Text, 20)) & "'"
    If Not .EOF Then
    If Len(!question) < 60 Then
      Printer.Print "* " & !question & " - " & !correctanswer;
    End If
    If Len(!question) >= 60 And Len(!question) <= 120 Then
      Printer.Print "* " & Left(!question, 60);
      Printer.Print "      "
      Printer.Print "  " & Right(!question, Len(!question) - 60) & " - " & !correctanswer;
    End If
    If Len(!question) > 120 And Len(!question) <= 180 Then
      Printer.Print "* " & Left(!question, 60);
      VarQuestion = Right(!question, Len(!question) - 60)
      Printer.Print "      "
      Printer.Print "  " & Left(VarQuestion, 60);
      VarQuestion = Right(VarQuestion, Len(VarQuestion) - 60)
      Printer.Print "      "
      Printer.Print "  " & VarQuestion & " - " & !correctanswer;
    End If
    
    Printer.Print "      "
    End If
    .MoveNext
Next I
End With
Printer.Print "     "
Printer.Print "     "
Printer.Print "     "
Printer.Print "     "
Printer.Print "     "
'Printer.Print Tab(25); "                     _______________"; Date;
'Printer.Print Tab(25); "                        Instructor ";

Next j


Printer.EndDoc
 Exit Sub
errhandler:
  Exit Sub
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Populate_List
End Sub
Sub Populate_List()
Dim ValueString As String
Dim ReplaceNumber As String
ValueString = Trim(Left(FrmCopyQuizForm.txtFireMode.Text, 20))
List_ViewQuizForm.Clear
With rsQuestion
  .MoveFirst
  .Find "compositenumquestion= '" & ValueString & "'"
  If Not .EOF Then
     .MoveFirst
  End If
  Do Until .EOF
   If !compositenumquestion = ValueString Then
   ReplaceNumber = !itemnumber
   If Len(ReplaceNumber) = 1 Then
      ReplaceNumber = "0" & ReplaceNumber
   End If
   List_ViewQuizForm.AddItem ReplaceNumber & Space(3 - Len(ReplaceNumber)) & " | " & _
   !question & Space(250 - Len(!question))
  
   End If
   .MoveNext
Loop
End With
End Sub



