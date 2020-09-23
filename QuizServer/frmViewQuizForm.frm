VERSION 5.00
Begin VB.Form frmViewQuizForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Quiz Items"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   4680
      Width           =   1695
   End
   Begin VB.ListBox List_ViewQuizForm 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   240
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      X1              =   4800
      X2              =   7440
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line6 
      X1              =   4800
      X2              =   7440
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   7440
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   2880
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   2880
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   7440
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7440
      Y1              =   4560
      Y2              =   4560
   End
End
Attribute VB_Name = "frmViewQuizForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
Unload Me
End Sub
Private Sub Form_Load()
Call Populate_List
End Sub
Sub Populate_List()
Dim ValueString As String
ValueString = Trim(frmQuizForm.txtUserName.Text & "|" & frmQuizForm.Combo1.Text)
List_ViewQuizForm.Clear
With rsQuestion
  .MoveFirst
  .Find "compositenumquestion= '" & ValueString & "'"
  If Not .EOF Then
     .MoveFirst
  End If
  Do Until .EOF
   List_ViewQuizForm.AddItem !compositenumquestion & Space(20 - Len(!compositenumquestion)) & " | " & _
   !itemnumber & Space(3 - Len(!itemnumber)) & " | " & _
   !question & Space(250 - Len(!question)) & " | " & _
   !Option1 & Space(50 - Len(!Option1)) & " | " & _
   !Option2 & Space(50 - Len(!Option2)) & " | " & _
   !Option3 & Space(50 - Len(!Option3)) & " | " & _
   !Option4 & Space(50 - Len(!Option4)) & " | " & _
   !Option5 & Space(50 - Len(!Option5)) & " | " & _
   !correctanswer & Space(50 - Len(!correctanswer))
   .MoveNext
Loop
End With
End Sub

