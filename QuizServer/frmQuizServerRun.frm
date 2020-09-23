VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmQuizServerRun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quiz Server"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSendMessage 
      Caption         =   "&Send Message To The Clients"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect All Clients"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   11
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3480
      Top             =   0
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000080&
      ForeColor       =   &H80000005&
      Height          =   4155
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      TabIndex        =   3
      Top             =   5400
      Width           =   3135
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000080&
      ForeColor       =   &H80000005&
      Height          =   4350
      ItemData        =   "frmQuizServerRun.frx":0000
      Left            =   0
      List            =   "frmQuizServerRun.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   600
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Computer/s connected are each represented with their IP number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Instructions Received:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblConnections 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Connections:"
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
      Left            =   0
      TabIndex        =   7
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "IP Address:"
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
      Left            =   0
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Host:"
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
      Left            =   0
      TabIndex        =   4
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblAddress 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label lblHostID 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
End
Attribute VB_Name = "frmQuizServerRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sServerMsg As String
Dim iSockets As Integer
Dim sREquestID As String
Dim IPctr As Integer
Dim Cnn1, rsStude, strCnn, rstAcquisition
Dim ClientArray(1 To 100) As String
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDisconnect_Click()
Dim DClientIndex As Integer
On Error Resume Next
For DClientIndex = 1 To 100
   Socket(DClientIndex).SendData "DisconnectClient"
Next DClientIndex
End Sub

Private Sub cmdSendMessage_Click()
frmSendMessage.Show 1
End Sub

Private Sub Form_Load()
lblConnections.Caption = 0
lblHostID.Caption = ""
lblAddress.Caption = ""

lblHostID.Caption = Socket(0).LocalHostName
lblAddress.Caption = Socket(0).LocalIP
Socket(0).LocalPort = 1007
sServerMsg = "Listening to port: " & Socket(0).LocalPort
List1.AddItem (sServerMsg)
Socket(0).Listen

strCnQu = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Qu = New ADODB.Connection
strCn1Qu.Open strCnQu
Call Initialization_Question

strCnQ = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Q = New ADODB.Connection
strCn1Q.Open strCnQ
Call Initialization_Quiz

strCnHistory = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1History = New ADODB.Connection
strCn1History.Open strCnHistory
Call Initialization_History

strCnS = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1S = New ADODB.Connection
strCn1S.Open strCnS
Call Initialization_Student

strCnLimit = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1Limit = New ADODB.Connection
strCn1Limit.Open strCnLimit
Call Initialization_Limit
End Sub

Private Sub frmDisconnect_Click()
Dim DClientIndex As Integer
On Error Resume Next
For DClientIndex = 1 To 100
   Socket(DClientIndex).SendData "DisconnectClient"
Next DClientIndex
End Sub

Private Sub lblConnections_Change()
If lblConnections.Caption > 0 Then
   cmdSendMessage.Enabled = True
   cmdDisconnect.Enabled = True
  Else
    cmdSendMessage.Enabled = False
    cmdDisconnect.Enabled = False
End If
End Sub

Private Sub Socket_Close(Index As Integer)
Dim I As Integer
Dim x As Integer
Dim ClientIndex As Integer

sServerMsg = "Connection closed: " & Socket(Index).RemoteHostIP
List1.AddItem (sServerMsg)
For ClientIndex = 1 To 100
    If ClientArray(ClientIndex) = Socket(Index).RemoteHostIP Then
       ClientArray(ClientIndex) = "~" & ClientArray(ClientIndex)
       Socket(Index).Close
    End If
Next ClientIndex
List2.Clear
For ClientIndex = 1 To 100
   If Left(ClientArray(ClientIndex), 1) <> "~" And ClientArray(ClientIndex) <> "" Then
       List2.AddItem ClientArray(ClientIndex)
   End If
Next ClientIndex
'For I = 0 To List2.ListCount - 1
'  If Trim(List2.List(I)) = Trim(Socket(Index).RemoteHostIP) Then
'    List2.RemoveItem (I)
  '  If Socket(Index).State = sckConnected Then
'        Socket(Index).Close
'        Unload Socket(Index)
'        iSockets = iSockets - 1
'    End If
'  End If
'Next I
'Socket(Index).Close
'Unload Socket(Index)
'Unload Socket(Index)
'iSockets = iSockets - 1
'lblConnections.Caption = iSockets
End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim ClientIndex As Integer
Dim Flag As Integer
'On Error Resume Next
sServerMsg = "Connection request id " & requestID & " from " & Socket(Index).RemoteHostIP
If Index = 0 Then
  List1.AddItem (sServerMsg)
  sREquestID = requestID
  iSockets = iSockets + 1
  'lblConnections.Caption = iSockets
  Load Socket(iSockets)
  Socket(iSockets).LocalPort = 1007
  Socket(iSockets).Accept requestID
  IPctr = Index
  
End If
Flag = 0
For ClientIndex = 1 To 100
      If (Left(ClientArray(ClientIndex), 1) <> "~") And ClientArray(ClientIndex) = "" And Flag = 0 Then
          ClientArray(ClientIndex) = Socket(Index).RemoteHostIP
          Flag = 1
      End If
Next ClientIndex

List2.Clear

For ClientIndex = 1 To 100
    If Left(ClientArray(ClientIndex), 1) <> "~" And ClientArray(ClientIndex) <> "" Then
        List2.AddItem ClientArray(ClientIndex) & " --Index: " & ClientIndex
    End If
Next ClientIndex

If List1.ListCount >= 25 Then
  
  List1.Clear
End If
'List2.AddItem Socket(Index).RemoteHostIP
 
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim sItemData As String
Dim strData As String
Dim strOutData As String
Dim strConnect As String
Dim variableItem(100) As String
Dim x As Integer
Dim RemainingString As String
Dim VarLogin As String
Dim VarQuizForms As String
Dim Mypos As Integer
Dim VarIndex As Integer
Dim LastItem As Integer
Dim Item As Integer
Dim QuestionsArray() As String
Dim Questions As String
Dim VarQuestionSequenceLength As Integer
Dim varDataQuizCompositeKey As String
Dim VarQuizScore As String
Dim varAlreadyTakenFlag As Integer
Dim varFirstVisitFlag As Integer
Dim varUnfinishedSequence As Integer
Dim varRoute As String
Dim varPutzero As String
Dim varFindCompositeKey As String
Dim varFindUserName As String
Dim Flagger As Integer
Dim I As Integer
Dim DClientIndex As Integer
Dim VarTimer As String
On Error Resume Next
 varAlreadyTakenFlag = 0
 Item = 1


'get data from client
Socket(Index).GetData sItemData, vbString
sServerMsg = "Instruction received: " & " from " & Socket(Index).RemoteHostIP & "(" & sREquestID & ")"
List1.AddItem (sServerMsg)

'List1.ListIndex = Val(List1.ListCount)
VarIndex = Index
For x = 0 To 100
  If x = VarIndex Then
      variableItem(x) = sItemData
      
      
      
' ******************************Quizforms
      If variableItem(x) = "Quizforms" Then
          With rsLimit
             .MoveFirst
             Do While Not .EOF
                VarQuizForms = VarQuizForms & "~" & !compositenumquestion
                .MoveNext
             Loop
          End With
         Socket(VarIndex).SendData "Quizforms" & VarQuizForms & "~"
         List1.AddItem ("Request for Quizforms Response")
      End If
' /******************************Quizforms
  If Left(variableItem(x), 13) = "GroupMessage~" Then
   ' frmSendMessage.Show 1
    frmSendMessage.List1.AddItem Socket(VarIndex).RemoteHostIP & "> " & Right(variableItem(x), Len(variableItem(x)) - 13)
    List1.AddItem ("Chatting...")
    For DClientIndex = 1 To 100
       frmQuizServerRun.Socket(DClientIndex).SendData "GroupMessage~" & "> " & Right(variableItem(x), Len(variableItem(x)) - 13)
    Next DClientIndex
       
   End If

' ******************************QuizAnswer
If Left(variableItem(x), 10) = "QuizAnswer" Then
         RemainingString = Right(variableItem(x), Len(variableItem(x)) - 10)
         Mypos = InStr(1, RemainingString, "~", 1)
         With rsQuiz
            .MoveFirst
         End With
       With rsQuiz
            Flagger = 0
            If Not .EOF Then
                 .MoveFirst
                 .Find "quizcompositekey= '" & Left(RemainingString, Mypos - 1) & "'"
                 varFindCompositeKey = Left(RemainingString, Mypos - 1)
                 varDataQuizCompositeKey = Left(RemainingString, Mypos - 1)
                 RemainingString = Right(RemainingString, Len(RemainingString) - Mypos)
                 Mypos = InStr(1, RemainingString, "~", 1)
                 varFindUserName = Left(RemainingString, Mypos - 1)
                 Do While Not .EOF And Flagger <> 1
                      If Not .EOF Then
 ' +++++++++
 ' +++++++++
                          If !studentusername = varFindUserName Then
                              Flagger = 1
                              RemainingString = Right(RemainingString, Len(RemainingString) - Mypos)
                              Mypos = InStr(1, RemainingString, "~", 1)
                              If IsNull(!quizscore) Then
                                  !quizscore = "0"
                              End If
                              !quizscore = Val(!quizscore) + Left(RemainingString, Mypos - 1)
                              VarQuizScore = !quizscore
                              RemainingString = Right(RemainingString, Len(RemainingString) - Mypos)
                              Mypos = InStr(1, RemainingString, "~", 1)
                              !questionsequence = Left(RemainingString, Mypos - 1)
                              !route = Right(RemainingString, 1)
                              VarQuestionSequenceLength = Len(!questionsequence)
                              .Update
                              With rsLimit
                                .MoveFirst
                                .Find "compositenumquestion= '" & varDataQuizCompositeKey & "'"
                                If Not .EOF Then
                                   If Len(!completed) = 5 Then
                                       If VarQuestionSequenceLength >= (Right(!completed, 2) * 2) Then
                                           Socket(VarIndex).SendData "Finished" & VarQuizScore
                                           List1.AddItem ("Quiz Finished Response")
                                       End If
                                    End If
                                    If Len(!completed) = 3 Then
                                        If VarQuestionSequenceLength >= (Right(!completed, 1) * 2) Then
                                           Socket(VarIndex).SendData "Finished" & VarQuizScore
                                           List1.AddItem ("Quiz Finished Response")
                                        End If
                                    End If   '====================
                                 End If
                               End With
                           End If
' ++++++++++++
' ++++++++++++
                           If !studentusername <> varFindUserName Then
                               .MoveNext
                               .Find "quizcompositekey= '" & varFindCompositeKey & "'"
                               Flagger = 0
                           End If
                     End If
                Loop
                 If .EOF And Flagger <> 1 Then
                      .MoveLast
                      .AddNew
                       If IsNull(!quizscore) Then
                           !quizscore = "0"
                       End If
                      !quizcompositekey = varDataQuizCompositeKey
                      RemainingString = Right(RemainingString, Len(RemainingString) - Mypos)
                      Mypos = InStr(1, RemainingString, "~", 1)
                      !studentusername = varFindUserName
                      !quizscore = !quizscore + Left(RemainingString, Mypos - 1)
                      RemainingString = Right(RemainingString, Len(RemainingString) - Mypos)
                      Mypos = InStr(1, RemainingString, "~", 1)
                      !questionsequence = !questionsequence & Left(RemainingString, Mypos - 1)
                     !route = Right(RemainingString, 1)
                      .Update
                    End If
              End If
      End With
   End If

' /******************************QuizAnswer

' ******************************Login
      If Left(variableItem(x), 5) = "Login" Then
          RemainingString = Mid(variableItem(x), 6, Len(variableItem(x)) - 5)
          Mypos = InStr(1, RemainingString, "~", 1)
          
          With rsStudent
             .MoveFirst
             .Find "username= '" & Left(RemainingString, Mypos - 1) & "'"
             If Not .EOF Then
                If !Password Like Right(RemainingString, Len(RemainingString) - Mypos) Then
                       GlobalStudentusername = Left(RemainingString, Mypos - 1)
                       VarLogin = "Go"
                       If !firstvisit = 0 Then
                         varFirstVisitFlag = 0
                        Else
                         varFirstVisitFlag = 1
                       End If
                    Else
                       VarLogin = "Stop"
                End If
             End If
             
          End With
         Socket(VarIndex).SendData "Login" & VarLogin & varFirstVisitFlag
         List1.AddItem ("Logging In")
      End If
' /******************************Login

' ******************************EnglishCredits
      If Left(variableItem(x), 14) = "EnglishCredits" Then
          RemainingString = Right(variableItem(x), Len(variableItem(x)) - 14)
          With rsStudent
             .MoveFirst
             .Find "username= '" & Trim(RemainingString) & "'"
             
             If Not .EOF Then
                   Socket(VarIndex).SendData "EnglishCredits" & !studentviolation
                   List1.AddItem ("English Credits Inquiry")
              End If
           End With
         
      End If

' ******************************Login1
      If Left(variableItem(x), 6) = "CLogin" Then
          RemainingString = Right(variableItem(x), Len(variableItem(x)) - 6)
          Mypos = InStr(1, RemainingString, "~", 1)
          
          With rsStudent
             .MoveFirst
             .Find "username= '" & Left(RemainingString, Mypos - 1) & "'"
             RemainingString = Right(RemainingString, Len(RemainingString) - Mypos)
             Mypos = InStr(1, RemainingString, "~", 1)
             If Not .EOF Then
                If !Password Like Left(RemainingString, Mypos - 1) Then
                        !Password = Right(RemainingString, Len(RemainingString) - Mypos)
                        List1.AddItem ("Change Password Response")
                       .Update
                End If
              End If
           End With
         
      End If
' /******************************Login1
      
' ******************************Change Password
      If Left(variableItem(x), 14) = "ChangePassword" Then
          RemainingString = Right(variableItem(x), Len(variableItem(x)) - 14)
          Mypos = InStr(1, RemainingString, "~", 1)
          
          With rsStudent
             .MoveFirst
             .Find "username= '" & Left(RemainingString, Mypos - 1) & "'"
             If Not .EOF Then
                 !Password = Right(RemainingString, Len(RemainingString) - Mypos)
                 !firstvisit = 1
                 List1.AddItem ("Change Password Response")
                 .Update
             End If
          End With
      End If
' /******************************Change Password
      
' ******************************GoQuiz
      If Left(variableItem(x), 6) = "GoQuiz" Then
         RemainingString = Mid(variableItem(x), 7, Len(variableItem(x)) - 5)
         GlobalQuizform = RemainingString
          
         
          With rsLimit
              .MoveFirst
              .Find "compositenumquestion= '" & GlobalQuizform & "'"
              If Not .EOF Then
                If Len(!completed) = 5 Then
                    LastItem = Right(!completed, 2)
                End If
                If Len(!completed) = 3 Then
                    LastItem = Right(!completed, 1)
                End If
                VarTimer = !ftimers

                If DateDiff("d", Date, !dateend) < 0 Then
                   Socket(VarIndex).SendData "Expired"
                   List1.AddItem ("Quiz Form Expired Response")
                   Exit Sub
                End If
              End If
           End With
          With rsQuiz
             .MoveFirst
          End With
          With rsQuiz
            If Not .EOF Then
               .MoveFirst
               .Find "studentusername= '" & GlobalStudentusername & "'"
            End If
            Do While Not .EOF
               If Not .EOF Then
                   If GlobalQuizform = !quizcompositekey Then
                         If Len(!questionsequence) < (LastItem * 2) And Len(!questionsequence) <> 0 Then
                             varUnfinishedSequence = Len(!questionsequence) / 2
                             varRoute = !route
                             varAlreadyTakenFlag = 1
                             If Len(varUnfinishedSequence) < 2 Then
                                varUnfinishedSequence = "0" & varUnfinishedSequence
                             End If
                         End If
                         If Len(!questionsequence) >= (LastItem * 2) Then
                             varAlreadyTakenFlag = 2
                         End If
                   End If
               End If
               .MoveNext
               .Find "studentusername= '" & GlobalStudentusername & "'"
            Loop
          End With

'************** if the quizform is already completed
' **************** varAlreadyTakenFlag = 2
        If varAlreadyTakenFlag = 2 Then
            Socket(VarIndex).SendData "CompletedAlready"
            List1.AddItem ("Completed Already Response")
        End If
' /**************** varAlreadyTakenFlag = 2

'************** if the quizform is not yet taken
' **************** varAlreadyTakenFlag = 0
        If varAlreadyTakenFlag = 0 Then
           With rsLimit
              .MoveFirst
              .Find "compositenumquestion= '" & RemainingString & "'"
              If Not .EOF Then
                If Len(!completed) = 5 Then
                    LastItem = Right(!completed, 2)
                End If
                If Len(!completed) = 3 Then
                    LastItem = Right(!completed, 1)
                End If
              End If
           End With
         
           With rsQuestion
             .MoveFirst
             Do While Item <= LastItem
                If Not .EOF Then
                    If (Trim(!compositenumquestion) Like Trim(RemainingString)) And _
                            !itemnumber = Item Then
                            Questions = Questions & "~" & !question & "~" & !Option1 _
                                        & "~" & !Option2 & "~" & !Option3 _
                                        & "~" & !Option4 & "~" & !Option5 _
                                        & "~" & !correctanswer ' & "~" & !itemnumber
                            .MoveFirst
                            Item = Item + 1
                     End If
                End If
                .MoveNext
             Loop
          End With
          Socket(VarIndex).SendData "Questions" & VarTimer & LastItem & Questions
          List1.AddItem ("Send Questions Response")
        End If
' /**************** varAlreadyTakenFlag = 0
        
'************** if the quizform was taken already but not yet completed
'************** due to uncontrollable situations like power interruption
' **************** varAlreadyTakenFlag = 1
        If varAlreadyTakenFlag = 1 Then
           With rsLimit
              .MoveFirst
              .Find "compositenumquestion= '" & RemainingString & "'"
              If Not .EOF Then
                If Len(!completed) = 5 Then
                    LastItem = Right(!completed, 2)
                End If
                If Len(!completed) = 3 Then
                    LastItem = Right(!completed, 1)
                End If
              End If
           End With

           With rsQuestion
             .MoveFirst
             Do While Item <= LastItem
                If Not .EOF Then
                    If (Trim(!compositenumquestion) Like Trim(RemainingString)) And _
                            !itemnumber = Item Then
                            Questions = Questions & "~" & !question & "~" & !Option1 _
                                        & "~" & !Option2 & "~" & !Option3 _
                                        & "~" & !Option4 & "~" & !Option5 _
                                        & "~" & !correctanswer
                            .MoveFirst
                            Item = Item + 1
                     End If
                End If
                .MoveNext
             Loop
          End With
          If varUnfinishedSequence < 10 Then
             varPutzero = "0"
            Else
             varPutzero = ""
          End If
          Socket(VarIndex).SendData "Continue" & VarTimer & varPutzero & varUnfinishedSequence & varRoute & _
                           LastItem & Questions
         List1.AddItem ("Continue Quiz Response")
        End If
' /**************** varAlreadyTakenFlag = 1
      End If
' /******************************GoQuiz
' /******************************GoQuiz
' /******************************GoQuiz
      
 End If
Next x

'If Left(variableItem(x), 10) = "QuizAnswer" Then
'         RemainingString = Right(variableItem(x), Len(variableItem(x)) - 10)
'         MyPos = InStr(1, RemainingString, "~", 1)
'         With rsQuiz
'            .MoveFirst
'         End With
'       With rsQuiz
'            If Not .EOF Then
'                 .MoveFirst
'                 .Find "quizcompositekey= '" & Left(RemainingString, MyPos - 1) & "'"
'                 varFindCompositeKey = Left(RemainingString, MyPos - 1)
'                 If Not .EOF Then
'                     varDataQuizCompositeKey = Left(RemainingString, MyPos - 1)
'                     RemainingString = Right(RemainingString, Len(RemainingString) - MyPos)
'                     MyPos = InStr(1, RemainingString, "~", 1)
'                     varFindUserName = Left(RemainingString, MyPos - 1)
'
'                         If !studentusername = varFindUserName Then
'                            RemainingString = Right(RemainingString, Len(RemainingString) - MyPos)
'                            MyPos = InStr(1, RemainingString, "~", 1)
'                            If IsNull(!quizscore) Then
'                               !quizscore = "0"
'                            End If
'                            !quizscore = Val(!quizscore) + Left(RemainingString, MyPos - 1)
'                            VarQuizScore = !quizscore
'                            RemainingString = Right(RemainingString, Len(RemainingString) - MyPos)
'                            MyPos = InStr(1, RemainingString, "~", 1)
'                            !questionsequence = Left(RemainingString, MyPos - 1)
'                            !route = Right(RemainingString, 1)
'                            VarQuestionSequenceLength = Len(!questionsequence)
'                            .Update
'                            With rsLimit
'                              .MoveFirst
'                              .Find "compositenumquestion= '" & varDataQuizCompositeKey & "'"
'                              If Not .EOF Then
'                                  If Len(!completed) = 5 Then
'                                     If VarQuestionSequenceLength >= (Right(!completed, 2) * 2) Then
'                                        Socket(VarIndex).SendData "Finished" & VarQuizScore
'                                     End If
'                                  End If
'                                  If Len(!completed) = 3 Then
'                                      If VarQuestionSequenceLength >= (Right(!completed, 1) * 2) Then
'                                         Socket(VarIndex).SendData "Finished" & VarQuizScore
'                                      End If
'                                  End If
'                              End If
'                           End With
'                        Else
'                            .MoveNext
'                            .Find "quizcompositekey= '" & varFindCompositeKey & "'"
'                            If .EOF Then
'                                .AddNew
'                                !quizcompositekey = varFindCompositeKey
'                                RemainingString = Right(RemainingString, Len(RemainingString) - MyPos)
'                                MyPos = InStr(1, RemainingString, "~", 1)
'                                !studentusername = varFindUserName
'                                'RemainingString = Right(RemainingString, Len(RemainingString) - MyPos)
'                                'MyPos = InStr(1, RemainingString, "~", 1)
'                                !quizscore = !quizscore + Left(RemainingString, MyPos - 1)
'                                RemainingString = Right(RemainingString, Len(RemainingString) - MyPos)
'                                MyPos = InStr(1, RemainingString, "~", 1)
'                                !questionsequence = !questionsequence & Left(RemainingString, 2)
'                                !route = Right(RemainingString, 1)
'                                .Update
'
'                            End If
'                        End If
'
'                 Else
'                      .AddNew
'                      !quizcompositekey = Left(RemainingString, MyPos - 1)
'                      RemainingString = Right(RemainingString, Len(RemainingString) - MyPos)
'                      MyPos = InStr(1, RemainingString, "~", 1)
'                      !studentusername = Left(RemainingString, MyPos - 1)
'                      RemainingString = Right(RemainingString, Len(RemainingString) - MyPos)
'                      MyPos = InStr(1, RemainingString, "~", 1)
'                      !quizscore = !quizscore + Left(RemainingString, MyPos - 1)
'                      RemainingString = Right(RemainingString, Len(RemainingString) - MyPos)
'                      MyPos = InStr(1, RemainingString, "~", 1)
'                      !questionsequence = !questionsequence & Left(RemainingString, 1)
'                      !route = Right(RemainingString, 1)
'                      .Update
'
'                 End If
'            End If
'      End With
'   End If
If List1.ListCount >= 25 Then
  'With rsHistory
  '   .AddNew
  '   !Fdate = Date & "|" & Time
  '   !History = List1.Text
   '  .Update
  'End With
  List1.Clear
End If

End Sub

Private Sub Timer1_Timer()
lblConnections.Caption = List2.ListCount
End Sub
