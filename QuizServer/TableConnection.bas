Attribute VB_Name = "TableConnection"
Public Qctr As Integer
Public facultyPassword As String

'Set Cnn1 = New ADODB.Connection
'Set Rsstude = CreateObject("ADODB.recordset")
'strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;pwd=touch;Data Source=OnlynQuiz.mdb"
'Cnn1.Open strCnn


' ---------------
Public strCn As String
Public strCn1 As New ADODB.Connection
Public rsFaculty As ADODB.Recordset
Public cnConnection As New ADODB.Connection
' ---------------
Public strCnC As String
Dim strCn1C As New ADODB.Connection
Public rsCourse As ADODB.Recordset
Public cnConnectionC As New ADODB.Connection
' ----------------
Public strCnS As String
Public strCn1S As New ADODB.Connection
Public rsStudent As ADODB.Recordset
Public cnConnectionS As New ADODB.Connection
' ----------------
Public strCnQ As String
Public strCn1Q As New ADODB.Connection
Public rsQuiz As ADODB.Recordset
Public cnConnectionQ As New ADODB.Connection

' ----------------
Public strCnViolation As String
Public strCn1Violation As New ADODB.Connection
Public rsViolation As ADODB.Recordset
Public cnConnectionViolation As New ADODB.Connection
' ----------------
Public strCnQuery As String
Public strCn1Query As New ADODB.Connection
Public rsQuizQuery As ADODB.Recordset
Public cnConnectionQuery As New ADODB.Connection
' ----------------
Public strCnQu As String
Public strCn1Qu As New ADODB.Connection
Public rsQuestion As ADODB.Recordset
Public cnConnectionQu As New ADODB.Connection
' ----------------
Public strCnLimit As String
Public strCn1Limit As New ADODB.Connection
Public rsLimit As ADODB.Recordset
Public cnConnectionLimit As New ADODB.Connection

Public strCnHistory As String
Public strCn1History As New ADODB.Connection
Public rsHistory As ADODB.Recordset
Public cnConnectionHistory As New ADODB.Connection
Sub Initialization()
Set cnConnection = New ADODB.Connection
cnConnection.Open strCn
Set rsFaculty = New ADODB.Recordset

With rsFaculty
   .CursorLocation = adUseClient
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .ActiveConnection = cnConnection
   .Open "tblFaculty", , , , adCmdTable
   If .EOF = False Then
      .MoveFirst
    End If
End With
End Sub
Sub Initialization_Course()
Set cnConnectionC = New ADODB.Connection
cnConnectionC.Open strCnC
Set rsCourse = New ADODB.Recordset

With rsCourse
   .CursorLocation = adUseClient
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .ActiveConnection = cnConnectionC
   .Open "tblCourse", , , , adCmdTable
   If .EOF = False Then
      .MoveFirst
    End If
End With
End Sub
Sub Initialization_Student()
Set cnConnectionS = New ADODB.Connection
cnConnectionS.Open strCnS
Set rsStudent = New ADODB.Recordset

With rsStudent
   .CursorLocation = adUseClient
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .ActiveConnection = cnConnectionS
   .Open "tblstudent", , , , adCmdTable
   If .EOF = False Then
      .MoveFirst
    End If
End With
End Sub
Sub Initialization_QueryViolation()
Set cnConnectionViolation = New ADODB.Connection
cnConnectionViolation.Open strCnViolation
Set rsViolation = New ADODB.Recordset

With rsViolation
   .CursorLocation = adUseClient
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .ActiveConnection = cnConnectionViolation
   .Open "queryviolation", , , , adCmdTable
              
   If .EOF = False Then
      .MoveFirst
    End If
End With
End Sub

Sub Initialization_QuizQuery()
Set cnConnectionQuery = New ADODB.Connection
cnConnectionQuery.Open strCnQuery
Set rsQuizQuery = New ADODB.Recordset

With rsQuizQuery
   .CursorLocation = adUseClient
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .ActiveConnection = cnConnectionQuery
   .Open "queryquiz", , , , adCmdTable
              
   If .EOF = False Then
      .MoveFirst
    End If
End With
End Sub

Sub Initialization_Quiz()
Set cnConnectionQ = New ADODB.Connection
cnConnectionQ.Open strCnQ
Set rsQuiz = New ADODB.Recordset

With rsQuiz
   .CursorLocation = adUseClient
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .ActiveConnection = cnConnectionQ
   .Open "tblquiz", , , , adCmdTable
   If .EOF = False Then
      .MoveFirst
    End If
End With
End Sub
Sub Initialization_Question()
Set cnConnectionQu = New ADODB.Connection
cnConnectionQu.Open strCnQu
Set rsQuestion = New ADODB.Recordset

With rsQuestion
   .CursorLocation = adUseClient
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .ActiveConnection = cnConnectionQu
   .Open "tblquestion", , , , adCmdTable
   If .EOF = False Then
      .MoveFirst
    End If
End With
End Sub
Sub Initialization_Limit()
Set cnConnectionLimit = New ADODB.Connection
cnConnectionLimit.Open strCnLimit
Set rsLimit = New ADODB.Recordset

With rsLimit
   .CursorLocation = adUseClient
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .ActiveConnection = cnConnectionLimit
   .Open "tblQuizFormLimitation", , , , adCmdTable
   If .EOF = False Then
      .MoveFirst
    End If
End With
End Sub

Sub Initialization_History()
Set cnConnectionHistory = New ADODB.Connection
cnConnectionHistory.Open strCnHistory
Set rsHistory = New ADODB.Recordset

With rsHistory
   .CursorLocation = adUseClient
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .ActiveConnection = cnConnectionHistory
   .Open "tblServerHistory", , , , adCmdTable
   If .EOF = False Then
      .MoveFirst
    End If
End With
End Sub
