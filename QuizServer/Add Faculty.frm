VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFaculty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faculty"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "English Police"
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtFacultyFname 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   3720
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Add Faculty.frx":0000
      Left            =   1680
      List            =   "Add Faculty.frx":000A
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&l"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   2760
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
      TabIndex        =   13
      Top             =   4080
      Width           =   5175
   End
   Begin VB.ListBox ListFaculty 
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
      ItemData        =   "Add Faculty.frx":0026
      Left            =   120
      List            =   "Add Faculty.frx":002D
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   1800
      Width           =   3615
   End
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
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtFacultyLName 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtFacultyID 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Name"
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lvl"
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ID Number"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      X1              =   6480
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line9 
      X1              =   6480
      X2              =   3720
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   6480
      X2              =   3720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line7 
      X1              =   6480
      X2              =   3720
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000E&
      X1              =   6120
      X2              =   6480
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line5 
      X1              =   6120
      X2              =   6480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   6480
      X2              =   5880
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line3 
      X1              =   6480
      X2              =   5880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   3720
      X2              =   4320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   4320
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label5 
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
      Left            =   3360
      TabIndex        =   15
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Level:"
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
      TabIndex        =   14
      Top             =   1200
      Width           =   1455
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
      TabIndex        =   12
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label2 
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
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Faculty ID:"
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
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmFaculty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
Unload Me
End Sub
Private Sub cmdAdd_Click()
txtFacultyID.Enabled = True

txtFacultyLName.Enabled = True
txtFacultyFname.Enabled = True
If cmdAdd.Caption = "&Add" Then
   Check1.Value = 0
   cmdAdd.Caption = "&Save"
   cmdEdit.Enabled = False
   txtFacultyID.SetFocus
ElseIf cmdAdd.Caption = "&Save" Then
   cmdAdd.Caption = "&Add"
   With rsFaculty
      .MoveFirst
      .Find "facultyid= '" & Trim(txtFacultyID.Text) & "'"
      If Not .EOF Then
         MsgBox "The Faculty ID already exists!", vbCritical
      Else
         .AddNew
         !facultyid = txtFacultyID.Text
         !facultylName = txtFacultyLName.Text
         !facultyfname = txtFacultyFname.Text
         !UserName = UCase(Trim(txtFacultyLName.Text) & Left(txtFacultyFname.Text, 1))
         !facultyPassword = UCase("pass")
         !facultyviolations = "0"
         If Combo1.ListIndex = -1 Then
            !Level = 0
         End If
         If Combo1.ListIndex <> -1 Then
            !Level = Combo1.ListIndex
         End If
         If Check1.Value = 1 Then
            !Level = 1
         End If
         .Update
      End If
    End With
Call Populate_List
txtFacultyID.Text = ""
txtFacultyFname.Text = ""
txtFacultyLName.Text = ""
cmdEdit.Enabled = True
End If
End Sub
Sub Populate_List()
ListFaculty.Clear
With rsFaculty
  If Not .EOF Then
     .MoveFirst
  End If
Do Until .EOF
   ListFaculty.AddItem !facultyid & Space(12 - Len(!facultyid)) & " | " & !Level & " | " & !UserName
   .MoveNext
Loop
End With
End Sub
Private Sub cmdDelete_Click()
With rsFaculty
  .MoveFirst
  .Find "facultyid= '" & Left(txtFireMode.Text, 10) & "'"
  If Not .EOF Then
     If MsgBox("Do you really want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
        If GlobalFacultyUserRights <= !Level Then
          MsgBox "You cannot delete any user with the same or greater rights level!", vbCritical
          Call Populate_List
          txtFireMode.Text = ""
          Exit Sub
        End If
        
        .Delete
        Call Populate_List
        txtFireMode.Text = ""
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
  txtFacultyID.Enabled = True
  txtFacultyLName.Enabled = True
  txtFacultyFname.Enabled = True
  txtFacultyID.SetFocus
  cmdAdd.Enabled = False
  cmdEdit.Caption = "&Save"
With rsFaculty
  .MoveFirst
  .Find "facultyid= '" & Left(txtFireMode.Text, 10) & "'"
  If Not .EOF Then
      txtFacultyID.Text = !facultyid
      txtFacultyLName.Text = !facultylName
      txtFacultyFname.Text = !facultyfname
      If !Level = 0 Then
        Combo1.Text = "Faculty"
       Else
        Combo1.Text = "Administrator"
      End If
      If GlobalFacultyUserRights <= !Level Then
         MsgBox "You cannot modify any user with the same or greater rights level!", vbCritical
         cmdAdd.Enabled = True
         Call Populate_List
         cmdEdit.Caption = "&Edit"
         txtFacultyID.Enabled = False
         txtFacultyLName.Enabled = False
         txtFacultyFname.Enabled = False
         Exit Sub
      End If
    Else
      MsgBox "No Record to Edit", vbInformation
   End If
End With
Else
  With rsFaculty
    !facultyid = txtFacultyID.Text
    !facultylName = txtFacultyLName.Text
    !facultyfname = txtFacultyFname.Text
    If Combo1.ListIndex = 0 Then
       !Level = 0
     Else
       !Level = 2
    End If
    .Update
  End With
  cmdAdd.Enabled = True
  Call Populate_List
  cmdEdit.Caption = "&Edit"
  txtFacultyID.Enabled = False
  txtFacultyLName.Enabled = False
  txtFacultyFname.Enabled = False
End If
End Sub
Private Sub cmdCancel_Click()
With rsFaculty
   .CancelUpdate
End With
txtFacultyID.Text = ""
txtFacultyLName.Text = ""
txtFacultyFname.Text = ""
txtFacultyID.Enabled = False
txtFacultyLName.Enabled = False
txtFacultyFname.Enabled = False
cmdAdd.Enabled = True
cmdEdit.Enabled = True
cmdAdd.Caption = "&Add"
cmdEdit.Caption = "&Edit"
End Sub
Private Sub Form_Load()
strCn = "DSN=DSNSample;server=server;uid=sa;pwd=touch;database=OnlynQuiz"
Set strCn1 = New ADODB.Connection
strCn1.Open strCn
Call Initialization
Call Populate_List
End Sub
Private Sub ListFaculty_Click()
txtFireMode.Text = ListFaculty.Text
End Sub
Private Sub txtFacultyFname_LostFocus()
cmdAdd.SetFocus
End Sub

