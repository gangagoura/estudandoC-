VERSION 5.00
Begin VB.Form frmStudReg 
   Caption         =   "New Student Entry..."
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   9510
   Begin VB.CommandButton Command3 
      Caption         =   "New Entry"
      Height          =   615
      Left            =   3360
      TabIndex        =   34
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   615
      Left            =   7920
      TabIndex        =   33
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   5040
      TabIndex        =   29
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton Option2 
         Caption         =   "Female"
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Male"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Special Case                                                (including SUG, Sports Persons, Handicapped...)"
         Height          =   615
         Left            =   360
         TabIndex        =   30
         Top             =   600
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Entry"
      Height          =   615
      Left            =   5640
      TabIndex        =   28
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Allocation Details"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   360
      TabIndex        =   27
      Top             =   3720
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Next of Kin"
      Height          =   2055
      Left            =   5040
      TabIndex        =   22
      Top             =   3480
      Width           =   4215
      Begin VB.TextBox txtNext 
         Height          =   375
         Left            =   1200
         TabIndex        =   24
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtNextAdd 
         Height          =   1095
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label12 
         Caption         =   "Next of Kin"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Next of Kin's Address"
         Height          =   615
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sponsor"
      Height          =   2055
      Left            =   5040
      TabIndex        =   7
      Top             =   1320
      Width           =   4215
      Begin VB.TextBox txtSponsorAdd 
         Height          =   1095
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtSponsor 
         Height          =   375
         Left            =   1200
         TabIndex        =   19
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label10 
         Caption         =   "Sponsor's Address"
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Sponsor"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Academic Info"
      Height          =   2295
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   4575
      Begin VB.ComboBox CboCourse 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Text            =   "CboCourse"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtCGPA 
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox cboLevel 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Text            =   "cboLevel"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox cboDept 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Text            =   "cboDept"
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cboSchool 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Text            =   "cboSchool"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "CGPA"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Level:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Course of Study"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Department:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "School:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox txtRegNo 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtFirstname 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtSurname 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Registration Number:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "First Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Surname:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmStudReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String
Dim rs As New Recordset

Private Sub cboDept_Click()
    sSQL = "select CourseofStudy from Courses where school = '" & cboSchool & "' and department = '" & cboDept & "'"
    Set rs = cn.Execute(sSQL)
    With CboCourse
        .Clear
        Do While Not rs.EOF
            .AddItem rs.Fields(0)
            rs.MoveNext
        Loop
    End With
End Sub

Private Sub cboSchool_Click()
    sSQL = "Select department from departments where school = '" & cboSchool & "'"
    Set rs = cn.Execute(sSQL)
    With cboDept
        .Clear
        Do While Not rs.EOF
            .AddItem rs.Fields(0)
            rs.MoveNext
        Loop
    End With
End Sub

Private Sub Command1_Click()
    Dim Special As Boolean
    Dim Sex As String
    If Option1 Then
        Sex = "Male"
    ElseIf Option2 Then
        Sex = "Female"
    Else
        MsgBox "Please select the Sex of the Student!", vbCritical
        Exit Sub
    End If
    
    If Check1.Value = vbChecked Then
        Special = True
    Else
        Special = False
    End If
    
    On Error GoTo ErrorHandler
    cn.BeginTrans
    
    sSQL = "insert into studentinfo (MatricNumber,LastName,FirstName,Sex,School,Department,tblLevel,Allocated,CourseofStudy,NameofSponsor,SponsorAddress,NextofKin,KinAddress,ReceiptNo,Special,CGPA) values('" & _
    txtRegNo & "','" & txtSurname & "','" & txtFirstname & "','" & Sex & "','" & cboSchool & _
    "','" & cboDept & "','" & cboLevel & "', false,'" & CboCourse & "','" & txtSponsor & "','" & _
    txtSponsorAdd & "','" & txtNext & "','" & txtNextAdd & "',0," & Special & ", " & txtCGPA & ")"
    
    'MsgBox sSQL
    
    cn.Execute sSQL
    
    MsgBox "Entry created"
    cn.CommitTrans
    
    Command3_Click
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
    cn.RollbackTrans
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    Me.txtRegNo = ""
    Me.txtSurname = ""
    Me.txtFirstname = ""
    Me.Option1 = False
    Me.Option2 = False
    Me.cboSchool = ""
    Me.cboDept = ""
    Me.cboLevel = ""
    Me.CboCourse = ""
    Me.txtSponsor = ""
    Me.txtSponsorAdd = ""
    Me.txtNext = ""
    Me.txtNextAdd = ""
    Me.Check1 = vbUnchecked
    Me.txtCGPA = ""
End Sub

Private Sub Form_Load()
    sSQL = "select  school from schools"
    Set rs = cn.Execute(sSQL)
    With cboSchool
        .Clear
        Do While Not rs.EOF
            .AddItem rs.Fields(0)
            rs.MoveNext
        Loop
    End With
    
    With cboLevel
        .Clear
        .AddItem "ND I FT"
        .AddItem "ND II FT"
        .AddItem "HND I FT"
        .AddItem "HND II FT"
        .AddItem "ND I PT"
        .AddItem "ND II PT"
        .AddItem "ND III PT"
        .AddItem "HND I PT"
        .AddItem "HND II PT"
        .AddItem "HND III PT"
        .AddItem "TTTP"
        .AddItem "PGD"
    End With
End Sub

