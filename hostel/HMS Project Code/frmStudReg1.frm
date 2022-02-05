VERSION 5.00
Begin VB.Form frmStudReg1 
   Caption         =   "Edit Student Entry..."
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9360
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   615
      Left            =   7680
      TabIndex        =   33
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   5040
      TabIndex        =   29
      Top             =   120
      Width           =   4095
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
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Edited Entry"
      Height          =   615
      Left            =   5400
      TabIndex        =   28
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Allocation Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   4575
      Begin VB.TextBox txtReceipt 
         Height          =   285
         Left            =   1800
         TabIndex        =   36
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label14 
         Caption         =   "Receipt Number"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Room Identification"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Next of Kin"
      Height          =   2055
      Left            =   4800
      TabIndex        =   22
      Top             =   3480
      Width           =   4335
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
      Left            =   4800
      TabIndex        =   7
      Top             =   1320
      Width           =   4335
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
      Left            =   120
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
      Width           =   1575
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
Attribute VB_Name = "frmStudReg1"
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
    
    'CREATE many sQL statements here for update
    sSQL = "update studentinfo set LastName = '" & txtSurname & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set FirstName = '" & txtFirstname & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set Sex = '" & Sex & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set School = '" & cboSchool & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set department = '" & cboDept & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set tbllevel = '" & cboLevel & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set courseofstudy = '" & CboCourse & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set nameofsponsor = '" & txtSponsor & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set sponsoraddress = '" & txtSponsorAdd & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set nextofkin = '" & txtNext & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set kinaddress = '" & txtNextAdd & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    If txtReceipt = "" Then
        txtReceipt = 0
    End If
    sSQL = "update studentinfo set receiptno = '" & txtReceipt & "' where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set special = " & Special & " where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    sSQL = "update studentinfo set cgpa = " & txtCGPA & " where MatricNumber = '" & MatricNo & "'"
    cn.Execute sSQL
    
    'sSQL = "update studentinfo set matricnumber = '" & txtRegNo & "' where MatricNumber = '" & MatricNo & "'"
    'cn.Execute sSQL
    
    MsgBox "Entry saved!"
    cn.CommitTrans
    
    frmStudRec.loadRS
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
    cn.RollbackTrans
End Sub

Private Sub Command2_Click()
Unload Me
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

