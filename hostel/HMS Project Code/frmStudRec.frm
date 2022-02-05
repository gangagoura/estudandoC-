VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStudRec 
   Caption         =   "Students' Record"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   8700
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   855
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   6615
      Begin VB.CommandButton Command4 
         Caption         =   "Delete Student's Record"
         Height          =   495
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Edit Student Details"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   495
      Left            =   7080
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search by Name"
      Height          =   1935
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   615
         Left            =   1200
         TabIndex        =   10
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Surname"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Department Select"
      Height          =   1935
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      Begin VB.ComboBox cboSchool 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox cboDEpt 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Select Department"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Select School"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8705
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Doubleclick on a record below to view the detailed information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      TabIndex        =   11
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmStudRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String
Dim rs As New Recordset

Private Sub cboDept_Click()
    loadRS
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
    loadRS
End Sub

Private Sub Command1_Click()
    loadRS
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    Load frmStudReg1
    '
    ' assign details to controls
    On Error Resume Next
    With frmStudReg1
        .txtRegNo = grd.TextMatrix(grd.Row, 2)
        MatricNo = grd.TextMatrix(grd.Row, 2)
        .txtSurname = grd.TextMatrix(grd.Row, 3)
        .txtFirstname = grd.TextMatrix(grd.Row, 4)
        If grd.TextMatrix(grd.Row, 5) = "Male" Then
            .Option1 = True
            .Option2 = False
        Else
            .Option1 = False
            .Option2 = True
        End If
        .cboSchool = grd.TextMatrix(grd.Row, 6)
        .cboDept = grd.TextMatrix(grd.Row, 7)
        .cboLevel = grd.TextMatrix(grd.Row, 8)
        .CboCourse = grd.TextMatrix(grd.Row, 10)
        .txtSponsor = grd.TextMatrix(grd.Row, 11)
        .txtSponsorAdd = grd.TextMatrix(grd.Row, 12)
        .txtNext = grd.TextMatrix(grd.Row, 13)
        .txtNextAdd = grd.TextMatrix(grd.Row, 14)
        .txtReceipt = grd.TextMatrix(grd.Row, 15)
        If grd.TextMatrix(grd.Row, 16) = True Then
            .Check1.Value = vbChecked
        Else
            .Check1.Value = vbUnchecked
        End If
        .txtCGPA = grd.TextMatrix(grd.Row, 17)
        
        sSQL = "select RoomID from RoomAllocation where MatricNo = '" & grd.TextMatrix(grd.Row, 2) & "'"
        Set rs = cn.Execute(sSQL)
        .Text1 = rs.Fields("RoomID")
    '
    ' show form
        .Show
    End With
End Sub

Private Sub Command4_Click()
    Dim s As String
    If grd.Row = 0 Then Exit Sub
    If grd.TextMatrix(grd.Row, 9) = True Then
        MsgBox "Please deallocate student.", vbInformation
        Load frmHostelMgt
        With frmHostelMgt
            .Check1.Value = vbChecked
            
            sSQL = "select * from RoomAllocation where MatricNo = '" & grd.TextMatrix(grd.Row, 2) & "'"
            Set rs = cn.Execute(sSQL)
            
            s = rs.Fields("RoomID")
            s = Right$(s, Len(s) - 2)
            .txtRoomSearch.Text = s
            
            sSQL = "select HostelName from HostelName where Prefix = '" & UCase$(Left$(rs.Fields("RoomID"), 1)) & "'"
            Set rs = cn.Execute(sSQL)
            .cboHostel.Text = rs.Fields("HostelName")
            
            .loadRS
            .Show
        End With
        Exit Sub
    End If
    sSQL = "delete from studentinfo where matricnumber = '" & grd.TextMatrix(grd.Row, 2) & "'"
    If MsgBox("Are you sure u want 2 delete this students record", vbQuestion + vbYesNo + vbDefaultButton2, "Irreversible Action") = vbNo Then
        Exit Sub
    End If
    cn.Execute sSQL
    loadRS
End Sub

Private Sub Form_Activate()
loadRS
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
    loadRS
End Sub

Public Sub loadRS()
    If cboSchool.Text <> "" Then
        sSQL = "select * from studentinfo where school = '" & cboSchool & "'"
    Else
        sSQL = "select * from studentinfo"
    End If
    
    If cboDept.Text <> "" Then
        If cboSchool.Text = "" Then
            sSQL = sSQL & " where department = '" & cboDept & "'"
        Else
            sSQL = sSQL & " and department = '" & cboDept & "'"
        End If
    End If
    
    If Text1.Text <> "" Then
        If cboSchool.Text = "" Then
            sSQL = sSQL & " where lastname like '" & Text1 & "%'"
        Else
            sSQL = sSQL & " and lastname like '" & Text1 & "%'"
        End If
    End If
        
    Set rs = cn.Execute(sSQL)
    LoadRecordsetIntoGrid rs, grd
End Sub

Private Sub Form_Resize()
    With grd
        .Width = ScaleWidth - 200
        .Height = ScaleHeight - .Top - 200
    End With
End Sub

Private Sub grd_dblClick()
    With DataEnvironment1.rsCommand1
        If .State Then
            .Close
        End If
    End With
    DataEnvironment1.Command1 grd.TextMatrix(grd.Row, 2)
    dRptStudent.Show

End Sub
