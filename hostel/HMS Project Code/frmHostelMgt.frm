VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHostelMgt 
   Caption         =   "Hostel Management"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   10380
   Begin VB.CommandButton Command3 
      Caption         =   "exit"
      Height          =   615
      Left            =   4800
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cboStud 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdAlloc 
      Caption         =   "Allocate Room"
      Height          =   735
      Left            =   8640
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Only Allocated Rooms"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   735
      Left            =   3240
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtRoomSearch 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Room"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7435
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.ComboBox cboHostel 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label4 
      Height          =   1095
      Left            =   6480
      TabIndex        =   12
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "click to view student matric no"
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Search for Room"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Select Hostel"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmHostelMgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String
Dim rs As New Recordset
Dim rsStud1 As New Recordset

Private Sub cboHostel_Click()
    loadRS
End Sub

Private Sub cboStud_Click()
    sSQL = "select MatricNumber, LastName, Firstname, Sex, DepartMent from StudentInfo where MatricNumber = '" & cboStud.Text & "'"
    Set rs = cn.Execute(sSQL)
    Label4 = ""
    Label4 = Label4 & "Matric Number: " & rs.Fields("MatricNumber") & vbCrLf
    Label4 = Label4 & "(" & rs.Fields("LastName") & ", " & rs.Fields("Firstname") & ")" & vbCrLf
    Label4 = Label4 & "Sex: " & rs.Fields("Sex") & vbCrLf
    Label4 = Label4 & rs.Fields("Department")
    
End Sub

Private Sub Check1_Click()
    loadRS
End Sub

Private Sub cmdAlloc_Click()
    Dim strStud As String
    Dim intAlloc As Integer

    If cboStud.Text = "" Then
        MsgBox "Select a student's matric number first!", vbExclamation
        Exit Sub
    End If
    
    If grd.Row < 0 Then
        MsgBox "Please select a room to allocate!", vbCritical
        Exit Sub
    End If
    
    strStud = cboStud.Text
    sSQL = "select Allocated, Sex from Studentinfo where MatricNumber = '" & strStud & "'"
    Set rsStud1 = cn.Execute(sSQL)
    If rsStud1.BOF And rsStud1.EOF Then
        MsgBox "sorry, no student like that..."
        Exit Sub
    End If
    If rsStud1.Fields("Allocated") = True Then
        MsgBox "Student already allocated to a room.", vbCritical
        Exit Sub
    Else
        'check students sex
        If LCase$(grd.TextMatrix(grd.Row, 7)) <> LCase$(rsStud1.Fields("Sex")) Then
            MsgBox "Wrong Sex Allocation... Check Student's Sex", vbCritical
            Exit Sub
        End If
        
        'do the allocation here
        
        intAlloc = CInt(grd.TextMatrix(grd.Row, 6))
        intCapacity = CInt(grd.TextMatrix(grd.Row, 5))
        If intAlloc = intCapacity Then
            MsgBox "Sorry, room full!", vbCritical
            Exit Sub
        End If
        
        intAlloc = CInt(grd.TextMatrix(grd.Row, 6)) + 1
        
        'insert to room allocation table
        sSQL = "insert into RoomAllocation(RoomId, MatricNo) values ('" & grd.TextMatrix(grd.Row, 2) & "','" & strStud & "')"
        cn.Execute sSQL
        
        'indicate in student table that student is allocated
        sSQL = "update StudentInfo set Allocated = " & True & " where MatricNumber = '" & strStud & "'"
        cn.Execute sSQL
        
        'update current status of room allocation
        
        sSQL = "update Hostels set Allocated = " & intAlloc & " where RoomId = '" & grd.TextMatrix(grd.Row, 2) & "'"
        cn.Execute sSQL
        
        loadRS
        loadUnAllocated
    End If
End Sub

Private Sub Command1_Click()
    Load frmAddRoom
    frmAddRoom.cboHostels.Text = cboHostel.Text
    frmAddRoom.Show
    loadRS
End Sub

Private Sub Command2_Click()
    loadRS
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Activate()
loadRS
End Sub

Private Sub Form_Load()
    sSQL = "select hostelname from hostelname"
    Set rs = cn.Execute(sSQL)
    
    With cboHostel
        .Clear
        Do While Not rs.EOF
            .AddItem rs.Fields(0)
            rs.MoveNext
        Loop
    End With
    
    loadUnAllocated
End Sub

Private Sub Form_Resize()
    With grd
        .Height = ScaleHeight - .Top - 300
        .Width = ScaleWidth - 300
    End With
End Sub

Public Sub grd_dblClick()
    'MsgBox grd.TextMatrix(grd.Row, 2)
    
    Dim rs1 As New Recordset
    sSQL = "select id,RoomId,MatricNo from RoomAllocation where RoomID = '" & grd.TextMatrix(grd.Row, 2) & "'"
    Set rs1 = cn.Execute(sSQL)
    If rs1.EOF And rs1.BOF Then
        MsgBox "No Allocation for Room ", vbCritical
        Exit Sub
    End If
    Load frmRoomDetails
    With frmRoomDetails
        LoadRecordsetIntoGrid rs1, .grd
        .Show
    End With
End Sub

Private Sub txtRoomSearch_Change()
    loadRS
End Sub

Public Sub loadRS()
    If cboHostel.Text = "" Then Exit Sub
    sSQL = "select * from hostels where hostelname = '" & cboHostel.Text & "'"
    
    If txtRoomSearch <> "" Then
        sSQL = sSQL & " and RoomNumber like '" & txtRoomSearch.Text & "%'"
    End If

    If Check1.Value = vbChecked Then
        sSQL = sSQL & " and allocated > 0"
    End If
    
loadRS1:
    Set rs = cn.Execute(sSQL)
    If rs.EOF And rs.BOF Then
        MsgBox "Room Information not found", vbCritical
        sSQL = "select * from hostels where hostelname = '" & cboHostel.Text & "'"
        'GoTo loadRS1
    Else
        LoadRecordsetIntoGrid rs, grd
    End If
    
    loadUnAllocated
End Sub


Sub loadUnAllocated()
    'On Error Resume Next
    cboStud.Clear
    sSQL = "select MatricNumber from StudentInfo where Allocated = " & False
    Set rs = cn.Execute(sSQL)
    Do While Not rs.EOF
        cboStud.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
End Sub
