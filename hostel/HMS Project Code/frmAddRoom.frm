VERSION 5.00
Begin VB.Form frmAddRoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Hostel Room"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "exit"
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Create Room and Close"
      Height          =   615
      Left            =   2520
      TabIndex        =   9
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmdCreateRoom 
      Caption         =   "Create Room and New..."
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Room Details"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
      Begin VB.ComboBox txtRoomSex 
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtRoomCapacity 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtRoomNumber 
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Sex of Room Members"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Room Capacity"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Room Number"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.ComboBox cboHostels 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Hostel"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsHostel As New Recordset

Private Sub cboHostels_Click()
sSQL = "select sex from hostelname where hostelname = '" & cboHostels & "'"
Set rsHostel = cn.Execute(sSQL)
If LCase$(rsHostel.Fields(0)) = "male" Then
    txtRoomSex.Clear
    txtRoomSex.AddItem "Male"
ElseIf LCase$(rsHostel.Fields(0)) = "female" Then
    txtRoomSex.Clear
    txtRoomSex.AddItem "Female"
Else
    txtRoomSex.Clear
    txtRoomSex.AddItem "Male"
    txtRoomSex.AddItem "Female"
End If

End Sub

'dim
Private Sub cmdCancel_Click()
    Create_Room
    Unload Me
End Sub

Private Sub cmdCreateRoom_Click()
    Create_Room
    Clear_Fields
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    txtRoomSex.Clear
    txtRoomSex.AddItem "Male"
    txtRoomSex.AddItem "Female"
    
    sSQL = "select * from HostelName"
    Set rsHostel = cn.Execute(sSQL)
    rsHostel.MoveFirst
    cboHostels.Clear
    Do While Not rsHostel.EOF
        cboHostels.AddItem rsHostel.Fields(1)
        rsHostel.MoveNext
    Loop
End Sub

Sub Create_Room()
    Dim intCapacity As Integer
    Dim strRoomNumber, strSex As String
    
    strRoomNumber = Me.txtRoomNumber
    intCapacity = CInt(txtRoomCapacity)
    strSex = Me.txtRoomSex
    
    If strRoomNumber = "" Then
        MsgBox "please enter an entry for the room number"
        Exit Sub
    End If
    
    If intCapacity = 0 Then
        MsgBox "please enter an entry for the room capacity"
        Exit Sub
    End If
    
    If strSex = "" Then
        MsgBox "please enter an entry for the room sex"
        Exit Sub
    End If
    
    
    mess = MsgBox("create room entry - number:" & strRoomNumber & " capacity:" & intCapacity & " members sex:" & strSex & " - in hostel:" & cboHostels.Text & "?", vbYesNo)
    If mess = vbNo Then
        Exit Sub
    End If
    
    sSQL = "select capacity from hostelname where hostelname = '" & cboHostels.Text & "'"
    Set rshotel = cn.Execute(sSQL)
    
    'insert room record
    sSQL = "insert into Hostels(RoomID, HostelName, RoomNumber, Capacity, Allocated, Sex) values ('" & Left$(cboHostels.Text, 1) & "-" & strRoomNumber & "','" & cboHostels.Text & "','" & strRoomNumber & "'," & intCapacity & ",0,'" & strSex & "')"
    cn.Execute sSQL
    
    sSQL = "select * from HostelName where HostelName='" & cboHostels & "'"
    Set rsHostel = cn.Execute(sSQL)
    
    Cap = CInt(rsHostel.Fields("capacity"))
    'update hostel parent record - total capacity
    sSQL = "update hostelname set capacity = " & Cap + CInt(intCapacity) & " where HostelName = '" & cboHostels & "'"
    cn.Execute sSQL
End Sub

Sub Clear_Fields()
    Me.txtRoomCapacity = ""
    Me.txtRoomNumber = ""
    Me.txtRoomSex = ""
End Sub
