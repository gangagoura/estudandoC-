VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRoomDetails 
   Caption         =   "Room Details"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   6330
   Begin VB.CommandButton Command3 
      Caption         =   "exit"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DeAllocate Student"
      Height          =   735
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Report"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      _Version        =   393216
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmRoomDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New Recordset

Private Sub Command1_Click()
    grd_dblClick
End Sub

Private Sub Command2_Click()
    Dim strStud As String
    Dim intAlloc As Integer

    If grd.Row < 0 Then
        MsgBox "Please Student to de-allocate!", vbCritical
        Exit Sub
    End If
    
    strStud = grd.TextMatrix(grd.Row, 3)
    
    sSQL = "delete from RoomAllocation where RoomID = '" & grd.TextMatrix(grd.Row, 2) & "' and MatricNo = '" & grd.TextMatrix(grd.Row, 3) & "'"
    cn.Execute (sSQL)
    
    sSQL = "update StudentInfo set Allocated = " & False & " where MatricNumber = '" & strStud & "'"
    cn.Execute sSQL
    
    sSQL = "select Allocated from Hostels where RoomId = '" & grd.TextMatrix(grd.Row, 2) & "'"
    Set rs = cn.Execute(sSQL)
    
    intAlloc = CInt(rs.Fields(0)) - 1
    
    sSQL = "update Hostels set Allocated = " & intAlloc & " where RoomId = '" & grd.TextMatrix(grd.Row, 2) & "'"
    cn.Execute sSQL

    frmHostelMgt.grd_dblClick
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Resize()
With grd
    .Height = ScaleHeight - 150
    .Width = ScaleHeight - .Top - 150
End With
End Sub

Private Sub grd_dblClick()
    With DataEnvironment1.rsCommand1
        If .State Then
            .Close
        End If
    End With
    DataEnvironment1.Command1 grd.TextMatrix(grd.Row, 3)
    dRptStudent.Show

End Sub
