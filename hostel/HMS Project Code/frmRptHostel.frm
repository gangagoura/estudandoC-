VERSION 5.00
Begin VB.Form frmRptHostel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Hostels Report."
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "View"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "frmRptHostel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As New Recordset

Private Sub Command1_Click()
    If Combo1.Text = "" Then
        MsgBox "Please select a Hostel name!", vbCritical
        Exit Sub
    End If
    
    With DataEnvironment1.rsCommand2
        If .State Then
            .Close
        End If
    End With
    DataEnvironment1.Command2 Combo1.Text
    dRptHostel.Show
End Sub

Private Sub Form_Load()
    sSQL = "select hostelname from hostelname"
    Set rs = cn.Execute(sSQL)
    Combo1.Clear
    With rs
        Do While Not .EOF
            Combo1.AddItem .Fields(0)
            .MoveNext
        Loop
    End With
End Sub
