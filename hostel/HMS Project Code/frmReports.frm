VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReports 
   Caption         =   "Some Reports..."
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   480
      ScaleHeight     =   855
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   0
      Width           =   15
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8493
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   480
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New Recordset

Private Sub Command1_Click()

sSQL = "SELECT MatricNumber, LastName, FirstName, Sex, School, Department, CourseofStudy,  Allocated FROM StudentInfo where allocated = " & True
Set rs = cn.Execute(sSQL)
LoadRecordsetIntoGrid rs, grd
End Sub

Private Sub Form_Resize()
    With grd
        .Height = ScaleHeight - .Top - 300
        .Width = ScaleWidth - 300
    End With

End Sub
