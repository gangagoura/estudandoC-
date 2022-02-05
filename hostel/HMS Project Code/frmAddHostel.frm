VERSION 5.00
Begin VB.Form frmAddHostel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create New Hostel Entry"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtCapacity 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Hostel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtNickname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   1080
      Width           =   3135
   End
   Begin VB.ComboBox cboSex 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtPrefix 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtHostelName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Capacity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Hostel Nickname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Prefix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Hostel Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddHostel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String

Private Sub Command1_Click()
    On Error GoTo ErrorHandler
    
    cn.BeginTrans
    sSQL = "insert into hostelname(HostelName, HostelNickname, Sex, Prefix, Capacity, CapacityUsed) values ('" & txtHostelName & "','" & txtNickname & "','" & cboSex & "','" & txtPrefix & "', " & CInt(txtCapacity) & ", 0)"
    
    cn.Execute sSQL
    MsgBox "Hostel " & txtHostelName & " created.", vbInformation
    cn.CommitTrans
    Exit Sub
ErrorHandler:
    cn.RollbackTrans
    MsgBox "Hostel " & txtHostelName & " not created.", vbInformation
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    With cboSex
        .Clear
        .AddItem "Female"
        .AddItem "Male"
        .AddItem "Mixed"
    End With
End Sub

Private Sub txtHostelName_Change()
    txtPrefix = Left$(txtHostelName, 1)
    Command1.Caption = "Create Hostel '" & txtHostelName & "'"
End Sub
