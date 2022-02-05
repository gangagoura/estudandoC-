VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Hostel Allocation System"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8220
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   6435
      Left            =   0
      ScaleHeight     =   6375
      ScaleWidth      =   1935
      TabIndex        =   0
      Top             =   0
      Width           =   2000
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   4320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Rooms Management"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Hostel Management"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "New Hostel"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Student Records"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Student Registration"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuHyp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAllocation 
      Caption         =   "Allocation"
      Begin VB.Menu mnuAllocate 
         Caption         =   "Allocate"
      End
   End
   Begin VB.Menu mnuMgt 
      Caption         =   "Management"
      Begin VB.Menu mnuReg 
         Caption         =   "Student Registration"
      End
      Begin VB.Menu mnuMgtStudent 
         Caption         =   "Student Records"
      End
      Begin VB.Menu mnuHyp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateHostel 
         Caption         =   "Create Hostel"
      End
      Begin VB.Menu mnuHostelMgt 
         Caption         =   "Hostel Management"
      End
      Begin VB.Menu mnuHyp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRoomsMgt 
         Caption         =   "Rooms Management"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuRptViewAll 
         Caption         =   "All Allocated Rooms"
      End
      Begin VB.Menu mnuRptAllUn 
         Caption         =   "All Unallocated Rooms"
      End
      Begin VB.Menu mnuRptFully 
         Caption         =   "Fully Allocated Rooms"
      End
      Begin VB.Menu mnuRptPartial 
         Caption         =   "Partially Allocated Rooms"
      End
      Begin VB.Menu mnuHyp4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptHos 
         Caption         =   "Hostels"
      End
      Begin VB.Menu mnuHyp6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptStudDept 
         Caption         =   "Students List by Department"
      End
      Begin VB.Menu mnuRptStudDeptM 
         Caption         =   "Students List by Department (Male)"
      End
      Begin VB.Menu mnuRptStudDeptF 
         Caption         =   "Students List by Department (Female)"
      End
      Begin VB.Menu mnuRptStudSpc 
         Caption         =   "Students (Special Cases)"
      End
      Begin VB.Menu mnuhyp5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNonHND2 
         Caption         =   "Non HND II Students"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      Begin VB.Menu mnuCascade 
         Caption         =   "ReArrange"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuDocumentation 
         Caption         =   "Documentation"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmStudReg.Show
End Sub

Private Sub Command2_Click()
    frmStudRec.Show
End Sub

Private Sub Command3_Click()
    frmAddHostel.Show
End Sub

Private Sub Command4_Click()
    frmHostelMgt.Show
End Sub

Private Sub Command6_Click()
DataReport5.Show
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuAllocate_Click()
    MsgBox "Automatic Allocation wiating for Supervisor Recommendation!", vbInformation
    
End Sub

Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuCreateHostel_Click()
    frmAddHostel.Show
End Sub

Private Sub mnuDocumentation_Click()
    frmBrowser.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuRptHostels_Click()
    frmRptHostel.Show
End Sub

Private Sub mnuHostelMgt_Click()
    frmHostelMgt.Show
End Sub

Private Sub mnuMgtStudent_Click()
    frmStudRec.Show
End Sub

Private Sub mnuNonHND2_Click()
    DataReport9.Show
End Sub

Private Sub mnuReg_Click()
    frmStudReg.Show
End Sub

Private Sub mnuRptAllUn_Click()
    DataEnvironment1.Command4_Grouping
    DataReport2.Show
End Sub

Private Sub mnuRptFully_Click()
    DataEnvironment1.Command5_Grouping
    DataReport3.Show
End Sub

Private Sub mnuRptHos_Click()
    frmRptHostel.Show
End Sub

Private Sub mnuRptPartial_Click()
    DataEnvironment1.Command3_Grouping
    DataReport1.Show
End Sub

Private Sub mnuRptStudDept_Click()
    DataReport6.Show
End Sub

Private Sub mnuRptStudDeptF_Click()
DataReport7.Show
End Sub

Private Sub mnuRptStudDeptM_Click()
    DataReport10.Show
End Sub

Private Sub mnuRptStudSpc_Click()
DataReport8.Show
End Sub

Private Sub mnuRptViewAll_Click()
    DataEnvironment1.Command7_Grouping
    DataReport4.Show
End Sub
