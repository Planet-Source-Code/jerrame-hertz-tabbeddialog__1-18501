VERSION 5.00
Begin VB.Form TabbedDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabbed Dialog"
   ClientHeight    =   5295
   ClientLeft      =   7545
   ClientTop       =   4455
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      Height          =   375
      Index           =   3
      Left            =   6480
      TabIndex        =   13
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   12
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   10
      Top             =   4800
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   4590
      Left            =   100
      ScaleHeight     =   4590
      ScaleWidth      =   7965
      TabIndex        =   0
      Top             =   100
      Width           =   7970
      Begin VB.Frame Box 
         BorderStyle     =   0  'None
         Height          =   4215
         Index           =   0
         Left            =   10
         TabIndex        =   4
         Top             =   360
         Width           =   7935
         Begin VB.Frame Frame1 
            Caption         =   "Tab 1"
            Height          =   3975
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   7695
         End
      End
      Begin VB.Frame Box 
         BorderStyle     =   0  'None
         Height          =   4215
         Index           =   3
         Left            =   10
         TabIndex        =   3
         Top             =   360
         Width           =   7935
         Begin VB.Frame Frame2 
            Caption         =   "Tab 4"
            Height          =   3975
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   7695
         End
      End
      Begin VB.Frame Box 
         BorderStyle     =   0  'None
         Height          =   4215
         Index           =   2
         Left            =   10
         TabIndex        =   2
         Top             =   360
         Width           =   7935
         Begin VB.Frame Frame3 
            Caption         =   "Tab 3"
            Height          =   3975
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   7695
         End
      End
      Begin VB.Frame Box 
         BorderStyle     =   0  'None
         Height          =   4215
         Index           =   1
         Left            =   10
         TabIndex        =   1
         Top             =   360
         Width           =   7935
         Begin VB.Frame Frame4 
            Caption         =   "Tab 2"
            Height          =   3975
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   7695
         End
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5790
         TabIndex        =   9
         Top             =   0
         Width           =   2145
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Tab 1"
         Height          =   365
         Index           =   0
         Left            =   10
         TabIndex        =   8
         Top             =   10
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tab 2"
         Height          =   365
         Index           =   1
         Left            =   1450
         TabIndex        =   7
         Top             =   10
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tab 3"
         Height          =   365
         Index           =   2
         Left            =   2900
         TabIndex        =   6
         Top             =   10
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tab 4"
         Height          =   365
         Index           =   3
         Left            =   4350
         TabIndex        =   5
         Top             =   10
         Width           =   1455
      End
   End
End
Attribute VB_Name = "TabbedDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Unload Me
        Case 2
            MsgBox ("Apply command goes here.")
        Case 3
            MsgBox ("Help command goes here.")
    End Select
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width / 2) - (Me.Width / 2), (Screen.Height / 2) - (Me.Height / 2)
End Sub

Private Sub Label1_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To Label1.Count - 1
        Label1(i).BorderStyle = 1
        Box(i).Visible = False
    Next
    Select Case Index
        Case 0
            Label1(0).BorderStyle = 0
            Box(0).Visible = True
        Case 1
            Label1(1).BorderStyle = 0
            Box(1).Visible = True
        Case 2
            Label1(2).BorderStyle = 0
            Box(2).Visible = True
        Case 3
            Label1(3).BorderStyle = 0
            Box(3).Visible = True
    End Select
End Sub
