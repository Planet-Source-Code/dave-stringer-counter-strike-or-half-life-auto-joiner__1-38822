VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Thingy Settings"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2880
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Where is hl.exe or cstrike.exe?"
      Filter          =   "*.exe"
      InitDir         =   "c:\"
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Path to hl.exe or cstrike.exe:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
CD1.Filter = "Game exe's (hl.exe;cstrike.exe)|hl.exe;cstrike.exe"
CD1.ShowOpen
Text1.Text = CD1.FileName
End Sub

Private Sub Command2_Click()
Game_Path$ = Text1
Open App.Path & "\settings.txt" For Output As #1
Print #1, Game_Path$
Close #1
End Sub

Private Sub Form_Load()
On Error Resume Next
Open "settings.txt" For Input As #1
Line Input #1, temp$
Text1.Text = temp$
Close #1
Err.Clear
Call putMeOnTop(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Game_Path$ = Text1.Text
DoEvents
End Sub
