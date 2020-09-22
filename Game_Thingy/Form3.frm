VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Thingy"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4605
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.ListBox LST1 
      Height          =   1230
      ItemData        =   "Form3.frx":030A
      Left            =   120
      List            =   "Form3.frx":030C
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Select a server to remove:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
LST1.RemoveItem (LST1.ListIndex)
DoEvents

Open "favorites.txt" For Output As #4

For a = 0 To LST1.ListCount - 1
tmp1$ = Left(LST1.List(a), InStr(1, LST1.List(a), Chr(30)) - 1)
tmp2$ = Mid(LST1.List(a), InStr(1, LST1.List(a), Chr(30)) + 5)
Print #4, tmp1$
Print #4, tmp2$
Next a

Close #4

Call Form1.ReLoad_Favorites
End Sub

Private Sub Form_Load()
Call putMeOnTop(Me)
End Sub
