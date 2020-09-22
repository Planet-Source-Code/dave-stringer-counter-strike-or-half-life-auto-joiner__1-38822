VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Game Thingy"
   ClientHeight    =   4770
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6090
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5520
      TabIndex        =   28
      Text            =   "0"
      Top             =   3960
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   3015
      Begin VB.TextBox txtSERVEROS 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtDEDICATEDSERVER 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtPLAYERS 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   720
         TabIndex        =   25
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtGAME 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtMAP 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtHOST 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtSERVERNAME 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblSERVEROS 
         Alignment       =   1  'Right Justify
         Caption         =   "Server OS:"
         Height          =   255
         Left            =   65
         TabIndex        =   24
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblDEDICATEDSERVER 
         Alignment       =   1  'Right Justify
         Caption         =   "Dedicated Server:"
         Height          =   255
         Left            =   90
         TabIndex        =   23
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblPLAYERS 
         Alignment       =   1  'Right Justify
         Caption         =   "Players:"
         Height          =   255
         Left            =   80
         TabIndex        =   22
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblGAME 
         Alignment       =   1  'Right Justify
         Caption         =   "Game:"
         Height          =   255
         Left            =   90
         TabIndex        =   21
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblMAP 
         Alignment       =   1  'Right Justify
         Caption         =   "Map:"
         Height          =   255
         Left            =   105
         TabIndex        =   20
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblSERVERNAME 
         Alignment       =   1  'Right Justify
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblHOST 
         Alignment       =   1  'Right Justify
         Caption         =   "Host:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   375
      End
   End
   Begin MSWinsockLib.Winsock sckUDP2 
      Left            =   4080
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Text            =   "3"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Text            =   "27015"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Refresh Rate In Seconds:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Server IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto Join"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   4320
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   4440
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3495
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "player"
         Text            =   "Player"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "kills"
         Text            =   "Kills"
         Object.Width           =   794
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Query Server"
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   4320
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock sckUDP1 
      Left            =   3480
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "http://www.c6d.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "How Many Reserved Slots:"
      Height          =   255
      Left            =   3480
      TabIndex        =   26
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lblPLAYERSTOTAL 
      Caption         =   "Active Players"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuSETTINGS 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEXIT 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuFAVORITES 
      Caption         =   "Favorites"
      Begin VB.Menu mnuADD 
         Caption         =   "Add Current Server"
      End
      Begin VB.Menu mnuREMOVE 
         Caption         =   "Remove a Server"
      End
      Begin VB.Menu mnufav 
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error Resume Next
'handle refresh
If Text3 <> "" Then
 If CInt(Text3) > 0 And CInt(Text3) < 61 Then
  Timer1.Enabled = True
  Timer1.Interval = Text3 & "000"
 Else
  Timer1.Interval = 0
  Timer1.Enabled = False
 End If
End If

'send the players query
sckUDP1.RemoteHost = Text1
sckUDP1.RemotePort = Text2
sckUDP1.SendData Chr(255) & Chr(255) & Chr(255) & Chr(255) & "players"

'send the details query
sckUDP2.RemoteHost = Text1
sckUDP2.RemotePort = Text2
sckUDP2.SendData Chr(255) & Chr(255) & Chr(255) & Chr(255) & "details"

End Sub

Public Sub Form_Load()
'Load Form3
'Form3.Show
Call Load_Favorites
On Error Resume Next
Open "settings.txt" For Input As #1
Line Input #1, temp$
Game_Path$ = temp$
Close #1
If Err.Number > 0 Then Call mnuSETTINGS_Click
If Len(Game_Path$) < 5 Then Call mnuSETTINGS_Click
Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Timer1.Interval = 0
Timer1.Enabled = False
Unload Form2
Unload Form3
End
End Sub

Private Sub Label5_Click()
On Error Resume Next
Shell "explorer http://www.c6d.net"
End Sub

Private Sub mnuADD_Click()
If txtSERVERNAME = "" Then
 MsgBox "There must be a servername! :)" & vbCrLf & "Query the server and then try adding it!"
 Exit Sub
End If

Open "favorites.txt" For Append As #1
Print #1, txtSERVERNAME
Print #1, txtHOST
Close #1

Call ReLoad_Favorites
End Sub

Private Sub mnuEXIT_Click()
Unload Form2
Unload Me
End Sub

Private Sub mnuFAV_Click(Index As Integer)
temp$ = Mid(mnufav(Index).Caption, InStrRev(mnufav(Index).Caption, Chr(30)) + 1)
Text1 = Left(temp$, InStr(1, temp$, ":") - 1)
Text2 = Mid(temp$, InStr(1, temp$, ":") + 1)
DoEvents
Call Command1_Click
End Sub

Private Sub mnuREMOVE_Click()

Load Form3
Form3.Show

Open "favorites.txt" For Input As #3

Do Until EOF(3)
 Line Input #3, tmp1$
 Line Input #3, tmp2$
 temp$ = tmp1$ & Chr(30) & " - " & Chr(30) & tmp2$
 Form3.LST1.AddItem temp$
Loop

Close #3

End Sub

Private Sub mnuSETTINGS_Click()
Load Form2
Form2.Show
End Sub

Private Sub sckUDP1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
sckUDP1.GetData temp$

If Left(temp$, 5) = Chr(255) & Chr(255) & Chr(255) & Chr(255) & Chr(68) Then

'Text10 = ""
'For a = 1 To Len(temp$)

'Text10 = Text10 & Chr(Asc(Mid(temp$, a, 1)))
'Next a


'**********players********

Dim CurrentPlayers

'seperate beginning crap
temp$ = Mid(temp$, 6)

'get current number of players
CurrentPlayers = Asc(Left(temp$, 1))

'strip off current players
temp$ = Mid(temp$, 2)

lblPLAYERSTOTAL.Caption = CurrentPlayers & " Active Players"

LV1.ListItems.Clear

'take care of first player

'calc time
 'ptmp$ = Right(temp$, 4)
 'For b = 1 To 4
  'ptime$ = ptime$ & Asc(Mid(ptmp$, b, 1)) & " "
 'Next b
temp$ = Left(temp$, Len(temp$) - 4) 'chop time
pkills$ = Asc(Mid(temp$, Len(temp$) - 3, 1))
temp$ = Left(temp$, Len(temp$) - 5) 'chop kills
player$ = Mid(temp$, InStrRev(temp$, Chr(CurrentPlayers)))
pnum$ = Asc(Left(player$, 1))
temp$ = Left(temp$, InStrRev(temp$, Chr(CurrentPlayers)) - 1)
'LV1.ListItems.Add , , pnum$ & ". " & Mid(player$, 2)
LV1.ListItems.Add , , Mid(player$, 2)
LV1.ListItems(1).SubItems(1) = pkills$
'LV1.ListItems(1).SubItems(2) = ptime$
'ptime$ = ""


'take of the rest
For a = 1 To CurrentPlayers
 If a = CurrentPlayers Then Exit For
 'calc time
  'ptmp$ = Right(temp$, 4)
  'For b = 1 To 4
   'ptime$ = ptime$ & Asc(Mid(ptmp$, b, 1))
  'Next b
 temp$ = Left(temp$, Len(temp$) - 4)
 pkills$ = Asc(Mid(temp$, Len(temp$) - 3, 1))
 temp$ = Left(temp$, Len(temp$) - 5)
 player$ = Mid(temp$, InStrRev(temp$, Chr(CurrentPlayers - a)))
 pnum$ = Asc(Left(player$, 1))
 temp$ = Left(temp$, InStrRev(temp$, Chr(CurrentPlayers - a)) - 1)
 'LV1.ListItems.Add , , pnum$ & ". " & Mid(player$, 2)
 LV1.ListItems.Add , , Mid(player$, 2)
 LV1.ListItems(a + 1).SubItems(1) = pkills$
 'LV1.ListItems(a + 1).SubItems(2) = ptime$
 'ptime$ = ""
Next a
End If
Err.Clear
End Sub
Public Sub sckUDP2_DataArrival(ByVal bytesTotal As Long)
sckUDP2.GetData temp$

'Text10 = ""

If Left(temp$, 5) = Chr(255) & Chr(255) & Chr(255) & Chr(255) & Chr(109) Then
 
'For a = 1 To Len(temp$)
 'Text10 = Text10 & Chr(Asc(Mid(temp$, a, 1)))
'Next a

'strip first crap
 temp$ = Mid(temp$, 6)

txtHOST.Text = Left(temp$, InStr(1, temp$, Chr(0)) - 1)
temp$ = Mid(temp$, InStr(1, temp$, Chr(0)) + 1)
txtSERVERNAME.Text = Left(temp$, InStr(1, temp$, Chr(0)) - 1)
temp$ = Mid(temp$, InStr(1, temp$, Chr(0)) + 1)
txtMAP.Text = Left(temp$, InStr(1, temp$, Chr(0)) - 1)
temp$ = Mid(temp$, InStr(1, temp$, Chr(0)) + 1)
'txtGAMEDIRECTORY.text = "Game Directory: " & Left(temp$, InStr(1, temp$, Chr(0)) - 1)
Game_Dir$ = Left(temp$, InStr(1, temp$, Chr(0)) - 1)
temp$ = Mid(temp$, InStr(1, temp$, Chr(0)) + 1)
txtGAME.Text = Left(temp$, InStr(1, temp$, Chr(0)) - 1)
temp$ = Mid(temp$, InStr(1, temp$, Chr(0)) + 1)
txtPLAYERS.Text = Asc(Left(temp$, 1)) & "/" & Asc(Mid(temp$, 2, 1))
'AUTO-JOIN
If (Asc(Left(temp$, 1)) + Text4) < Asc(Mid(temp$, 2, 1)) And Check1.Value = 1 Then
 Timer1.Interval = 0
 Timer1.Enabled = False
 Shell Game_Path$ & " -console -game " & Game_Dir$ & " +connect " & Text1.Text & ":" & Text2.Text
 DoEvents
 Unload Me
End If
'END AUTO-JOIN
temp$ = Mid(temp$, 3)
'txtPROTOCOL.text = "Protocol: " & Left(temp$, InStr(1, temp$, Chr(0)) - 1)
temp$ = Mid(temp$, 2)
If LCase(Left(temp$, 1)) = "l" Then txtDEDICATEDSERVER.Text = "No"
If LCase(Left(temp$, 1)) = "d" Then txtDEDICATEDSERVER.Text = "Yes"
temp$ = Mid(temp$, 2)
If LCase(Left(temp$, 1)) = "l" Then txtSERVEROS.Text = "Linux"
If LCase(Left(temp$, 1)) = "w" Then txtSERVEROS.Text = "Windows"
temp$ = Mid(temp$, 2)
End If
'temp$ = Mid(temp$, InStr(1, temp$, Chr(0)) + 1)
'MsgBox InStr(1, temp$, Chr(0))
End Sub

Private Sub Text3_Change()
If Text3 = "" Then Exit Sub
If CInt(Text3) > 0 And CInt(Text3) < 61 Then
 Timer1.Enabled = True
 Timer1.Interval = Text3 & "000"
Else
 Timer1.Interval = 0
 Timer1.Enabled = False
End If
End Sub

Private Sub Text4_Change()
Check1.Value = 0
End Sub

Private Sub Timer1_Timer()
Call Command1_Click
End Sub

Private Sub oldcode()
player$ = Mid(temp$, InStr(temp$, Chr(a)))
 pnum$ = Asc(Left(player$, 1))
 player$ = Mid(player$, 2)
 player$ = Left(player$, InStr(1, player$, Chr(a + 1)))
 temp$ = Mid(temp$, InStr(temp$, Chr(a)))
 ptime$ = Right(player$, 4)
  
  'time on map
  For b = 1 To 3
   ptmp$ = ptmp$ & Asc(Left(ptime$, 1))
   ptime$ = Mid(ptime$, 2)
  Next b
  ptmp$ = ptmp$ & Asc(ptime$)
  ptime$ = ptmp$
  ptmp$ = ""
  
  player$ = Left(player$, Len(player$) - 4)
  
  'kills
  pkills$ = Asc(Right(player$, 1))
  player$ = Left(player$, Len(player$) - 1)
  
 LV1.ListItems.Add , , pnum$ & ". " & player$
 'LV1.ListItems.Add , , "hi"
 LV1.ListItems(a).SubItems(1) = pkills$
 LV1.ListItems(a).SubItems(2) = ptime$
 
 DoEvents


'add last player
 'time on map
 ptime$ = Right(temp$, 4)
  For b = 1 To 3
   ptmp$ = ptmp$ & Asc(Left(ptime$, 1))
   ptime$ = Mid(ptime$, 2)
  Next b
  ptmp$ = ptmp$ & Asc(ptime$)
  ptime$ = ptmp$
  ptmp$ = ""
temp$ = Mid(temp$, InStr(temp$, Chr(a)))
pnum$ = Asc(Left(temp$, 1))
LV1.ListItems.Add , , pnum$ & ". " & temp$
LV1.ListItems(a).SubItems(1) = "kills"
LV1.ListItems(a).SubItems(2) = ptime$
End Sub

Public Sub Load_Favorites()
On Error Resume Next
Open "favorites.txt" For Input As #2

mCount = 0
Do Until EOF(2)
 mCount = mCount + 1
 Line Input #2, tmp1$
 Line Input #2, tmp2$
 temp$ = tmp1$ & Chr(30) & " - " & Chr(30) & tmp2$
 Load mnufav(mCount)
 mnufav(mCount).Caption = temp$
 If Err.Number > 0 Then Exit Do
Loop

Close #2

End Sub

Public Sub ReLoad_Favorites()
For a = 1 To mnufav.Count - 1
Unload mnufav(a)
Next a
Call Load_Favorites
End Sub
