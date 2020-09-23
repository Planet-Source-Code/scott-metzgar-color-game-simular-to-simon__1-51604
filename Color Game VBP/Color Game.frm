VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Color Game"
   ClientHeight    =   4395
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5295
   Icon            =   "Color Game.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   125
      Left            =   3960
      Top             =   720
   End
   Begin VB.CommandButton Quit 
      BackColor       =   &H00808080&
      Caption         =   "Give Up"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Start 
      BackColor       =   &H00808080&
      Caption         =   "Start"
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Blue 
      BackColor       =   &H00400000&
      Caption         =   "Blue"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Yellow 
      BackColor       =   &H00008080&
      Caption         =   "Yellow"
      Height          =   975
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Green 
      BackColor       =   &H00004000&
      Caption         =   "Green"
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Red 
      BackColor       =   &H00000080&
      Caption         =   "Red"
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   1920
      ScaleHeight     =   675
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuChangelevel 
         Caption         =   "&Change Level"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Max, Min, C, Mem(255), M15, Play, Game, Pause, Turn, GameTurn As Integer
Public Times As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Form_Load()
Form1.Hide
Dialog1.Show

End Sub

Private Sub Timer1_Timer()

Pause = Pause + 1
If Pause >= 4 And Play = 1 Then Call Gameplay
If Pause >= 8 Then Call Repeat

End Sub

Private Sub Blue_Click()

If Game = 1 Then Call Bluec

End Sub

Private Sub Green_Click()

If Game = 1 Then Call Greenc

End Sub

Private Sub Red_Click()

If Game = 1 Then Call Redc

End Sub


Private Sub Yellow_Click()

If Game = 1 Then Call Yellowc

End Sub


Private Sub Start_Click()

Max = 4
Min = 1
GameTurn = 1
For Turn = 1 To Times
C = (Int((Max - Min + 1) * Rnd + Min))
Mem(Turn) = C
Next Turn
Game = 1
Play = 1
M15 = 0
Turn = 0
Picture1.Picture = LoadPicture("")
Call Repeat

End Sub

Private Sub Quit_Click()

Game = 0
Play = 0
GameTurn = 0
Red.BackColor = &H80&
Yellow.BackColor = &H8080&
Green.BackColor = &H4000&
Blue.BackColor = &H400000
Picture1.BackColor = &H0
Picture1.Picture = LoadPicture("")
Dialog.Show


End Sub

Public Sub Gameplay()

Pause = -4
If Play = 1 Then M15 = M15 + 1
If M15 < 1 Or M15 > Turn Then Play = 0
If Play = 0 Then GoTo 10
C = Mem(M15)
Pause = 0
If C = 1 Then Call Redplay
If C = 2 Then Call Yellowplay
If C = 3 Then Call Greenplay
If C = 4 Then Call Blueplay
Pause = 0
10 If Play = 0 Then M15 = 0

End Sub

Public Sub Redplay()
Pause = -4
SndPlay = sndPlaySound("Redbeep.wav", 1)
Red.BackColor = &HFF&
Yellow.BackColor = &H8080&
Green.BackColor = &H4000&
Blue.BackColor = &H400000
Picture1.BackColor = &HFF&
Pause = 0

End Sub


Public Sub Yellowplay()
Pause = -4
SndPlay = sndPlaySound("Yellowbeep.wav", 1)
Red.BackColor = &H80&
Yellow.BackColor = &HC0C0&
Green.BackColor = &H4000&
Blue.BackColor = &H400000
Picture1.BackColor = &HFFFF&
Pause = 0

End Sub

Public Sub Greenplay()
Pause = -4
SndPlay = sndPlaySound("Greenbeep.wav", 1)
Red.BackColor = &H80&
Yellow.BackColor = &H8080&
Green.BackColor = &H8000&
Blue.BackColor = &H400000
Picture1.BackColor = &H8000&
Pause = 0
End Sub

Public Sub Blueplay()
Pause = -4
SndPlay = sndPlaySound("Bluebeep.wav", 1)
Red.BackColor = &H80&
Yellow.BackColor = &H8080&
Green.BackColor = &H4000&
Blue.BackColor = &HC00000
Picture1.BackColor = &HC00000
Pause = 0

End Sub

Public Sub Bluec()

M15 = M15 + 1
If Mem(M15) <> 4 Then Call Lose
SndPlay = sndPlaySound("Bluebeep.wav", 1)
Red.BackColor = &H80&
Yellow.BackColor = &H8080&
Green.BackColor = &H4000&
Blue.BackColor = &HC00000
Picture1.BackColor = &HC00000
If M15 = Times And Mem(M15) = 4 Then Call Win
Pause = -4


End Sub

Public Sub Greenc()

M15 = M15 + 1
If Mem(M15) <> 3 Then Call Lose
SndPlay = sndPlaySound("Greenbeep.wav", 1)
Red.BackColor = &H80&
Yellow.BackColor = &H8080&
Green.BackColor = &H8000&
Blue.BackColor = &H400000
Picture1.BackColor = &H8000&
If M15 = Times And Mem(M15) = 3 Then Call Win
Pause = -4

End Sub


Public Sub Redc()

M15 = M15 + 1
If Mem(M15) <> 1 Then Call Lose
SndPlay = sndPlaySound("Redbeep.wav", 1)
Red.BackColor = &HFF&
Yellow.BackColor = &H8080&
Green.BackColor = &H4000&
Blue.BackColor = &H400000
Picture1.BackColor = &HFF&
If M15 = Times And Mem(M15) = 1 Then Call Win
Pause = -4

End Sub

Public Sub Yellowc()

M15 = M15 + 1
If Mem(M15) <> 2 Then Call Lose
SndPlay = sndPlaySound("Yellowbeep.wav", 1)
Red.BackColor = &H80&
Yellow.BackColor = &HC0C0&
Green.BackColor = &H4000&
Blue.BackColor = &H400000
Picture1.BackColor = &HFFFF&
If M15 = Times And Mem(M15) = 2 Then Call Win
Pause = -4

End Sub

Public Sub Win()
Play = 0
Game = 0
Picture1.Picture = LoadPicture("smiley.bmp")

End Sub

Public Sub Lose()
Play = 0
Game = 0
GameTurn = 0
Picture1.Picture = LoadPicture("frown.bmp")

End Sub
Private Sub mnuHelpAbout_Click()
    MsgBox "Version 1.0  " & Chr$(169) & " 2003 Scott Metzgar"
    
End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuChangelevel_Click()

Game = 0
Play = 0
Red.BackColor = &H80&
Yellow.BackColor = &H8080&
Green.BackColor = &H4000&
Blue.BackColor = &H400000
Picture1.BackColor = &H0
Picture1.Picture = LoadPicture("")
Dialog1.Show

End Sub

Public Sub Repeat()

Game = 1
Play = 1
M15 = 0
Picture1.Picture = LoadPicture("")
Turn = Turn + 1
If Turn = Times + 1 Then GameTurn = 0
If GameTurn = 0 Then Game = 0
If GameTurn = 0 Then Play = 0
If GameTurn = 1 Then Call Gameplay
End Sub
