VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Puzzle Board"
   ClientHeight    =   7530
   ClientLeft      =   8220
   ClientTop       =   5775
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   5775
   Begin VB.Timer tmrScore 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   5760
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   15
      Left            =   3960
      TabIndex        =   18
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   14
      Left            =   2880
      TabIndex        =   17
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   13
      Left            =   1800
      TabIndex        =   16
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   12
      Left            =   720
      TabIndex        =   15
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   11
      Left            =   3960
      TabIndex        =   14
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   10
      Left            =   2880
      TabIndex        =   13
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   9
      Left            =   1800
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   8
      Left            =   720
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   7
      Left            =   3960
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   6
      Left            =   2880
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   1800
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   4
      Left            =   720
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   3960
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   2880
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdTile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Elasped Time:"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.Label lblTimer 
         Alignment       =   2  'Center
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   15
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   14
      Left            =   5280
      Picture         =   "Puzzle_ChenVMain.frx":0000
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   13
      Left            =   5280
      Picture         =   "Puzzle_ChenVMain.frx":327C
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   12
      Left            =   5280
      Picture         =   "Puzzle_ChenVMain.frx":6501
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   11
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":9A08
      Stretch         =   -1  'True
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   10
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":CCE9
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   9
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":FF17
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   8
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":1310B
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   7
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":163BF
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   6
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":192B5
      Stretch         =   -1  'True
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   5
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":1C252
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   4
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":1F1FD
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   3
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":2215E
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   2
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":256FE
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   1
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":28B1A
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgData 
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "Puzzle_ChenVMain.frx":2BF3B
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgBlank 
      Height          =   495
      Left            =   5280
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   15
      Left            =   9720
      Picture         =   "Puzzle_ChenVMain.frx":2F537
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   14
      Left            =   8640
      Picture         =   "Puzzle_ChenVMain.frx":328C8
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   13
      Left            =   7560
      Picture         =   "Puzzle_ChenVMain.frx":35B44
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   12
      Left            =   6480
      Picture         =   "Puzzle_ChenVMain.frx":38DC9
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   11
      Left            =   9720
      Picture         =   "Puzzle_ChenVMain.frx":3C2D0
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   10
      Left            =   8640
      Picture         =   "Puzzle_ChenVMain.frx":3F5B1
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   9
      Left            =   7560
      Picture         =   "Puzzle_ChenVMain.frx":427DF
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   8
      Left            =   6480
      Picture         =   "Puzzle_ChenVMain.frx":459D3
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   7
      Left            =   9720
      Picture         =   "Puzzle_ChenVMain.frx":48C87
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   6
      Left            =   8640
      Picture         =   "Puzzle_ChenVMain.frx":4BB7D
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   5
      Left            =   7560
      Picture         =   "Puzzle_ChenVMain.frx":4EB1A
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   4
      Left            =   6480
      Picture         =   "Puzzle_ChenVMain.frx":51AC5
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   3
      Left            =   9720
      Picture         =   "Puzzle_ChenVMain.frx":54A26
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   2
      Left            =   8640
      Picture         =   "Puzzle_ChenVMain.frx":57FC6
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   1
      Left            =   7560
      Picture         =   "Puzzle_ChenVMain.frx":5B3E2
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgTile 
      Height          =   1095
      Index           =   0
      Left            =   6480
      Picture         =   "Puzzle_ChenVMain.frx":5E803
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblMoves 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   2
      Top             =   6480
      Width           =   4575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuHighscores 
         Caption         =   "Highscores"
         Shortcut        =   ^H
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuNumerical 
         Caption         =   "Numerical Board"
      End
      Begin VB.Menu mnuGraphical 
         Caption         =   "Graphical Board"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name: Vincent Chen
'Date: 2019/06/03
'Purpose: Puzzle Board Version 2

Option Explicit

Dim Start As Single

Private Sub cmdTile_Click(Index As Integer)
    Dim Valid As Boolean
    Dim Win As Boolean
    Dim X As Integer
    
    
    Win = True
    Valid = CheckValid(Index)
    
    If tmrScore.Enabled = False And Valid = True Then
        tmrScore.Enabled = True
        Start = Timer
        lblTimer.Visible = True
    End If
    
    If Valid = True Then
        cmdTile(Blank).Visible = True
        cmdTile(Index).Visible = False
        MoveTileNum Index, cmdTile
        NumMoves = NumMoves + 1
    End If
    
    lblMoves.Caption = "Number of moves made: " + Str$(NumMoves)
    Win = CheckWinNum(cmdTile, Index)

    If Win = True Then
        MsgBox "Victory......You've won?", vbOKOnly, "Win"
        For X = 0 To 15
            cmdTile(X).Enabled = False
        Next X
        tmrScore.Enabled = False
        HighScoreNum
    End If
End Sub

'Name: Vincent Chen
'Date: 2019/06/03

Private Sub Form_Load()
    Dim X As Integer
    
    GameMode = True
    ResetForm cmdTile, imgTile
    NewGameNum cmdTile, lblMoves, tmrScore, lblTimer
    mnuNumerical.Enabled = False
    
    For X = 0 To 15
        imgTile(X).Left = cmdTile(X).Left
        imgTile(X).Top = cmdTile(X).Top
    Next X
End Sub

Private Sub imgTile_Click(Index As Integer)
    Dim Valid As Boolean
    Dim Win As Boolean
    Dim X As Integer
    
    
    Win = True
    Valid = CheckValid(Index)
    
    If tmrScore.Enabled = False And Valid = True Then
        tmrScore.Enabled = True
        Start = Timer
        lblTimer.Visible = True
    End If
    
    If Valid = True Then
        imgTile(Blank).Visible = True
        imgTile(Index).Visible = False
        MoveTileGra Index, imgTile, imgBlank
        NumMoves = NumMoves + 1
    End If
    
    lblMoves.Caption = "Number of moves made: " + Str$(NumMoves)
    Win = CheckWinGra(imgTile, Index)

    If Win = True Then
        MsgBox "Victory......You've won?", vbOKOnly, "Win"
        For X = 0 To 15
            imgTile(X).Enabled = False
        Next X
        tmrScore.Enabled = False
        HighScoreGra
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuExit_Click()
    Dim Msg As Integer
    
    Msg = MsgBox("Are you sure you want to exit?", vbYesNo + vbInformation, "Exit")
    
    If Msg = vbYes Then
        End
    End If
End Sub

Private Sub mnuGraphical_Click()
    Dim Msg As Integer
    
    If tmrScore.Enabled = True Then
        Msg = MsgBox("Are you sure you want to start a new game? There is still an active game running!", vbYesNo + vbExclamation, "New Game")
        If Msg = vbYes Then
            GameMode = False
            ResetForm cmdTile, imgTile
            NewGameGra imgTile, lblMoves, tmrScore, lblTimer, imgBlank, imgData
            mnuGraphical.Enabled = False
            mnuNumerical.Enabled = True
        End If
    Else
        GameMode = False
        ResetForm cmdTile, imgTile
        NewGameGra imgTile, lblMoves, tmrScore, lblTimer, imgBlank, imgData
        mnuGraphical.Enabled = False
        mnuNumerical.Enabled = True
    End If
End Sub

Private Sub mnuHighscores_Click()
    frmScores.Show 1
End Sub

Private Sub mnuNewGame_Click()
    If GameMode = True Then
        NewGameNum cmdTile, lblMoves, tmrScore, lblTimer
    Else
        NewGameGra imgTile, lblMoves, tmrScore, lblTimer, imgBlank, imgData
    End If
End Sub

Private Sub mnuNumerical_Click()
    Dim Msg As Integer
    
    If tmrScore.Enabled = True Then
        Msg = MsgBox("Are you sure you want to start a new game? There is still an active game running!", vbYesNo + vbExclamation, "New Game")
        If Msg = vbYes Then
            GameMode = True
            ResetForm cmdTile, imgTile
            NewGameNum cmdTile, lblMoves, tmrScore, lblTimer
            mnuGraphical.Enabled = True
            mnuNumerical.Enabled = False
        End If
    Else
        GameMode = True
        ResetForm cmdTile, imgTile
        NewGameNum cmdTile, lblMoves, tmrScore, lblTimer
        mnuGraphical.Enabled = True
        mnuNumerical.Enabled = False
    End If
End Sub

Private Sub tmrScore_Timer()
    Dim Current As Single
    Dim Hours As Integer
    Dim Mins As Integer
    Dim Secs As Integer
    
    Current = Timer
    
    Diff = Current - Start
    
    Hours = Int(Diff / 3600)
    Mins = Int((Diff Mod 3600) / 60)
    Secs = (Diff Mod 3600) Mod 60
    lblTimer.Caption = Format$(Hours, "00") & ":" & Format$(Mins, "00") & ":" & Format$(Secs, "00")
End Sub
