VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmScores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   8655
   ClientLeft      =   2925
   ClientTop       =   4050
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   4
      Top             =   7800
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid grdGra 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   4680
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5318
      _Version        =   393216
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   0
      BorderStyle     =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdNum 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
      _Version        =   393216
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   0
      BorderStyle     =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Graphical Version"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   3960
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Numerical Version"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name: Vincent Chen
'Date: 2019/06/03
'Purpose: Puzzle Board Version 2

Option Explicit

Private Sub cmdReturn_Click()
    Unload frmScores
End Sub

Private Sub Form_Activate()
    Dim X As Integer
    Dim Y As Integer
    Dim Time As Integer
    Dim Fake As Boolean
    Dim RecordLen As Integer
    Dim GFake As Boolean
    Dim GTime As Integer
    
    RecordLen = Len(ScoresNum(1))
    X = 0
    
    Open App.Path & FILENAMENUM For Random As #1 Len = RecordLen
    
    Do While Not EOF(1)
        X = X + 1
        If X > 5 Then
            Exit Do
        End If
        Get #1, X, ScoresNum(X)
    Loop
    
    Close #1
    
    If ScoresNum(1).Time = 0 And ScoresNum(1).Name = "" And ScoresNum(1).Moves = 0 Then
        ScoresNum(1).Time = 29999
        ScoresNum(1).Moves = 29999
        ScoresNum(1).Name = "Empty@#*$@#&$%("
    End If
    
    X = 0
    
    Open App.Path & FILENAMEGRA For Random As #1 Len = RecordLen
    
    Do While Not EOF(1)
        X = X + 1
        If X > 5 Then
            Exit Do
        End If
        Get #1, X, ScoresGra(X)
    Loop
    
    Close #1
    
    If ScoresGra(1).Time = 0 And ScoresGra(1).Name = "" And ScoresGra(1).Moves = 0 Then
        ScoresGra(1).Time = 29999
        ScoresGra(1).Moves = 29999
        ScoresGra(1).Name = "Empty@#*$@#&$%(*"
    End If
    
    For X = 1 To 5
        If ScoresNum(X).Time = 0 And ScoresNum(X).Moves = 0 Then
            ScoresNum(X).Time = 29999 'To fill up an empty record
            ScoresNum(X).Moves = 29999
            ScoresNum(X).Name = "Empty@#*$@#&$%("
        End If
        If ScoresGra(X).Time = 0 And ScoresGra(X).Moves = 0 Then
            ScoresGra(X).Time = 29999 'To fill up an empty record
            ScoresGra(X).Moves = 29999
            ScoresGra(X).Name = "Empty@#*$@#&$%("
        End If
    Next X
    
    For Y = 1 To 5
        grdNum.Row = Y
        grdGra.Row = Y
        
        If ScoresNum(Y).Moves = 29999 And ScoresNum(Y).Time = 29999 And ScoresNum(Y).Name = "Empty@#*$@#&$%(" Then
            Fake = True
        Else
            Fake = False
        End If

        If ScoresGra(Y).Moves = 29999 And ScoresGra(Y).Time = 29999 And ScoresGra(Y).Name = "Empty@#*$@#&$%(" Then
            GFake = True
        Else
            GFake = False
        End If
        
        For X = 1 To 3
            grdNum.Col = X - 1
            grdGra.Col = X - 1
            If X = 1 Then
                If Fake = False Then
                    grdNum.Text = ScoresNum(Y).Name
                Else
                    grdNum.Text = "EMPTY"
                End If
                If GFake = False Then
                    grdGra.Text = ScoresGra(Y).Name
                Else
                    grdGra.Text = "EMPTY"
                End If
            ElseIf X = 2 Then
                If Fake = False Then
                    Time = ScoresNum(Y).Time
                    grdNum.Text = Format$(Int((Time / 3600)), "00") & ":" & Format$(Int(((Time Mod 3600) / 60)), "00") & ":" & Format$(((Time Mod 60)), "00")
                Else
                    grdNum.Text = "EMPTY"
                End If
                If GFake = False Then
                    GTime = ScoresGra(Y).Time
                    grdGra.Text = Format$(Int((GTime / 3600)), "00") & ":" & Format$(Int(((GTime Mod 3600) / 60)), "00") & ":" & Format$(((GTime Mod 60)), "00")
                Else
                    grdGra.Text = "EMPTY"
                End If
            ElseIf X = 3 Then
                If Fake = False Then
                    grdNum.Text = ScoresNum(Y).Moves
                Else
                    grdNum.Text = "EMPTY"
                End If
                If GFake = False Then
                    grdGra.Text = ScoresGra(Y).Moves
                Else
                    grdGra.Text = "EMPTY"
                End If
            End If
        Next X
    Next Y
End Sub

Private Sub Form_Load()
    Dim C As Integer, R As Integer, W As Integer
    
    With grdNum
        .Rows = 6
        .Cols = 3
        .FixedCols = 0
        
        .ColWidth(0) = 3000
        For C = 1 To .Cols - 1
            .ColWidth(C) = 1000
        Next C
        
        For R = 1 To .Rows - 1
            .RowHeight(R) = 500
        Next R
        
        .Row = 0
        .Col = 0
        .Text = "Name"
        .Col = 1
        .Text = "Time"
        .Col = 2
        .Text = "Moves"
    End With
    
    With grdGra
        .Rows = 6
        .Cols = 3
        .FixedCols = 0
        
        .ColWidth(0) = 3000
        For C = 1 To .Cols - 1
            .ColWidth(C) = 1000
        Next C
        
        For R = 1 To .Rows - 1
            .RowHeight(R) = 500
        Next R
        .Row = 0
        .Col = 0
        .Text = "Name"
        .Col = 1
        .Text = "Time"
        .Col = 2
        .Text = "Moves"
    End With
End Sub
