Attribute VB_Name = "Module1"
'Name: Vincent Chen
'Date: 2019/06/03

Option Explicit

Type ScoreRec
    Name As String * 15
    Time As Integer
    Moves As Integer
End Type

Global Const FILENAMENUM = "\Numerical.rec"
Global Const FILENAMEGRA = "\Graphical.rec"
Global Const MAX = 15
Global Const TOTALMOVES = 250
Global Const DEFAULTNAME = "TheLegend27"

Global Blank As Integer
Global NumMoves As Integer
Global Current(0 To MAX) As Integer
Global GameMode As Boolean 'True = Numerical, False = Graphical
Global Diff As Integer

Global ScoresNum(1 To 5) As ScoreRec
Global ScoresGra(1 To 5) As ScoreRec

Public Function CheckValid(ByVal Index As Integer) As Boolean
    Dim Valid As Boolean
    
    Select Case Index
        Case (Blank + 4), (Blank - 4)
            Valid = True
        Case (Blank + 1)
            If Index Mod 4 <> 0 Then
                Valid = True
            End If
        Case (Blank - 1)
            If Index Mod 4 <> 3 Then
                Valid = True
            End If
    End Select
    
    CheckValid = Valid
End Function

Public Sub NewGameGra(imgTile As Variant, lblMoves As Label, tmrScore As Timer, lblTimer As Label, imgBlank As Variant, imgData As Variant)
    Dim K As Integer
    Dim Tile As Integer
    Dim Valid As Boolean
    
    Randomize
    tmrScore.Enabled = False
    Blank = 15
    NumMoves = 0
    Diff = 0
    
    lblMoves.Caption = "Number of moves made: " + Str$(NumMoves)
    lblTimer.Caption = "00:00:00"
    
    For K = 0 To 15
        Current(K) = K
    Next K
    
    For K = 0 To 15
        imgTile(K).Enabled = True
        imgTile(K).Picture = imgData(K).Picture
        imgTile(K).Visible = False
    Next K

    For K = 1 To TOTALMOVES
        Do
            Tile = Int(Rnd() * 16)
            Valid = CheckValid(Tile)
        Loop While Valid = False
        MoveTileGra Tile, imgTile, imgBlank
    Next K
    
    For K = 0 To 15
        imgTile(K).Visible = True
    Next K
    
    imgTile(Blank).Visible = False
End Sub

Public Sub MoveTileGra(ByVal Index As Integer, imgTile As Variant, imgBlank As Variant)
    Current(Blank) = Current(Index)
    imgTile(Blank).Picture = imgTile(Index).Picture
    
    Current(Index) = 0
    imgTile(Index).Picture = imgBlank.Picture
    
    Blank = Index
End Sub

Public Function CheckWinGra(imgTile As Variant, ByVal Index As Integer)
    Dim X As Integer
    Dim Win As Integer
    
    Win = 0
    
    For X = 0 To 14
        If Current(X) = X Then
            Win = Win + 1
        End If
    Next X

    If Win = 15 Then
        CheckWinGra = True
    Else
        CheckWinGra = False
    End If
End Function

Public Sub NewGameNum(cmdTile As Variant, lblMoves As Label, tmrScore As Timer, lblTimer As Label)
    Dim K As Integer
    Dim Tile As Integer
    Dim Valid As Boolean
    
    cmdTile(Blank).Visible = False
    Randomize
    tmrScore.Enabled = False
    Blank = 15
    NumMoves = 0
    Diff = 0
    
    lblMoves.Caption = "Number of moves made: " + Str$(NumMoves)
    lblTimer.Caption = "00:00:00"
    
    For K = 0 To 15
        cmdTile(K).Enabled = True
        cmdTile(K).Caption = cmdTile(K).Index + 1
        cmdTile(K).Visible = False
    Next K
    
    For K = 1 To TOTALMOVES
        Do
            Tile = Int(Rnd() * 16)
            Valid = CheckValid(Tile)
        Loop While Valid = False
        MoveTileNum Tile, cmdTile
    Next K
    
    For K = 0 To 15
        cmdTile(K).Visible = True
    Next K
    
    cmdTile(Blank).Visible = False
End Sub

Public Sub MoveTileNum(ByVal Index As Integer, cmdTile As Variant)
    cmdTile(Blank).Caption = cmdTile(Index).Caption
    
    cmdTile(Index).Caption = "0"
    Blank = Index
End Sub

Public Function CheckWinNum(cmdTile As Variant, ByVal Index As Integer)
    Dim X As Integer
    Dim Win As Integer
    
    Win = 0
    
    For X = 0 To 14
        If cmdTile(X).Caption = Trim$(cmdTile(X).Index + 1) Then
            Win = Win + 1
        End If
    Next X

    If Win = 15 Then
        CheckWinNum = True
    Else
        CheckWinNum = False
    End If
End Function

Public Sub ResetForm(cmdTile As Variant, imgTile As Variant)
    Dim K As Integer
    
    For K = 0 To 15
        If GameMode = False Then
            cmdTile(K).Visible = False
        Else
            imgTile(K).Visible = False
        End If
    Next K
End Sub

Public Sub HighScoreNum()
    Dim X As Integer
    Dim Y As Integer
    Dim K As Integer
    Dim Hs As Boolean
    Dim HSName As String
    Dim SwapMoves As Integer
    Dim SwapName As String
    Dim SwapTime As Single
    Dim RecordLen As Integer
    Dim Time As Integer
    Dim Response As Integer
    
    Time = Diff
    RecordLen = Len(ScoresNum(1))
    SwapTime = 0
    SwapMoves = 0
    SwapName = ""
    Y = 0
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
    
    For X = 1 To 5
        If ScoresNum(X).Time = 0 And ScoresNum(X).Moves = 0 Then
            ScoresNum(X).Time = 29999 'To fill up an empty record
            ScoresNum(X).Moves = 29999
            ScoresNum(X).Name = "Empty@#*$@#&$%("
        End If
    Next X

    Do While Y < 5
        Y = Y + 1
        If Time < ScoresNum(Y).Time Then
            Hs = True
            
            X = Y
            Y = 6
        ElseIf Time = ScoresNum(Y).Time Then
            If Y <> 5 Then
                If (NumMoves < ScoresNum(Y).Moves) And (Time <= ScoresNum(Y + 1).Time) Then
                    Hs = True
                    X = Y
                    Y = 6
                End If
            Else
                If (NumMoves < ScoresNum(Y).Moves) Then
                    Hs = True
                    X = Y
                    Y = 6
                End If
            End If
        End If
    Loop
    
    If Hs = True Then
        Do
            HSName = InputBox$("Congratulations! You've made it onto the highscores!" & vbCrLf & vbCrLf & "Tell me your name: ", "High Score!", DEFAULTNAME)
            If HSName = "" Then
                MsgBox "You cannot enter a blank name.", vbOKOnly, "Invalid Name."
            End If
        Loop While HSName = ""
        
        For K = X To 5
            SwapTime = ScoresNum(K).Time
            ScoresNum(K).Time = Time
            Time = SwapTime
            
            SwapName = ScoresNum(K).Name
            ScoresNum(K).Name = HSName
            HSName = SwapName
            
            SwapMoves = ScoresNum(K).Moves
            ScoresNum(K).Moves = NumMoves
            NumMoves = SwapMoves
        Next K
    End If
    
    On Error GoTo ErrorHandler
    
    Kill FILENAMENUM
    
    Open App.Path & FILENAMENUM For Random As #1 Len = RecordLen
    
    For X = 1 To 5
        Put #1, X, ScoresNum(X)
    Next X
    
    Close #1
    
    If Hs = True Then
        frmScores.Show 1
    End If
ErrorHandler:
    Resume Next
End Sub

Public Sub HighScoreGra()
    Dim X As Integer
    Dim Y As Integer
    Dim K As Integer
    Dim Hs As Boolean
    Dim HSName As String
    Dim SwapMoves As Integer
    Dim SwapName As String
    Dim SwapTime As Single
    Dim RecordLen As Integer
    Dim Time As Integer
    Dim Response As Integer
    
    Time = Diff
    RecordLen = Len(ScoresGra(1))
    SwapTime = 0
    SwapMoves = 0
    SwapName = ""
    Y = 0
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
    
    For X = 1 To 5
        If ScoresGra(X).Time = 0 And ScoresGra(X).Moves = 0 Then
            ScoresGra(X).Time = 29999 'To fill up an empty record
            ScoresGra(X).Moves = 29999
            ScoresGra(X).Name = "Empty@#*$@#&$%("
        End If
    Next X
    
    Do While Y < 5
        Y = Y + 1
        If Time < ScoresGra(Y).Time Then
            Hs = True
            
            X = Y
            Y = 6
        ElseIf Time = ScoresGra(Y).Time Then
            If Y <> 5 Then
                If (NumMoves < ScoresGra(Y).Moves) And (Time <= ScoresGra(Y + 1).Time) Then
                    Hs = True
                    X = Y
                    Y = 6
                End If
            Else
                If (NumMoves < ScoresGra(Y).Moves) Then
                    Hs = True
                    X = Y
                    Y = 6
                End If
            End If
        End If
    Loop

    If Hs = True Then
        Do
            HSName = InputBox$("Congratulations! You've made it onto the highscores!" & vbCrLf & vbCrLf & "Tell me your name: ", "High Score!", DEFAULTNAME)
            If HSName = "" Then
                MsgBox "You cannot enter a blank name.", vbOKOnly, "Invalid Name."
            End If
        Loop While HSName = ""
        
        For K = X To 5
            SwapTime = ScoresGra(K).Time
            ScoresGra(K).Time = Time
            Time = SwapTime
            
            SwapName = ScoresGra(K).Name
            ScoresGra(K).Name = HSName
            HSName = SwapName
            
            SwapMoves = ScoresGra(K).Moves
            ScoresGra(K).Moves = NumMoves
            NumMoves = SwapMoves
        Next K
    End If
    
    On Error GoTo ErrorHandler
    
    Kill FILENAMEGRA
    
    Open App.Path & FILENAMEGRA For Random As #1 Len = RecordLen
    
    For X = 1 To 5
        Put #1, X, ScoresGra(X)
    Next X
    
    Close #1
    
    If Hs = True Then
        frmScores.Show 1
    End If
    
ErrorHandler:
    Resume Next
End Sub
