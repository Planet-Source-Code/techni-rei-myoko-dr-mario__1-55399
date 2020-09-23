VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DrMarioCLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Event WonGame()
Public Event LostGame()
Public Event TilesCleared(Number As Long)
Public Event VirusChanged(Virii As Long)
Public Event PulledTile()
Public Event VirusKilled(Number As Long)

Private Enum Angles
    NoAngle = 0
    leftside
    rightside
    bottomside
    topside
End Enum

Public Enum DrMAction
    movedown
    MoveLeft
    moveright
    RotateTiles
    MoveTiles
End Enum

Private Type Tile
    color As Long
    y As Long
    Clean As Boolean
    Z As Long
    Viral As Boolean
    Bound As Boolean
    Angle As Angles
    bX As Long
    bY As Long
End Type

Private Type MiniTile
    Color1 As Long
    Color2 As Long
    x As Long
    y As Long
    YOffset As Long
    Angle As Long
End Type

Private Grid() As Tile, VirusCounts() As Long, Tiles As Long, GridWidth As Long, GridHeight As Long, TileSize As Long, Colors As Long, Ghost As Long, Clean As Boolean, BoardHeight As Long, Viral As Long, Locked As Boolean, Gameisover As Boolean, Total As Long, OldVirii As Long
Private Source As PictureBox, Destination As PictureBox, NextBlock As PictureBox
Private CurrentTile As MiniTile, FutureTile As MiniTile

'DR MARIO CODE
Private Sub TransBLT(SrcHDC As Long, xSrc As Long, ySrc As Long, MaskHDC As Long, ByVal Xmsk As Long, ByVal Ymsk As Long, Width As Long, Height As Long, DestHDC As Long, x As Long, y As Long, Optional UseSrcAsMask As Boolean)
    Const SRCPAINT = &HEE0086
    If UseSrcAsMask Then Xmsk = xSrc: Ymsk = ySrc
    BitBlt DestHDC, x, y, Width, Height, MaskHDC, Xmsk, Ymsk, SRCPAINT
    BitBlt DestHDC, x, y, Width, Height, SrcHDC, xSrc, ySrc, vbSrcAnd
End Sub

Public Sub KeyDown(KeyCode As Integer)
    Select Case KeyCode
        Case 37, 100: DoAction MoveLeft
        Case 39, 102: DoAction moveright
        Case 40, 98: DoAction movedown
        Case 13, 32, 101, 9, 20, 33, 34, 96, 107, 38, 104: DoAction RotateTiles
    End Select
End Sub

Public Function TileCount() As Long
    TileCount = Tiles
End Function

Private Sub doGameOver()
    Gameisover = True
    RaiseEvent LostGame
End Sub

Public Sub DoAction(Index As DrMAction)
    If Gameisover Then Exit Sub
    On Error Resume Next
    With CurrentTile
        Select Case Index
            Case MoveLeft
                Select Case .Angle
                    Case leftside, rightside: If .x > 1 Then If Grid(.y, .x - 1).color = 0 Then .x = .x - 1
                    Case topside: If .x > 1 Then If Grid(.y, .x - 1).color = 0 And Grid(.y - 1, .x - 1).color = 0 Then .x = .x - 1
                    Case bottomside: If .x > 1 Then If Grid(.y, .x - 1).color = 0 And Grid(.y + 1, .x - 1).color = 0 Then .x = .x - 1
                End Select
            Case moveright
                Select Case .Angle
                    Case leftside: If .x < GridWidth - 1 Then If Grid(.y, .x + 2).color = 0 Then .x = .x + 1
                    Case rightside: If .x < GridWidth Then If Grid(.y, .x + 1).color = 0 Then .x = .x + 1
                    Case bottomside: If .x < GridWidth Then If Grid(.y, .x + 1).color = 0 And Grid(.y + 1, .x + 1).color = 0 Then .x = .x + 1
                    Case topside: If .x < GridWidth Then If Grid(.y, .x + 1).color = 0 And Grid(.y - 1, .x + 1).color = 0 Then .x = .x + 1
                End Select
            Case movedown
                CurrentTile.YOffset = 1
                DecrementCurrentTile
            Case RotateTiles: RotateTile
        End Select
    End With
End Sub

Private Sub RotateTile()
    Dim temp As Long
    With CurrentTile
        Select Case .Angle
            Case leftside
                If .y < GridHeight Then
                    If Grid(.y + 1, .x).color = 0 Then .Angle = bottomside
                Else
                    .Angle = bottomside
                End If
            Case topside: If .x < GridWidth Then If Grid(.y, .x + 1).color = 0 Then .Angle = leftside
            Case rightside: If .y > 1 Then If Grid(.y - 1, .x).color = 0 Then .Angle = topside
            Case bottomside ': If Grid(.Y, .X - 1).Color = 0 Then .Angle = rightside
                If .y > 0 Then '< GridHeight Then
                    If .x = GridWidth Then If Grid(.y, .x - 1).color = 0 Then .x = .x - 1
                    If .y < GridHeight Then
                        If .x > 1 Then If Grid(.y, .x - 1).color = 0 And Grid(.y, .x + 1).color > 0 Then .x = .x - 1
                        If Grid(.y + 1, .x).color = 0 And .x < GridWidth Then
                            If Grid(.y, .x + 1).color = 0 Then
                                temp = .Color1
                                .Color1 = .Color2
                                .Color2 = temp
                                .Angle = leftside
                            End If
                        End If
                    End If
                End If
        End Select
    End With
End Sub

Private Sub RandomFutureTile()
    With FutureTile
        .Color1 = RandomTile
        .Color2 = RandomTile
        NextBlock.Cls
        DrawExternal NextBlock, .Color1, .Color2
        NextBlock.Refresh
    End With
End Sub

Public Sub PullTile()
    If Tiles <= 0 Then Exit Sub
    CurrentTile = FutureTile
    RandomFutureTile
    With CurrentTile
        .Angle = 1
        .x = (GridWidth / 2)
        .y = GridHeight
        .YOffset = TileSize
        If Grid(.y, .x).color <> 0 Or Grid(.y, .x + 1).color Then doGameOver
    End With
    Locked = False
    RaiseEvent PulledTile
    OldVirii = Total
End Sub
    
Public Function ViralCount(color As Long) As Long
    ViralCount = VirusCounts(color)
End Function

Public Sub SwitchViral()
    Viral = 1 - Viral
    Clean = False
End Sub

Public Sub InitDrMario(src As PictureBox, Dest As PictureBox, Future As PictureBox, Optional TileWidth As Long = 16, Optional Width As Long = 8, Optional Height As Long = 16, Optional ColorCount = 3, Optional Remains As Long = 8)
    Set Source = src
    Set Destination = Dest
    Set NextBlock = Future
    TileSize = TileWidth
    ReDim Grid(1 To Height, 1 To Width)
    ReDim VirusCounts(1 To ColorCount)
    Total = 0
    GridWidth = Width
    Colors = ColorCount
    GridHeight = Height
    Destination.Width = GridWidth * TileSize
    Destination.Height = GridHeight * TileSize
    BoardHeight = Destination.Height
    Ghost = Remains
End Sub

Public Sub RandomizeBoard(ByVal Count As Long, ByVal Level As Long)
    ClearBoard
    Do Until Count = 0
        RandomizeViral Level
        DrawScreen
        Count = Count - 1
        RaiseEvent VirusChanged(Total)
    Loop
    RandomFutureTile
    Gameisover = False
End Sub

Private Sub RandomizeViral(Optional ByVal Level As Long)
    Dim x As Long, y As Long, color As Long, temp As Boolean
    If Level = 0 Then Level = GridHeight
    Do Until temp 'Grid(Y, X).Color = 0
        Randomize Timer
        x = Rnd * (GridWidth - 1) + 1
        y = Rnd * (Level - 1) + 1
        color = RandomTile
        If Grid(y, x).color = 0 Then temp = Not IsAScore(x, y, color)
    Loop
    ClearTile x, y, True, color
    Total = Total + 1
    VirusCounts(color) = VirusCounts(color) + 1
    Tiles = Tiles + 1
    Grid(y, x).Viral = True
End Sub

'Draws a blank square
Private Sub DrawSquare(x As Long, y As Long)
    Destination.Line (x, y)-(x + TileSize - 1, y + TileSize - 1), Destination.BackColor, B
End Sub

Public Sub ClearBoard()
    Dim temp As Long, temp2 As Long
    Tiles = 0
    With CurrentTile
        .Color1 = 0
        .Color2 = 0
    End With
    RandomFutureTile
    For temp = 1 To Colors
        VirusCounts(temp) = 0
    Next
    Total = 0
    RaiseEvent VirusChanged(Total)
    For temp = 1 To GridWidth
        For temp2 = 1 To GridHeight
            ClearTile temp, temp2
        Next
    Next
End Sub

Private Sub DrawTile(x As Long, y As Long, Optional color As Long, Optional Angle As Long = 0)
    Dim SrcX As Long
    SrcX = Angle * TileSize
    TransBLT Source.hDC, SrcX, color * TileSize, Source.hDC, SrcX, 0, TileSize, TileSize, Destination.hDC, x, y
End Sub

'Chooses a random tile color (0 is the cursor, 1 is the placeholder)
Public Function RandomTile() As Long
    Randomize Timer
    RandomTile = (Rnd * (Colors - 1)) + 1
End Function

Private Function ClearTile(ByVal x As Long, ByVal y As Long, Optional Setit As Boolean = False, Optional color As Long) As Long
    If color = 0 And Setit Then color = RandomTile
    If x < 1 Or x > GridWidth Or y < 1 Or y > GridHeight Then Exit Function
    With Grid(y, x)
        .color = color
        .y = 0
        .Z = 0
        .Bound = False
        .Viral = False
        .bX = 0
        .bY = 0
        .Angle = 0
        .Clean = False
    End With
    ClearTile = color
End Function

Private Function CanMoveTileDown(x As Long, y As Long) As Boolean
    Dim temp As Boolean
    If y > 1 Then
        With Grid(y, x)
          If Not .Viral And .Z = 0 Then
            temp = Grid(y - 1, x).color = 0
            If .Bound Then
                If .Angle = leftside Then temp = temp And Grid(y - 1, x + 1).color = 0
                If .Angle = rightside Then temp = temp And Grid(y - 1, x - 1).color = 0
            End If
            CanMoveTileDown = temp
          End If
        End With
    End If
End Function

Private Function MoveCurrentTileDown() As Boolean
    Dim temp As Boolean
    With CurrentTile
        If .y > 1 Then
            temp = Grid(.y - 1, .x).color = 0
            If .Angle = leftside Then temp = temp And Grid(.y - 1, .x + 1).color = 0
            If .Angle = rightside Then temp = temp And Grid(.y - 1, .x - 1).color = 0
        End If
    End With
    MoveCurrentTileDown = temp
End Function

Private Function MoveTileDown(x As Long, y As Long) As Boolean
    If y = 1 Or x = 0 Then Exit Function
    If Grid(y - 1, x).color > 0 Or Grid(y, x).y > 0 Then Exit Function
    Grid(y - 1, x) = Grid(y, x)
    ClearTile x, y
    If x > 0 Then
    With Grid(y - 1, x)
        .y = TileSize
        .Clean = False
        If .Bound Then
            .bY = .bY - 1
            If Grid(.bY, .bX).color = 0 Then
                .Bound = False
                .Angle = NoAngle
            End If
        End If
    End With
    MoveTileDown = True
    End If
End Function

Private Function MoveTilesDown(ByVal x As Long, ByVal y As Long) As Boolean
    If CanMoveTileDown(x, y) Then
        With Grid(y, x)
            If .Bound Then
                Select Case .Angle
                    Case leftside: MoveTileDown x + 1, y
                    Case rightside: MoveTileDown x - 1, y
                    Case topside: MoveTileDown x, y - 1
                    Case bottomside
                        MoveTileDown x, y
                        MoveTileDown x, x + 1
                        Exit Function
                End Select
            End If
        End With
        MoveTileDown x, y
        MoveTilesDown = True
    End If
End Function

Public Sub DrawScreen()
    Dim temp As Long, temp2 As Long
    If Not Clean Then Destination.Cls
    For temp = 1 To GridWidth
        For temp2 = 1 To GridHeight
            DrawGrid temp, temp2
        Next
    Next
    DrawCurrentTile
End Sub

Private Function GetAngle(Z As Long, isViral As Boolean, Bound As Boolean, Angle As Long) As Long
    If Z > 0 Then
        GetAngle = 2
    Else
        If isViral Then
            GetAngle = Viral
        Else
            If Bound Then
                GetAngle = Angle + 3
            Else
                GetAngle = 3
            End If
        End If
    End If
End Function

Private Sub DrawGrid(x As Long, y As Long)
    Dim Angle As Long
    With Grid(y, x)
        If .Clean = False Or Clean = False Then
            If .y <> 0 And .color = 0 Then Exit Sub
            If .color > 0 Then
                DrawTile (x - 1) * TileSize, BoardHeight - (y * TileSize) - .y, .color, GetAngle(.Z, .Viral, .Bound, .Angle)
            'Else
                'DrawSquare (x - 1) * TileSize, BoardHeight - (y * TileSize) - .y
            End If
            .Clean = True
        End If
    End With
End Sub

Public Sub DrawCurrentTile()
   With CurrentTile
        If .Color1 > 0 And .Color2 > 0 Then
        DrawTile (.x - 1) * TileSize, BoardHeight - (.y * TileSize) - .YOffset, .Color1, GetAngle(0, False, True, .Angle)
        Select Case .Angle
            Case leftside: DrawTile (.x) * TileSize, BoardHeight - (.y * TileSize) - .YOffset, .Color2, rightside + 3
            Case topside: DrawTile (.x - 1) * TileSize, BoardHeight - ((.y - 1) * TileSize) - .YOffset, .Color2, bottomside + 3
            Case bottomside: DrawTile (.x - 1) * TileSize, BoardHeight - ((.y + 1) * TileSize) - .YOffset, .Color2, topside + 3
            Case rightside: DrawTile (.x + 1) * TileSize, BoardHeight - (.y * TileSize) - .YOffset, .Color2, leftside + 3
        End Select
        End If
   End With
End Sub

Private Function TransferTile() As Boolean
    Dim temp As Boolean
    With CurrentTile
         ClearTile .x, .y, True, .Color1
         temp = CheckTile(.x, .y) > 0
         Grid(.y, .x).Angle = .Angle
         If Tiles = 0 Then Exit Function
         Select Case .Angle
            Case leftside: temp = temp Or Bind(.x, .y, .x + 1, .y, rightside, .Color2)
            Case topside: temp = temp Or Bind(.x, .y, .x, .y - 1, bottomside, .Color2)
            Case bottomside: temp = temp Or Bind(.x, .y, .x, .y + 1, topside, .Color2)
            Case rightside: temp = temp Or Bind(.x, .y, .x - 1, .y, leftside, .Color2)
        End Select
        .Color1 = 0
        .Color2 = 0
    End With
    TransferTile = temp
End Function

Private Sub DropAll()
    Dim temp As Long
    For temp = 1 To GridWidth
        DropColumn temp, 1
    Next
End Sub

Private Function Bind(x As Long, y As Long, X2 As Long, Y2 As Long, Angle As Angles, color As Long) As Long
    'MsgBox x & ", " & y & " is bound to " & X2 & ", " & Y2
    With Grid(y, x)
        .Bound = True
        .bX = X2
        .bY = Y2
    End With
    ClearTile X2, Y2, True, color
    With Grid(Y2, X2)
        .Bound = True
        .bX = x
        .bY = y
        .Angle = Angle
    End With
    Bind = CheckTile(X2, Y2)
End Function

Private Sub DrawExternal(Destination As PictureBox, Color1 As Long, Color2 As Long)
    TransBLT Source.hDC, 4 * TileSize, Color1 * TileSize, Source.hDC, 4 * TileSize, 0, TileSize, TileSize, NextBlock.hDC, 0, 0
    TransBLT Source.hDC, 5 * TileSize, Color2 * TileSize, Source.hDC, 5 * TileSize, 0, TileSize, TileSize, NextBlock.hDC, TileSize, 0
End Sub

Public Sub CreateTile(x As Long, y As Long, Optional color As Long)
    If color = 0 Then color = RandomTile
    ClearTile x, y, True, color
    MoveTilesDown x, y
End Sub

Public Function DecrementOffsets(Optional Pixels As Long = 1) As Boolean
    If Gameisover Then Exit Function
    Dim temp As Long, temp2 As Long, waslocked As Boolean
    waslocked = Locked
    Locked = False
    For temp = 1 To GridWidth
        For temp2 = 1 To GridHeight
            With Grid(temp2, temp)
                If .y <> 0 Or .Z <> 0 Then
                    DecrementOffsets = True
                    Locked = True
                    'Falling Tile
                    If .y > 0 Then
                        .y = .y - Pixels
                        If .y = 0 Then If Not MoveTilesDown(temp, temp2) Then CheckTile temp, temp2
                    End If
                    'Disappearing Tile
                    If .Z > 0 Then
                        .Z = .Z - 1
                        If .Z = 0 Then
                            'Decrease viral count by 1, and check for win game status
                            'Moved here so the user gets to see some animation after he beats the level
                            If .Viral And .color > 0 Then
                                VirusCounts(.color) = VirusCounts(.color) - 1
                                Total = Total - 1
                                RaiseEvent VirusChanged(Total)
                                .Viral = False
                                Tiles = Tiles - 1
                                If Tiles <= 0 Then
                                    RaiseEvent VirusKilled(OldVirii)
                                    doWinGame
                                End If
                            End If
                            .color = 0
                        End If
                    End If
                End If
            End With
        Next
    Next
    DropAll
    If waslocked And Not Locked Then
        If OldVirii <> Total Then
            RaiseEvent VirusKilled(OldVirii - Total)
            OldVirii = Total
        End If
    End If
End Function

Public Sub DecrementCurrentTile(Optional Pixels As Long = 1)
    If Gameisover Then Exit Sub
    With CurrentTile
        If .Color1 > 0 And .Color2 > 0 Then
            .YOffset = .YOffset - Pixels
            If .YOffset = 0 Then
                If MoveCurrentTileDown Then
                    .y = .y - 1
                    .YOffset = TileSize
                Else
                    Locked = TransferTile
                End If
            End If
        End If
        If Not Locked Then
            If .Color1 = 0 And .Color2 = 0 Then .YOffset = 0
            If .YOffset = 0 Then PullTile
        End If
    End With
End Sub

'Find the start of concurrent tiles in a row
Private Function StartX(x As Long, y As Long, Optional color As Long = -1) As Long
    Dim temp As Long
    If color = -1 Then color = Grid(y, x).color
    For temp = x - 1 To 1 Step -1
        If Grid(y, temp).color <> color Then
            StartX = temp + 1
            Exit Function
        End If
    Next
    StartX = 1
End Function

'Find the start of concurrent tiles in a column
Private Function StartY(x As Long, y As Long, Optional color As Long = -1) As Long
    Dim temp As Long
    If color = -1 Then color = Grid(y, x).color
    For temp = y - 1 To 1 Step -1
        If Grid(temp, x).color <> color Then
            StartY = temp + 1
            Exit Function
        End If
    Next
    StartY = 1
End Function

'Find the end of concurrent tiles in a row
Private Function EndX(x As Long, y As Long, Optional color As Long = -1) As Long
    Dim temp As Long
    If color = -1 Then color = Grid(y, x).color
    For temp = x + 1 To GridWidth
        If Grid(y, temp).color <> color Then
            EndX = temp - 1
            Exit Function
        End If
    Next
    EndX = GridWidth
End Function

'Find the end of concurrent tiles in a column
Private Function EndY(x As Long, y As Long, Optional color As Long = -1) As Long
    Dim temp As Long
    If color = -1 Then color = Grid(y, x).color
    For temp = y + 1 To GridHeight
        If Grid(temp, x).color <> color Then
            EndY = temp - 1
            Exit Function
        End If
    Next
    EndY = GridHeight
End Function

'Find the number of concurrent tiles in a row
Private Function CountX(x As Long, y As Long, Optional color As Long = -1) As Long
    Dim temp As Long, temp2 As Long
    temp = StartX(x, y, color)
    temp2 = EndX(x, y, color)
    CountX = temp2 - temp + 1
End Function

'Find the number of concurrent tiles in a column
Private Function CountY(x As Long, y As Long, Optional color As Long = -1) As Long
    Dim temp As Long, temp2 As Long
    temp = StartY(x, y, color)
    temp2 = EndY(x, y, color)
    CountY = temp2 - temp + 1
End Function

'Check if the resulting tile placement will result in a score
Private Function IsAScore(x As Long, y As Long, color As Long) As Boolean
    IsAScore = color > 0 And (CountX(x, y, color) > 3 Or CountY(x, y, color) > 3)
End Function

'Ghosts a single tile
Private Sub GhostTile(x As Long, y As Long)
    If x > 0 And x <= GridWidth And y >= 0 And y <= GridHeight Then 'AHAHAH, stupidly put gridwidth in the Y checked
        With Grid(y, x)
            'MsgBox "Killing " & x & " , " & y
            'MsgBox "Bound = " & .Bound & vbNewLine & "Target = " & .bX & ", " & .bY
            .Clean = False
            .Z = Ghost
            If .Bound Or .bX <> 0 Or .bY <> 0 Or .Angle > 0 Then
                Grid(.bY, .bX).Bound = False
                Grid(.bY, .bX).Angle = NoAngle
                Grid(.bY, .bX).Clean = False
                DropColumn .bX, .bY
                .Bound = False
                .Angle = NoAngle
            End If
        End With
    End If
End Sub

'Drop a tiles above a certain tile
Private Sub DropColumn(x As Long, ByVal y As Long)
    For y = y To GridHeight
            If x > 0 Then If Grid(y, x).color > 0 Then MoveTilesDown x, y
    Next
End Sub

'Ghosts a horizontal Score
Private Function GhostX(x As Long, y As Long, X2 As Long) As Long
    Dim temp As Long
    For temp = x To X2
        GhostTile temp, y
    Next
    GhostX = X2 - x + 1
End Function

'Ghosts a horizontal Score
Private Function GhostY(x As Long, y As Long, Y2 As Long) As Long
    Dim temp As Long
    For temp = y To Y2
        GhostTile x, temp
    Next
    GhostY = Y2 - y + 1
End Function

Private Function CheckTile(x As Long, y As Long) As Long
    Dim pLeft As Long, pRight As Long, pTop As Long, pBottom As Long, color As Long, HOR As Boolean, VER As Boolean
    color = Grid(y, x).color
    
    pLeft = StartX(x, y, color)
    pRight = EndX(x, y, color)
    HOR = pRight - pLeft + 1 > 3
        
    pTop = StartY(x, y, color)
    pBottom = EndY(x, y, color)
    VER = pBottom - pTop + 1 > 3
        
    color = 0
    If HOR Then color = GhostX(pLeft, y, pRight)
    If VER Then color = color + GhostY(x, pTop, pBottom)
    CheckTile = color
    If color > 3 Then RaiseEvent TilesCleared(color)
End Function

Private Sub doWinGame()
    DrawScreen
    Gameisover = True
    RaiseEvent WonGame
    ClearBoard
End Sub

Public Function Score(ByVal Speed As Long, ViriiKilled As Long) As Long
    Dim temp As Long
    Speed = Speed + 1
    Select Case ViriiKilled
        Case 0
        Case 1, 2: temp = ViriiKilled
        Case Else: temp = 2 ^ ViriiKilled
    End Select
    Score = temp * 100 * Speed
End Function
