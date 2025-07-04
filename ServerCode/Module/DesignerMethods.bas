Attribute VB_Name = "DesignerMethods"


Option Explicit

Private Const ArgSeperator As String = ", "

Public Function MakeTile(TileName As String, PlayerName As String, ScriptName As String) As String
    MakeTile = GetTile(TileName) & ArgSeperator & GetPlayer(PlayerName) & ArgSeperator & GetScript(ScriptName)
End Function

Public Function GetTile(Name   As String) As Long: GetTile   = GetIndexByName(ThisWorkbook.Worksheets("Tiles")   , Name) : End Function
Public Function GetPlayer(Name As String) As Long: GetPlayer = GetIndexByName(ThisWorkbook.Worksheets("Players") , Name) : End Function
Public Function GetScript(Name As String) As Long: GetScript = GetIndexByName(ThisWorkbook.Worksheets("Scripts") , Name) : End Function

Private Function GetIndexByName(WS AS Worksheet, Name As String) As Long
    Dim i As Long
    Dim Rng As Range

    Set Rng = WS.Range("A2")
    Do While Rng.Offset(i, 0).Formula <> Empty
        If Rng.Offset(i, 0).Value = Name Then
            GetIndexByName = Rng.Offset(i, -1).Value
        End If
        i = i + 1
    Loop
    GetIndexByName = -1
End Function

Public Function MergeStringArray(Values() As String) As String
    Dim i As Long
    MergeStringArray = Values(0)
    For i = 1 To Ubound(Values)
        MergeStringArray = MergeStringArray & ArgSeperator & Values(i)
    Next i
End Function

Public Sub UpdatePlayer()
    Dim i As Long
    Dim Words() As String
    Dim x As Long, y As Long

    Call SetupGameServer()
    With FumonGame.MeServer
        For i = 0 To Ubound(.GameMap.Tiles)
            Words = .GameMap.Tile(i).Value
            Words(1) = -1
            .GameMap.Tile(i).Value = MergeStringArray(Words)
        Next i
        
        For i = 0 To Ubound(.Players)
            x = .Player(i).Row.Value
            y = .Player(i).Column.Value
            Words = Split(.GameMap.MapPointer.Offset(y, x).Value)
            Words(1) = .Player(i).Number
            .GameMap.MapPointer.Offset(y, x).Value = MergeStringArray(Words)
        Next i
    End With
End Sub


Public Sub UpdateMap()
    'Refresh Tile Icon and Player Icons
End Sub