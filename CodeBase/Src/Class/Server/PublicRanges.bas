Attribute VB_Name = "PublicRanges"


Option Explicit

Public ItemsStart          As IRange
Public QuestsStart         As IRange
Public ScriptsStart        As IRange
Public AttacksStart        As IRange
Public FumonsStart         As IRange
Public FumonSpawnersStart  As IRange
Public ElementTypesStart   As IRange
Public PlayersStart        As IRange
Public ServerUpdatesStart  As IRange
Public PlayerUpdatesStart  As IRange
Public FightsStart         As IRange
Public WildPlayersStart    As IRange
Public TilesStart          As IRange
Public MapDataStart        As IRange
Public GameMapsStart       As IRange

Public Sub InitializeAllRanges(ByVal WB As WorkBook)
    Set ItemsStart          = CreateFasterRange(WB.Sheets("Items").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set QuestsStart         = CreateFasterRange(WB.Sheets("Quests").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set ScriptsStart        = CreateFasterRange(WB.Sheets("Scripts").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set AttacksStart        = CreateFasterRange(WB.Sheets("Attacks").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set FumonsStart         = CreateFasterRange(WB.Sheets("Fumons").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set FumonSpawnersStart  = CreateFasterRange(WB.Sheets("FumonSpawners").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set ElementTypesStart   = CreateFasterRange(WB.Sheets("Fumons").Range("V1")).CreateFromSlice(1, 22, 1, 1)
    Set PlayersStart        = CreateFasterRange(WB.Sheets("Players").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set ServerUpdatesStart  = CreateFasterRange(WB.Sheets("ServerUpdates").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set PlayerUpdatesStart  = CreateFasterRange(WB.Sheets("PlayerUpdates").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set FightsStart         = CreateFasterRange(WB.Sheets("Fights").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set WildPlayersStart    = CreateFasterRange(WB.Sheets("WildPlayers").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set TilesStart          = CreateFasterRange(WB.Sheets("Tiles").Range("A1")).CreateFromSlice(2, 1, 1, 1)
    Set MapDataStart        = CreateFasterRange(WB.Sheets("MapData").Range("A1")).CreateFromSlice(2, 2, 1, 1)
    Set GameMapsStart       = CreateFasterRange(WB.Sheets("GameMaps").Range("A1")).CreateFromSlice(2, 1, 1, 1)
End Sub

Public Function CreateFasterRange(ByVal StartRange As Range) As FasterRange
    Dim RowRange     As Range : Set RowRange     = StartRange.End(xlDown)
    Dim ColumnRange  As Range : Set ColumnRange  = StartRange.End(xlToRight)
    Dim RowOffset    As Long  : Let RowOffset    = RowRange.Row - StartRange.Row
    Dim ColumnOffset As Long  : Let ColumnOffset = ColumnRange.Column - StartRange.Column

    If RowOffset > 16384 Then RowOffset = 16384 ' Limit so nobody exceeds to many items
    Set CreateFasterRange = FasterRange.CreateFromRange(Range(StartRange, StartRange.Offset(RowOffset, ColumnOffset)))
End Function