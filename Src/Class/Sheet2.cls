VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit


Public Sub RefreshMap()
    Dim y As Long, x As Long
    Dim MapRow As Long, MapColumn As Long
    Dim CurrentTile As String
    Dim Map As Worksheet
    Dim Cell As Range

    Set MapSettings = std_Config.Create(ThisWorkbook.Sheets("MapData"))

    MapRow = MapSettings.Setting("Rows")
    MapColumn = MapSettings.Setting("Columns")

    Dim FilePath As String
    Dim Image As Object

    Set Map = ThisWorkbook.Sheets("Map")
    Map.Pictures.Delete
    For y = 1 To MapRow
        For x = 1 To MapColumn
            Set Cell = Map.Cells(y, x)
            CurrentTile = ExtractPoint(Cell.Value, "Tile")
            FilePath = MapSettings.Setting("SpriteFolder") & GetData(CurrentTile, "Tiles", 0, "Name") & ".png"
            Set Image = Map.Pictures.Insert(FilePath)
            With Image
                .Top  = Cell.Top
                .Left = Cell.Left
                .ShapeRange.Height = Cell.Height
                .ShapeRange.Width = Cell.Width
            End With
        Next x
    Next y
End Sub