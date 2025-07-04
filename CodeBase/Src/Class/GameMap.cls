VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Name         As Range
Private p_Rows         As Range
Private p_Columns      As Range
Private p_Year         As Range
Private p_Month        As Range
Private p_Day          As Range
Private p_Hour         As Range
Private p_Minute       As Range
Private p_Folder       As Range
Private p_SpawnRow     As Range
Private p_SpawnColumn  As Range
Private p_Activefights As Range
Private p_MapPointer   As Range
Private p_Tiles()      As Tile



Public Property Let Name(n_Name                 As Range): Set p_Name         = n_Name          : End Property
Public Property Let Rows(n_Rows                 As Range): Set p_Rows         = n_Rows          : End Property
Public Property Let Columns(n_Columns           As Range): Set p_Columns      = n_Columns       : End Property
Public Property Let Year(n_Year                 As Range): Set p_Year         = n_Year          : End Property
Public Property Let Month(n_Month               As Range): Set p_Month        = n_Month         : End Property
Public Property Let Day(n_Day                   As Range): Set p_Day          = n_Day           : End Property
Public Property Let Hour(n_Hour                 As Range): Set p_Hour         = n_Hour          : End Property
Public Property Let Minute(n_Minute             As Range): Set p_Minute       = n_Minute        : End Property
Public Property Let Folder(n_Folder             As Range): Set p_Folder       = n_Folder        : End Property
Public Property Let SpawnRow(n_SpawnRow         As Range): Set p_SpawnRow     = n_SpawnRow      : End Property
Public Property Let SpawnColumn(n_SpawnColumn   As Range): Set p_SpawnColumn  = n_SpawnColumn   : End Property
Public Property Let Activefights(n_Activefights As Range): Set p_Activefights = n_Activefights  : End Property
Public Property Let MapPointer(n_MapPointer     As Range): Set p_MapPointer   = n_MapPointer    : End Property
Public Property Let Tiles(n_Tiles()             As Tile) : Let p_Tiles        = n_Tiles         : End Property

Public Property Let Tile(Index As Long, n_Tile  As Tile) : Set p_Tiles(Index) = n_Tile          : End Property

Public Property Get Name()                      As Range : Set Name           = p_Name          : End Property
Public Property Get Rows()                      As Range : Set Rows           = p_Rows          : End Property
Public Property Get Columns()                   As Range : Set Columns        = p_Columns       : End Property
Public Property Get Year()                      As Range : Set Year           = p_Year          : End Property
Public Property Get Month()                     As Range : Set Month          = p_Month         : End Property
Public Property Get Day()                       As Range : Set Day            = p_Day           : End Property
Public Property Get Hour()                      As Range : Set Hour           = p_Hour          : End Property
Public Property Get Minute()                    As Range : Set Minute         = p_Minute        : End Property
Public Property Get Folder()                    As Range : Set Folder         = p_Folder        : End Property
Public Property Get SpawnRow()                  As Range : Set SpawnRow       = p_SpawnRow      : End Property
Public Property Get SpawnColumn()               As Range : Set SpawnColumn    = p_SpawnColumn   : End Property
Public Property Get Activefights()              As Range : Set Activefights   = p_Activefights  : End Property
Public Property Get MapPointer()                As Range : Set MapPointer     = p_MapPointer    : End Property
Public Property Get Tiles()                     As Tile(): Let Tiles          = p_Tiles         : End Property

Public Property Get Tile(Index As Long)         As Tile  : Set Tile           = p_Tiles(Index)  : End Property

Public Function Create(Map As Range, Rng As Range) As GameMap
    Set Create = New GameMap
    With Create
        .Name         = Rng.Offset(00, 0)
        .Rows         = Rng.Offset(01, 0)
        .Columns      = Rng.Offset(02, 0)
        .Year         = Rng.Offset(03, 0)
        .Month        = Rng.Offset(04, 0)
        .Day          = Rng.Offset(05, 0)
        .Hour         = Rng.Offset(06, 0)
        .Minute       = Rng.Offset(07, 0)
        .Folder       = Rng.Offset(08, 0)
        .SpawnRow     = Rng.Offset(09, 0)
        .SpawnColumn  = Rng.Offset(10, 0)
        .Activefights = Rng.Offset(11, 0)
        .MapPointer   = Map
        .Tiles        = .GetGameTiles(Map)
    End With
End Function

Public Function GetGameTiles(Map As Range) As Tile()
    Dim ReturnArray() As Tile
    Dim x As Long, y As Long, Count As Long
    Dim Rng As Range, Rng2 As Range
    ReDim ReturnArray((Rows.Value + 1) * (Columns.Value + 1) - 1)
    Set Rng = ActiveWorkbook.Worksheets("Tiles").Range("A2")
    For y = 0 To Rows.Value
        For x = 0 To Columns.Value
            Set Rng2 = Rng.Offset(CLng(ExtractPoint(Map.Offset(y, x).Value, 0)), 0)
            Set ReturnArray(Count) = FumonGame.Tile.Create(Rng2, Folder.Value)
            Count = Count + 1
        Next x
    Next y 
    GetGameTiles = ReturnArray
End Function

Private Function ExtractPoint(Text As String, Index As Long) As String
    Dim Numbers() As String
    Numbers = Split(Text, ", ")
    ExtractPoint = Numbers(Index)
End Function

Public Function GetTile(Y As Long, X As Long) As Tile
    Dim Words() As String
    Words = Split(MapPointer.Offset(Y, X).Value, ", ")
    GetTile = Tile(CLng(Words(0)))
End Function

Public Function GetPlayer(Y As Long, X As Long) As Tile
    Dim Words() As String
    Words = Split(MapPointer.Offset(Y, X).Value, ", ")
    GetTile = Tile(CLng(Words(1)))
End Function

Public Function GetScript(Y As Long, X As Long) As Tile
    Dim Words() As String
    Words = Split(MapPointer.Offset(Y, X).Value, ", ")
    GetTile = Tile(CLng(Words(2)))
End Function