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

Private p_Number       As IRange
Private p_Name         As IRange
Private p_Rows         As IRange
Private p_Columns      As IRange
Private p_SpawnRow     As IRange
Private p_SpawnColumn  As IRange
Private p_EntireMap    As IRange
Private p_MapPointer   As IRange
Private p_Tiles        As Tiles



Public Property Let Number(n_Number               As IRange)   : Set p_Number        = n_Number         : End Property
Public Property Let Name(n_Name                   As IRange)   : Set p_Name          = n_Name           : End Property
Public Property Let Rows(n_Rows                   As IRange)   : Set p_Rows          = n_Rows           : End Property
Public Property Let Columns(n_Columns             As IRange)   : Set p_Columns       = n_Columns        : End Property
Public Property Let SpawnRow(n_SpawnRow           As IRange)   : Set p_SpawnRow      = n_SpawnRow       : End Property
Public Property Let SpawnColumn(n_SpawnColumn     As IRange)   : Set p_SpawnColumn   = n_SpawnColumn    : End Property
Public Property Let EntireMap(n_EntireMap         As IRange)   : Set p_EntireMap     = n_EntireMap      : End Property
Public Property Let MapPointer(n_MapPointer       As IRange)   : Set p_MapPointer    = n_MapPointer     : End Property
Public Property Let Tiles(n_Tiles                 As Tiles)    : Set p_Tiles         = n_Tiles          : End Property

Public Property Get Number()                      As IRange    : Set Number          = p_Number         : End Property
Public Property Get Name()                        As IRange    : Set Name            = p_Name           : End Property
Public Property Get Rows()                        As IRange    : Set Rows            = p_Rows           : End Property
Public Property Get Columns()                     As IRange    : Set Columns         = p_Columns        : End Property
Public Property Get SpawnRow()                    As IRange    : Set SpawnRow        = p_SpawnRow       : End Property
Public Property Get SpawnColumn()                 As IRange    : Set SpawnColumn     = p_SpawnColumn    : End Property
Public Property Get EntireMap()                   As IRange    : Set EntireMap       = p_EntireMap      : End Property
Public Property Get MapPointer()                  As IRange    : Set MapPointer      = p_MapPointer     : End Property
Public Property Get Tiles()                       As Tiles     : Set Tiles           = p_Tiles          : End Property

Public Function Create(ByVal WB As Workbook, ByVal Rng As IRange, ByVal TileDefinitions As PropCollection) As GameMap
    Set Create = New GameMap
    Dim Temp As Tiles: Set Temp = New Tiles
    With Create
        .Number        = Rng.Offset(0, 0)
        .Name          = Rng.Offset(0, 1)
        .Rows          = Rng.Offset(0, 2)
        .Columns       = Rng.Offset(0, 3)
        .SpawnRow      = Rng.Offset(0, 4)
        .SpawnColumn   = Rng.Offset(0, 5)
        .EntireMap     = CreateFasterRange(WB.Sheets(.Name.Value).Cells(1, 1))
        .MapPointer    = .EntireMap.GetSelf.CreateFromSlice(1, 1, 1, 1)
        .Tiles         = Temp.Create(.MapPointer, CLng(.Rows.Value), CLng(.Columns.Value), TileDefinitions)
    End With
End Function

Public Function InBounds(ByVal Y As Long, ByVal X As Long) As Boolean
    InBounds = Y >= 0              And _
               Y =< Rows.Value     And _
               X >= 0              And _
               X =< Columns.Value
End Function

Public Function Traverseable(ByVal Y As Long, ByVal X As Long) As Boolean
    If InBounds(Y, X) Then
        Dim nTile As Tile
        Set nTile = Tiles.Tile(Y, X)
        If nTile.Tile.Speed = 0        Then Exit Function
        If IsSomething(nTile.Player)   Then Exit Function
        Traverseable = True
    End If
End Function

Public Sub ClearPlayers()
    Dim RowCount    As Long    : RowCount    = EntireMap.GetSelf.Parent.RowCount
    Dim ColumnCount As Long    : ColumnCount = EntireMap.GetSelf.Parent.ColumnCount
    Dim n_Values()  As Variant : ReDim n_Values(1 To RowCount, 1 To ColumnCount)

    Dim y As Long, x As Long
    For y = 1 To RowCount
        For x = 1 To ColumnCount
            n_Values(y, x) = Tiles.Tile(y - 1, x - 1).MapPointPlayer(-1)
        Next x
    Next y
    Call Update(n_Values)
End Sub

Public Sub SetSpawnPoints(ByRef Players() As IPlayer)
    Dim RowCount    As Long    : RowCount    = EntireMap.GetSelf.Parent.RowCount
    Dim ColumnCount As Long    : ColumnCount = EntireMap.GetSelf.Parent.ColumnCount
    Dim n_Values()  As Variant : ReDim n_Values(1 To RowCount, 1 To ColumnCount)

    Dim y As Long, x As Long
    Dim i As Long
    For i = 0 To USize(Players)
        With Players(i).MoveBase
            If .Map Is Me Then
                y = .SpawnRow.Value
                x = .SpawnColumn.Value
                n_Values(y + 1, x + 1) = Tiles.Tile(y, x).MapPointPlayer(Players(i).PlayerBase.Number.Value)
            End If
        End With
    Next i

    For y = 1 To RowCount
        For x = 1 To ColumnCount
            If IsEmpty(n_Values(y, x)) Then
                n_Values(y, x) = Tiles.Tile(y - 1, x - 1).MapPoint()
            End If
        Next x
    Next y
    Call Update(n_Values)
End Sub

Public Sub Update(ByRef n_Values() As Variant)
    EntireMap.Value = n_Values
End Sub