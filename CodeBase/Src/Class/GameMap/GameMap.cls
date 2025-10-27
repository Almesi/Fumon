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

Private p_Number       As Range
Private p_Name         As Range
Private p_Rows         As Range
Private p_Columns      As Range
Private p_SpawnRow     As Range
Private p_SpawnColumn  As Range
Private p_MapPointer   As Range
Private p_Tiles        As Tiles



Public Property Let Number(n_Number               As Range)   : Set p_Number        = n_Number         : End Property
Public Property Let Name(n_Name                   As Range)   : Set p_Name          = n_Name           : End Property
Public Property Let Rows(n_Rows                   As Range)   : Set p_Rows          = n_Rows           : End Property
Public Property Let Columns(n_Columns             As Range)   : Set p_Columns       = n_Columns        : End Property
Public Property Let SpawnRow(n_SpawnRow           As Range)   : Set p_SpawnRow      = n_SpawnRow       : End Property
Public Property Let SpawnColumn(n_SpawnColumn     As Range)   : Set p_SpawnColumn   = n_SpawnColumn    : End Property
Public Property Let MapPointer(n_MapPointer       As Range)   : Set p_MapPointer    = n_MapPointer     : End Property
Public Property Let Tiles(n_Tiles                 As Tiles)   : Set p_Tiles         = n_Tiles          : End Property

Public Property Get Number()                      As Range    : Set Number          = p_Number         : End Property
Public Property Get Name()                        As Range    : Set Name            = p_Name           : End Property
Public Property Get Rows()                        As Range    : Set Rows            = p_Rows           : End Property
Public Property Get Columns()                     As Range    : Set Columns         = p_Columns        : End Property
Public Property Get SpawnRow()                    As Range    : Set SpawnRow        = p_SpawnRow       : End Property
Public Property Get SpawnColumn()                 As Range    : Set SpawnColumn     = p_SpawnColumn    : End Property
Public Property Get MapPointer()                  As Range    : Set MapPointer      = p_MapPointer     : End Property
Public Property Get Tiles()                       As Tiles    : Set Tiles           = p_Tiles          : End Property

Public Function Create(ByVal WB As Workbook, ByVal Rng As Range, ByVal TileDefinitions As PropCollection) As GameMap
    Set Create = New GameMap
    Dim Temp As Tiles: Set Temp = New Tiles
    With Create
        .Number        = Rng.Offset(0, 0)
        .Name          = Rng.Offset(0, 1)
        .Rows          = Rng.Offset(0, 2)
        .Columns       = Rng.Offset(0, 3)
        .SpawnRow      = Rng.Offset(0, 4)
        .SpawnColumn   = Rng.Offset(0, 5)
        .MapPointer    = WB.Worksheets(.Name.Value).Range("A1")
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
    Dim y As Long, x As Long
    For y = 0 To Rows.Value
        For x = 0 To Columns.Value
            Tiles.Tile(y, x).SetMapPointPlayer(-1)
        Next x
    Next y
End Sub

Public Sub SetSpawnPoints(ByRef Players() As IPlayer)
    Dim y As Long, x As Long
    Dim i As Long
    For i = 0 To USize(Players)
        With Players(i).MoveBase
            If .Map Is Me Then
                y = .SpawnRow.Value
                x = .SpawnColumn.Value
                Call Tiles.Tile(y, x).SetMapPointPlayer(Players(i).Number.Value)
            End If
        End With
    Next i
End Sub