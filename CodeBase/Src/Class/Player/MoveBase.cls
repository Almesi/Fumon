VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MoveBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private p_PlayerBase        As PlayerBase
Private p_Money             As IRange
Private p_Map               As GameMap
Private p_Row               As IRange
Private p_Column            As IRange
Private p_SpawnRow          As IRange
Private p_SpawnColumn       As IRange
Private p_LookDirection     As IRange
Private p_DoInteract        As Boolean

Public Function Offset() As Long
    Offset = 7
End Function


Public Property Let PlayerBase(ByVal n_PlayerBase         As PlayerBase)  : Set p_PlayerBase      = n_PlayerBase        : End Property
Public Property Let Money(ByVal n_Money                   As IRange)      : Set p_Money           = n_Money             : End Property
Public Property Let Map(ByVal n_Map                       As GameMap)     : Set p_Map             = n_Map               : End Property
Public Property Let Row(ByVal n_Row                       As IRange)      : Set p_Row             = n_Row               : End Property
Public Property Let Column(ByVal n_Column                 As IRange)      : Set p_Column          = n_Column            : End Property
Public Property Let SpawnRow(ByVal n_SpawnRow             As IRange)      : Set p_SpawnRow        = n_SpawnRow          : End Property
Public Property Let SpawnColumn(ByVal n_SpawnColumn       As IRange)      : Set p_SpawnColumn     = n_SpawnColumn       : End Property
Public Property Let LookDirection(ByVal n_LookDirection   As IRange)      : Set p_LookDirection   = n_LookDirection     : End Property
Public Property Let DoInteract(ByVal n_DoInteract         As Boolean)     : Let p_DoInteract      = n_DoInteract        : End Property

Public Property Get PlayerBase()                          As PlayerBase   : Set PlayerBase        = p_PlayerBase        : End Property
Public Property Get Money()                               As IRange       : Set Money             = p_Money             : End Property
Public Property Get Map()                                 As GameMap      : Set Map               = p_Map               : End Property
Public Property Get Row()                                 As IRange       : Set Row               = p_Row               : End Property
Public Property Get Column()                              As IRange       : Set Column            = p_Column            : End Property
Public Property Get SpawnRow()                            As IRange       : Set SpawnRow          = p_SpawnRow          : End Property
Public Property Get SpawnColumn()                         As IRange       : Set SpawnColumn       = p_SpawnColumn       : End Property
Public Property Get LookDirection()                       As IRange       : Set LookDirection     = p_LookDirection     : End Property
Public Property Get DoInteract()                          As Boolean      : Let DoInteract        = p_DoInteract        : End Property

Public Function Create(ByVal Rng As IRange, ByVal n_PlayerBase As PlayerBase, ByVal GameMaps As PropCollection) As MoveBase
    Set Create = New MoveBase
    With Create
        .PlayerBase  = n_PlayerBase
        .Money         = Rng.Offset(0, 0)
        .Map           = GameMaps.Object(Rng.Offset(0, 1).Value)
        .Row           = Rng.Offset(0, 2)
        .Column        = Rng.Offset(0, 3)
        .SpawnRow      = Rng.Offset(0, 4)
        .SpawnColumn   = Rng.Offset(0, 5)
        .LookDirection = Rng.Offset(0, 6)
        .Row.Value     = .SpawnRow.Value
        .Column.Value  = .SpawnColumn.Value
    End With
End Function

Public Sub Teleport(ByVal x As Long, ByVal y As Long)
    Dim Seed As Date
    Dim RandomNumber As Double
    Seed = Date()
    RandomNumber = Rnd(Seed)
    With Map.Tiles
        If Map.Traverseable(y, x) Then
            ' If random value is smaller than chance to go to new tile [1>=0,1 = guaranteed]
            If .Tile(y, x).Tile.Speed >= RandomNumber Then
                Call .Tile(Row.Value, Column.Value).SetMapPointPlayer(-1)
                Call .Tile(y, x).SetMapPointPlayer(PlayerBase.Number.Value)
                Call Look(DirectionByMovement(x - Column.Value, y - Row.Value))
                Row.Value    = y
                Column.Value = x
                Call TriggerEvent()
            End If
        Else
            ' Soundtrack for "cannot enter tile"
        End If
    End With
End Sub

Public Sub Move(ByVal x As Long, ByVal y As Long)
    Call Teleport(Column.Value + x, Row.Value + y)
End Sub

Public Sub MovePath(ByRef x() As Long, ByRef y() As Long)
    Dim i As Long
    For i = 0 To USize(x)
        Call Move(x(i), y(i))
    Next i
End Sub

Public Sub Look(ByVal Direction As XlDirection)
    LookDirection.Value = Direction
End Sub

Public Sub Interact(ByVal MyPlayer As IPlayer, ByVal Offset As Long)
    If DoInteract = False Then Exit Sub
    Dim TileRange As Tile         : Set TileRange = GetTileRange(Offset)
    Dim Player    As IPlayer      : Set Player    = TileRange.Player

    If IsSomething(Player)  Then Call Player.SpecificBase.RunScript(Player, MyPlayer)
End Sub

Public Function InFront(ByVal Offset As Long) As Tile
    Select Case LookDirection.Value
        Case xlUp    : Set InFront = Map.Tiles.Tile(Row.Value - Offset, Column.Value)
        Case xlDown  : Set InFront = Map.Tiles.Tile(Row.Value + Offset, Column.Value)
        Case xlLeft  : Set InFront = Map.Tiles.Tile(Row.Value         , Column.Value - Offset)
        Case xlRight : Set InFront = Map.Tiles.Tile(Row.Value         , Column.Value + Offset)
    End Select
End Function

Public Function SubTextureName(ByVal Name As String) As String
    Select Case LookDirection.Value
        Case xlUp    : SubTextureName = Name & "Up"
        Case xlLeft  : SubTextureName = Name & "Left"
        Case xlDown  : SubTextureName = Name & "Down"
        Case xlRight : SubTextureName = Name & "Right"
    End Select
End Function


Private Function DirectionByMovement(ByVal x As Long, ByVal y As Long) As XlDirection
    Select Case True
        Case x >= +1 And y =  0  : DirectionByMovement = xlRight
        Case x =< -1 And y =  0  : DirectionByMovement = xlLeft
        Case x =  +0 And y >= +1 : DirectionByMovement = xlDown
        Case x =  +0 And y =< -1 : DirectionByMovement = xlUp
    End Select
End Function

Private Sub TriggerEvent()
    If DoInteract = False Then Exit Sub
    Dim TileRange  As Tile         : Set TileRange = GetTileRange(0)
    Dim Script     As Script       : Set Script    = TileRange.Script
    Dim Spawner    As FumonSpawner : Set Spawner   = TileRange.Spawner
    
    If IsSomething(Spawner) Then
        Call StartFightWithWilPlayer(GetPlayer(PlayerBase.Number.Value), Spawner)
    End If
    If IsSomething(Script) Then
        If Script.Active = 0 Then
            Dim Temp() As Variant
            Call Script.Run(Temp)
        End If
    End If
End Sub

Private Function GetTileRange(ByVal Offset As Long) As Tile
    Dim x As Long, y As Long
    x = Column.Value
    y = Row.Value
    Select Case LookDirection.Value
        Case xlUp    : Set GetTileRange = Map.Tiles.Tile(y - Offset, x)
        Case xlDown  : Set GetTileRange = Map.Tiles.Tile(y + Offset, x)
        Case xlLeft  : Set GetTileRange = Map.Tiles.Tile(y, x - Offset)
        Case xlRight : Set GetTileRange = Map.Tiles.Tile(y, x + Offset)
    End Select
End Function