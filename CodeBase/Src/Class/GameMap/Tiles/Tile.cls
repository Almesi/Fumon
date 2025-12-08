VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Tile     As TileDefinition
Private p_Obstacle As TileDefinition
Private p_Range    As IRange

Public Property Let Tile(n_Tile          As TileDefinition) : Set p_Tile     = n_Tile     : End Property
Public Property Let Obstacle(n_Obstacle  As TileDefinition) : Set p_Obstacle = n_Obstacle : End Property
Public Property Let Range(n_Range        As IRange)         : Set p_Range    = n_Range    : End Property

Public Property Get Tile()               As TileDefinition  : Set Tile       = p_Tile     : End Property
Public Property Get Obstacle()           As TileDefinition  : Set Obstacle   = p_Obstacle : End Property
Public Property Get Range()              As IRange          : Set Range      = p_Range    : End Property

Public Function Create(Rng As IRange, ByVal TileDefinitions As PropCollection) As Tile
    Dim Index As Long
    Set Create = New Tile
    With Create
        .Range      = Rng
        Index = CLng(.ExtractPoint("Tiles"))
        If Index <> -1 Then .Tile = TileDefinitions.Object(Index)
    End With
End Function

Public Function Player() As IPlayer
    Dim Index As Long
    Index = CLng(ExtractPoint("Players"))
    If Index <> -1 Then Set Player = MeServer.Player(Index)
End Function
Public Function Spawner() As FumonSpawner
    Dim Index As Long
    Index = CLng(ExtractPoint("Spawners"))
    If Index <> -1 Then Set Spawner = MeServer.FumonSpawner(Index)
End Function
Public Function Script() As Script
    Dim Index As Long
    Index = CLng(ExtractPoint("Scripts"))
    If Index <> -1 Then Script = MeServer.Script(Index)
End Function

Public Function MapPoint(Optional ByVal Tile     As String = "KEEP", _
                         Optional ByVal Obstacle As String = "KEEP", _
                         Optional ByVal Player   As String = "KEEP", _
                         Optional ByVal Spawner  As String = "KEEP", _
                         Optional ByVal Script   As String = "KEEP") As String
    Dim Values() As String
    Values = Split(Range.Value, ", ")
    If Tile     <> "KEEP" Then Values(0) = Tile
    If Obstacle <> "KEEP" Then Values(1) = Obstacle
    If Player   <> "KEEP" Then Values(2) = Player
    If Spawner  <> "KEEP" Then Values(3) = Spawner
    If Script   <> "KEEP" Then Values(4) = Script

    MapPoint = Values(0) & ", " & Values(1) & ", " & Values(2) & ", " & Values(3) & ", " & Values(4)
End Function

Public Function SetMapPoint(Optional ByVal Tile     As String = "KEEP", _
                            Optional ByVal Obstacle As String = "KEEP", _
                            Optional ByVal Player   As String = "KEEP", _
                            Optional ByVal Spawner  As String = "KEEP", _
                            Optional ByVal Script   As String = "KEEP")
    Range.Value = MapPoint(Tile, Obstacle, Player, Spawner, Script)
End Function

Public Function MapPointTile(ByVal Tile         As Long) : MapPointTile     = MapPoint(Tile  , "KEEP"  , "KEEP", "KEEP" , "KEEP") : End Function
Public Function MapPointObstacle(ByVal Obstacle As Long) : MapPointObstacle = MapPoint("KEEP", Obstacle, "KEEP", "KEEP" , "KEEP") : End Function
Public Function MapPointPlayer(ByVal Player     As Long) : MapPointPlayer   = MapPoint("KEEP", "KEEP"  , Player, "KEEP" , "KEEP") : End Function
Public Function MapPointSpawner(ByVal Spawner   As Long) : MapPointSpawner  = MapPoint("KEEP", "KEEP"  , "KEEP", Spawner, "KEEP") : End Function
Public Function MapPointScript(ByVal Script     As Long) : MapPointScript   = MapPoint("KEEP", "KEEP"  , "KEEP", "KEEP" , Script) : End Function

Public Sub SetMapPointTile(ByVal Tile         As Long)   : Call SetMapPoint(Tile  , "KEEP"  , "KEEP", "KEEP" , "KEEP") : End Sub
Public Sub SetMapPointObstacle(ByVal Obstacle As Long)   : Call SetMapPoint("KEEP", Obstacle, "KEEP", "KEEP" , "KEEP") : End Sub
Public Sub SetMapPointPlayer(ByVal Player     As Long)   : Call SetMapPoint("KEEP", "KEEP"  , Player, "KEEP" , "KEEP") : End Sub
Public Sub SetMapPointSpawner(ByVal Spawner   As Long)   : Call SetMapPoint("KEEP", "KEEP"  , "KEEP", Spawner, "KEEP") : End Sub
Public Sub SetMapPointScript(ByVal Script     As Long)   : Call SetMapPoint("KEEP", "KEEP"  , "KEEP", "KEEP" , Script) : End Sub


Public Function ExtractPoint(ByVal NameOrIndex As String) As Variant
    Dim Numbers() As String
    Numbers = Split(Range.Value, ", ")
    If IsNumeric(NameOrIndex) Then
        ExtractPoint = Numbers(CLng(NameOrIndex))
    Else
        Select Case UCase(NameOrIndex)
            Case "TILES"      : ExtractPoint = CLng(Numbers(0))
            Case "OBSTACLES"  : ExtractPoint = CLng(Numbers(1))
            Case "PLAYERS"    : ExtractPoint = CLng(Numbers(2))
            Case "SPAWNERS"   : ExtractPoint = CLng(Numbers(3))
            Case "SCRIPTS"    : ExtractPoint = CLng(Numbers(4))
            Case Else         : Call VBGLAssign(ExtractPoint, TileInformation(NameOrIndex))
        End Select
    End If
End Function

Public Function TileInformation(ByVal Name As String) As Variant
    Select Case Name
        Case "Number"     : Call VBGLAssign(TileInformation, Tile.Number)
        Case "Name"       : Call VBGLAssign(TileInformation, Tile.Name)
        Case "TileType"   : Call VBGLAssign(TileInformation, Tile.TileType)
        Case "Speed"      : Call VBGLAssign(TileInformation, Tile.Speed)
        Case "Depth"      : Call VBGLAssign(TileInformation, Tile.Depth)
    End Select
End Function