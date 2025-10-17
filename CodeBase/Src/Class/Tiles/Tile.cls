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

Private p_Tile              As TileDefinition
Private p_Range             As Range

Public Property Let Tile(n_Tile      As TileDefinition) : Set p_Tile    = n_Tile   : End Property
Public Property Let Range(n_Range    As Range)          : Set p_Range   = n_Range  : End Property

Public Property Get Tile()           As TileDefinition  : Set Tile      = p_Tile   : End Property
Public Property Get Range()          As Range           : Set Range     = p_Range  : End Property

Public Function Create(Rng As Range) As Tile
    Dim Index As Long
    Set Create = New Tile
    With Create
        .Range      = Rng
        Index = CLng(.ExtractPoint("Tiles"))
        If Index <> -1 Then .Tile = FactoryServer.Tile(Index)
    End With
End Function

Public Function Player() As IPlayer
    Dim Index As Long
    Index = CLng(ExtractPoint("Players"))
    If Index <> -1 Then Set Player = FactoryServer.Player(Index)
End Function
Public Function Script() As Script
    Dim Index As Long
    Index = CLng(ExtractPoint("Scripts"))
    If Index <> -1 Then Script = FactoryServer.Script(Index)
End Function

Public Sub SetMapPoint(Optional ByVal Tile As String = "KEEP", Optional ByVal Player As String = "KEEP", Optional ByVal Script As String = "KEEP")
    Dim Values() As String
    Values = Split(Range.Value, ", ")
    If Tile   <> "KEEP" Then Values(0) = Tile
    If Player <> "KEEP" Then Values(1) = Player
    If Script <> "KEEP" Then Values(2) = Script
    Range.Value = Values(0) & ", " & Values(1) & ", " & Values(2)
End Sub

Public Sub SetMapPointTile(ByVal   Tile   As Long) : Call SetMapPoint(Tile  , "KEEP", "KEEP") : End Sub
Public Sub SetMapPointPlayer(ByVal Player As Long) : Call SetMapPoint("KEEP", Player, "KEEP") : End Sub
Public Sub SetMapPointScript(ByVal Script As Long) : Call SetMapPoint("KEEP", "KEEP", Script) : End Sub

Public Function ExtractPoint(NameOrIndex As String) As Variant
    Dim Numbers() As String
    Numbers = Split(Range.Value, ", ")
    If IsNumeric(NameOrIndex) Then
        ExtractPoint = Numbers(CLng(NameOrIndex))
    Else
        Select Case UCase(NameOrIndex)
            Case "TILES"   : ExtractPoint = CLng(Numbers(0))
            Case "PLAYERS" : ExtractPoint = CLng(Numbers(1))
            Case "SCRIPTS" : ExtractPoint = CLng(Numbers(2))
            Case Else      : Call VBGLAssign(ExtractPoint, TileInformation(NameOrIndex))
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