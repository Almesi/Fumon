Attribute VB_Name = "GameGraphicsTileSet"


Option Explicit

Public  TileSet As VBGLDualGrid
Public  PlayerPositions As VBGLMesh


Public Sub SetUpTileSet()
    Dim Layout As VBGLLayout
    Set Layout = VBGLPrCoLayoutXYZTxTy

    Dim Tiles() As Long
    Tiles = MeGameMap.Tiles.TileData("Tiles")

    Set TileSet = VBGLDualGrid.CreateFromFolder(MeServer.MapData.Folder.Value & "\Sprites\Tiles", False)
    Call TileSet.SetUp(Layout, Tiles)
    Call SortTileSet()


    Dim z() As Single
    z = MeGameMap.Tiles.TileData("Depth")
    Call TileSet.ParseData(0, TileSet.GetPositionData(z))
    Call TileSet.ParseData(1, TileSet.GetSubTextureData())
    Call TileSet.Build()

    Dim Players() As Long
    Players = MeGameMap.Tiles.TileData("Players")
    Dim Data   As IDataSingle : Set Data   = GetPlayerData(GameTextures.ObjectByName("Players"), Players)

    Dim Shader As VBGLShader
    Set Shader = TileSet.Mesh.Shader
    Set PlayerPositions = VBGLMesh.Create(Shader, Layout, Data)
    Call PlayerPositions.AddTexture(GameTextures.ObjectByName("Players"))
End Sub

Public Function UpdateMapData() As IDataSingle
    Dim Players() As Long
    Players = MeGameMap.Tiles.TileData("Players")
    Set UpdateMapData   = GetPlayerData(GameTextures.ObjectByName("Players"), Players)
End Function

Private Function GetPlayerData(ByVal Texture As VBGLTexture, ByRef MapData() As Long) As IDataSingle
    Dim ReturnArr() As Single
    Dim x As Long, y As Long
    Dim Index As Long


    For y = 0 To USize(MapData, 1)
        For x = 0 To USize(MapData, 2)
            Index = MapData(y, x)
            If Index <> -1 Then
                Dim SpriteIndex As Long
                Dim Player As IPlayer
                Set Player = MeServer.Player(Index)
                SpriteIndex = PlayerIndex(Index, Player.MoveBase.LookDirection.Value)
                Call AddTriangles(ReturnArr, x, y, 0.5, Texture.SubTexture(SpriteIndex), ScreenSpriteX, ScreenSpriteY)
            End If
        Next x
    Next y

    Set GetPlayerData = VBGLData.CreateSingle(ReturnArr)
End Function

Private Sub SortTileSet()
    Dim NamesVar   As Variant : NamesVar = MeServer.Tiles.Properties("Name")
    Dim SubNames() As String  : SubNames = TileSet.TypeTypeString()
    Dim Names() As String
    Dim i As Long
    ReDim Names(USize(NamesVar))
    For i = 0 To USize(NamesVar)
        Names(i) = CStr(NamesVar(i))
    Next i

    Dim Manager As VBGLTextureManager
    Set Manager = New VBGLTextureManager
    Call Manager.SortByArrayFamily(TileSet.TileSet, Names, SubNames)
End Sub

Private Sub AddTriangles(ByRef Arr() As Single, ByVal x As Long, ByVal y As Long, ByVal z As Single, ByVal SubTexture As VBGLSubTexture, ByVal MaxX As Long, ByVal MaxY As Long)
    Dim VertexSize As Long: VertexSize = 5
    Dim x1 As Single : x1 = x
    Dim x2 As Single : x2 = x + 1
    Dim y1 As Single : y1 = y
    Dim y2 As Single : y2 = y + 1
    With SubTexture
        Call AddVertex(Arr, x1, y1, z, .GetX("TopLeft")    , .GetY("TopLeft"))
        Call AddVertex(Arr, x2, y1, z, .GetX("TopRight")   , .GetY("TopRight"))
        Call AddVertex(Arr, x1, y2, z, .GetX("BottomLeft") , .GetY("BottomLeft"))
        Call AddVertex(Arr, x2, y1, z, .GetX("TopRight")   , .GetY("TopRight"))
        Call AddVertex(Arr, x2, y2, z, .GetX("BottomRight"), .GetY("BottomRight"))
        Call AddVertex(Arr, x1, y2, z, .GetX("BottomLeft") , .GetY("BottomLeft"))
    End With
End Sub

Private Sub AddVertex(ByRef Arr() As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal Tx As Single, ByVal Ty As Single)
    Dim Temp() As Single
    ReDim Temp(4)
    Temp(0) = x
    Temp(1) = -y
    Temp(2) = z
    Temp(3) = Tx
    Temp(4) = Ty
    Call VBGLMerge(Arr, Temp)
End Sub

Private Function PlayerIndex(ByVal Index As Long, ByVal Direction As xlDirection) As Long
    Dim SpritesPerPlayer As Long
    SpritesPerPlayer = 6
    Select Case Direction
        Case xlUp    : PlayerIndex = Index * SpritesPerPlayer + 0
        Case xlLeft  : PlayerIndex = Index * SpritesPerPlayer + 1
        Case xlDown  : PlayerIndex = Index * SpritesPerPlayer + 2
        Case xlRight : PlayerIndex = Index * SpritesPerPlayer + 3
    End Select
End Function