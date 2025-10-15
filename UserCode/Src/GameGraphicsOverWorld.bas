Attribute VB_Name = "GameGraphicsOverWorld"


Option Explicit

Public  TileSet As VBGLDualGrid
Public  PlayerPositions As VBGLMesh

Public Const ScreenSpriteX As Long = 17
Public Const ScreenSpriteY As Long = 09

Public Sub SetUpOverWorldGraphics()
    Dim Layout As VBGLLayout
    Set Layout = VBGLPrCoLayoutXYZTxTy

    Dim Tiles() As Long
    Tiles = MeServer.GameMap.Tiles.TileData("Tiles")
    Set TileSet = VBGLDualGrid.Create(MeServer.GameMap.Folder.Value & "\Sprites\Tiles.png", 64, 64, ScreenSpriteX, ScreenSpriteY)
    Call TileSet.SetUp(Layout, Tiles)

    Dim z() As Single
    z = MeServer.GameMap.Tiles.TileData("Depth")
    Call TileSet.ParseData(0, TileSet.GetPositionData(z))
    Call TileSet.ParseData(1, TileSet.GetSubTextureData())
    Call TileSet.Build()

    Dim Players() As Long
    Players = MeServer.GameMap.Tiles.TileData("Players")
    Dim Data   As IDataSingle : Set Data   = GetPlayerData(MeServer.GetTexture("Players"), Players)

    Dim Shader As VBGLShader
    Set Shader = TileSet.Mesh.Shader
    Set PlayerPositions = VBGLMesh.Create(Shader, Layout, Data)
    Call PlayerPositions.AddTexture(MeServer.GetTexture("Players"))


    Call TileSet.LookAt(MePlayerHuman.Column.Value, MePlayerHuman.Row.Value, ScreenSpriteX, ScreenSpriteY)

    Call OverWorldRenderObject.AddDrawable(TileSet)
    Call OverWorldRenderObject.AddDrawable(PlayerPositions)
    OverWorldRenderObject.Inputt = CreateInput()
End Sub

Public Sub UpdateOverWorld(ByVal PlayerMoved As Boolean)
    Call TileSet.LookAt(MePlayerHuman.Column.Value, MePlayerHuman.Row.Value, ScreenSpriteX, ScreenSpriteY)
    If PlayerMoved Then
        Call PlayerPositions.VAO.Buffer.Update(UpdateData())
    End If
End Sub

Private Function CreateInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Call Temp.AddKey(Asc("w")    , VBGLCallable.Create(MePlayerHuman , "Move"               , vbMethod, 1, +0, -1))
    Call Temp.AddKey(Asc("a")    , VBGLCallable.Create(MePlayerHuman , "Move"               , vbMethod, 1, -1, +0))
    Call Temp.AddKey(Asc("s")    , VBGLCallable.Create(MePlayerHuman , "Move"               , vbMethod, 1, +0, +1))
    Call Temp.AddKey(Asc("d")    , VBGLCallable.Create(MePlayerHuman , "Move"               , vbMethod, 1, +1, +0))
    Call Temp.AddKey(Asc("w")    , VBGLCallable.Create(Nothing       , "UpdateOverWorld"    , vbMethod, 0, True))
    Call Temp.AddKey(Asc("a")    , VBGLCallable.Create(Nothing       , "UpdateOverWorld"    , vbMethod, 0, True))
    Call Temp.AddKey(Asc("s")    , VBGLCallable.Create(Nothing       , "UpdateOverWorld"    , vbMethod, 0, True))
    Call Temp.AddKey(Asc("d")    , VBGLCallable.Create(Nothing       , "UpdateOverWorld"    , vbMethod, 0, True))

    Call Temp.AddKey(Asc("W")    , VBGLCallable.Create(MePlayerHuman , "Look"               , vbMethod, 0, xlUp))
    Call Temp.AddKey(Asc("A")    , VBGLCallable.Create(MePlayerHuman , "Look"               , vbMethod, 0, xlLeft))
    Call Temp.AddKey(Asc("S")    , VBGLCallable.Create(MePlayerHuman , "Look"               , vbMethod, 0, xlDown))
    Call Temp.AddKey(Asc("D")    , VBGLCallable.Create(MePlayerHuman , "Look"               , vbMethod, 0, xlRight))
    Call Temp.AddKey(Asc("W")    , VBGLCallable.Create(Nothing       , "UpdateOverWorld"    , vbMethod, 0, True))
    Call Temp.AddKey(Asc("A")    , VBGLCallable.Create(Nothing       , "UpdateOverWorld"    , vbMethod, 0, True))
    Call Temp.AddKey(Asc("S")    , VBGLCallable.Create(Nothing       , "UpdateOverWorld"    , vbMethod, 0, True))
    Call Temp.AddKey(Asc("D")    , VBGLCallable.Create(Nothing       , "UpdateOverWorld"    , vbMethod, 0, True))

    Call Temp.AddKeyUp(Asc(" ")    , VBGLCallable.Create(MePlayerHuman , "Interact"           , vbMethod, 0, 1))

    Dim MaxX As Long : MaxX = MeServer.GameMap.Columns.Value + 1
    Dim MaxY As Long : MaxY = MeServer.GameMap.Rows.Value + 1
    Call Temp.AddKeyUp(Asc("m")    , VBGLCallable.Create(Nothing  , "AddRenderObject"    , vbMethod, 0, MapRenderObject))
    Call Temp.AddKeyUp(Asc("m")    , VBGLCallable.Create(TileSet  , "LookAt"             , vbMethod,  3, MaxX / 2, MaxY / 2, MaxX, MaxY))

    Call Temp.AddKeyUp(Asc("i")    , VBGLCallable.Create(Nothing  , "AddRenderObject"    , vbMethod, 0, InventoryRenderObject))
    Call Temp.AddKeyUp(Asc("f")    , VBGLCallable.Create(Nothing  , "AddRenderObject"    , vbMethod, 0, FumonRenderObject))
    Call Temp.AddKeyUp(27          , VBGLCallable.Create(Nothing  , "RemoveRenderObject" , vbMethod, -1 ))
    Set CreateInput = Temp
End Function

Private Function UpdateData() As IDataSingle
    Dim Players() As Long
    Players = MeServer.GameMap.Tiles.TileData("Players")
    Set UpdateData   = GetPlayerData(MeServer.GetTexture("Players"), Players)
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
                If TypeName(Player) = "HumanPlayer" Then
                    Dim TempH As HumanPlayer
                    Set TempH = Player
                    SpriteIndex = PlayerIndex(Index, TempH.LookDirection.Value)
                Else
                    Dim TempC As ComPlayer
                    Set TempC = Player
                    SpriteIndex = PlayerIndex(Index, TempC.LookDirection.Value)
                End If
                Call AddTriangles(ReturnArr, x, y, 0.5, Texture.SubTexture(SpriteIndex), ScreenSpriteX, ScreenSpriteY)
            End If
        Next x
    Next y

    Set GetPlayerData = VBGLData.CreateSingle(ReturnArr)
End Function

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