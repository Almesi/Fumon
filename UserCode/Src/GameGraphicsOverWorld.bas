Attribute VB_Name = "GameGraphicsOverWorld"


Option Explicit

Public Const ScreenSpriteX As Long = 17
Public Const ScreenSpriteY As Long = 09

Public Function SetUpOverWorldGraphics() As VBGLRenderObject
    Call TileSet.LookAt(MePlayer.MoveBase.Column.Value, MePlayer.MoveBase.Row.Value, ScreenSpriteX, ScreenSpriteY)

    Set SetUpOverWorldGraphics = VBGLRenderObject.Create(CreateInput(), CurrentContext.CurrentFrame())
    Call SetUpOverWorldGraphics.AddDrawable(TileSet)
    Call SetUpOverWorldGraphics.AddDrawable(PlayerPositions)
End Function

Public Sub UpdateOverWorld(ByVal PlayerMoved As Boolean)
    Call TileSet.LookAt(MePlayer.MoveBase.Column.Value, MePlayer.MoveBase.Row.Value, ScreenSpriteX, ScreenSpriteY)
    If PlayerMoved Then
        Call PlayerPositions.VAO.Buffer.Update(UpdateMapData())
    End If
End Sub

Private Function CreateInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Call Temp.AddKey(Asc("w")    , VBGLCallable.Create(MePlayer, "Move", vbMethod, 1, +0, -1))
    Call Temp.AddKey(Asc("a")    , VBGLCallable.Create(MePlayer, "Move", vbMethod, 1, -1, +0))
    Call Temp.AddKey(Asc("s")    , VBGLCallable.Create(MePlayer, "Move", vbMethod, 1, +0, +1))
    Call Temp.AddKey(Asc("d")    , VBGLCallable.Create(MePlayer, "Move", vbMethod, 1, +1, +0))
    Call Temp.AddKey(Asc("w")    , ConvertCallable("UpdateOverWorld(True)"))
    Call Temp.AddKey(Asc("a")    , ConvertCallable("UpdateOverWorld(True)"))
    Call Temp.AddKey(Asc("s")    , ConvertCallable("UpdateOverWorld(True)"))
    Call Temp.AddKey(Asc("d")    , ConvertCallable("UpdateOverWorld(True)"))

    Call Temp.AddKey(Asc("W")    , VBGLCallable.Create(MePlayer, "Look", vbMethod, 0, xlUp))
    Call Temp.AddKey(Asc("A")    , VBGLCallable.Create(MePlayer, "Look", vbMethod, 0, xlLeft))
    Call Temp.AddKey(Asc("S")    , VBGLCallable.Create(MePlayer, "Look", vbMethod, 0, xlDown))
    Call Temp.AddKey(Asc("D")    , VBGLCallable.Create(MePlayer, "Look", vbMethod, 0, xlRight))
    Call Temp.AddKey(Asc("W")    , ConvertCallable("UpdateOverWorld(True)"))
    Call Temp.AddKey(Asc("A")    , ConvertCallable("UpdateOverWorld(True)"))
    Call Temp.AddKey(Asc("S")    , ConvertCallable("UpdateOverWorld(True)"))
    Call Temp.AddKey(Asc("D")    , ConvertCallable("UpdateOverWorld(True)"))

    Call Temp.AddKeyUp(Asc(" ")  , VBGLCallable.Create(MePlayer, "Interact", vbMethod, 0, 1))

    Dim MaxX As Long : MaxX = MeGameMap.Columns.Value + 1
    Dim MaxY As Long : MaxY = MeGameMap.Rows.Value + 1
    Call Temp.AddKeyUp(Asc("m")    , ConvertCallable("AddRenderObject($0)", MapRenderObject))
    Call Temp.AddKeyUp(Asc("m")    , VBGLCallable.Create(TileSet, "LookAt", vbMethod, 3, MaxX / 2, MaxY / 2, MaxX, MaxY))

    Call Temp.AddKeyUp(Asc("i")    , ConvertCallable("AddRenderObject($0)", InventoryRenderObject))
    Call Temp.AddKeyUp(Asc("f")    , ConvertCallable("AddRenderObject($0)", FumonRenderObject))
    Call Temp.AddKeyUp(27          , ConvertCallable("RemoveRenderObject()"))
    Set CreateInput = Temp
End Function