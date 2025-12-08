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

    Call Temp.AddKey(Asc("w")    , std_Callable.Create(MePlayer.MoveBase, "Move", vbMethod, 1).Bind(+0, -1).FixArgs(True))
    Call Temp.AddKey(Asc("a")    , std_Callable.Create(MePlayer.MoveBase, "Move", vbMethod, 1).Bind(-1, +0).FixArgs(True))
    Call Temp.AddKey(Asc("s")    , std_Callable.Create(MePlayer.MoveBase, "Move", vbMethod, 1).Bind(+0, +1).FixArgs(True))
    Call Temp.AddKey(Asc("d")    , std_Callable.Create(MePlayer.MoveBase, "Move", vbMethod, 1).Bind(+1, +0).FixArgs(True))

    Call Temp.AddKey(Asc("W")    , std_Callable.Create(MePlayer.MoveBase, "Look", vbMethod, 0).Bind(xlUp).FixArgs(True))
    Call Temp.AddKey(Asc("A")    , std_Callable.Create(MePlayer.MoveBase, "Look", vbMethod, 0).Bind(xlLeft).FixArgs(True))
    Call Temp.AddKey(Asc("S")    , std_Callable.Create(MePlayer.MoveBase, "Look", vbMethod, 0).Bind(xlDown).FixArgs(True))
    Call Temp.AddKey(Asc("D")    , std_Callable.Create(MePlayer.MoveBase, "Look", vbMethod, 0).Bind(xlRight).FixArgs(True))

    Call Temp.AddKeyUp(Asc(" ")  , std_Callable.Create(MePlayer.MoveBase, "Interact", VbMethod, 1).Bind(MePlayer, 1).FixArgs(True))

    Dim MaxX As Long : MaxX = MeGameMap.Columns.Value + 1
    Dim MaxY As Long : MaxY = MeGameMap.Rows.Value + 1
    Call Temp.AddKeyUp(Asc("m")    , CreateFixedCallable("AddRenderObject($0)", MapRenderObject))
    Call Temp.AddKeyUp(Asc("m")    , std_Callable.Create(TileSet, "LookAt", vbMethod, 3).Bind(MaxX / 2, MaxY / 2, MaxX, MaxY).FixArgs(True))

    Call Temp.AddKeyUp(Asc("i")    , CreateFixedCallable("AddDrawableToRenderObject($0, $1)", InventoryRenderObject, InventoryRenderObject.UserInput))
    Call Temp.AddKeyUp(Asc("f")    , CreateFixedCallable("AddDrawableToRenderObject($0, $1)", FumonRenderObject, FumonRenderObject.UserInput))

    Call Temp.AddKeyUp(Asc("i")    , UpdateInventory)
    Call Temp.AddKeyUp(Asc("f")    , UpdateFumon)
    
    Call Temp.AddKeyUp(27          , CreateFixedCallable("RemoveRenderObject()"))
    Set CreateInput = Temp
End Function