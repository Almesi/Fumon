Attribute VB_Name = "GameGraphicsMap"


Option Explicit

Public Function SetUpMapGraphics() As VBGLRenderObject
    Call TileSet.LookAt(MePlayer.MoveBase.Column.Value, MePlayer.MoveBase.Row.Value, MeGameMap.Columns.Value, MeGameMap.Rows.Value)

    Set SetUpMapGraphics = VBGLRenderObject.Create(CreateInput(), CurrentContext.CurrentFrame())
    Call SetUpMapGraphics.AddDrawable(TileSet)
    Call SetUpMapGraphics.AddDrawable(PlayerPositions)
End Function

Private Function CreateInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Dim TempCol As std_Callable: Set TempCol = std_Callable.Create(MePlayer.MoveBase.Column, "Value", vbGet, -1)
    Dim TempRow As std_Callable: Set TempRow = std_Callable.Create(MePlayer.MoveBase.Row, "Value", vbGet, -1)

    Call Temp.AddKeyUp(27, CreateFixedCallable("RemoveRenderObject()"))
    Call Temp.AddKeyUp(27, std_Callable.Create(TileSet, "LookAt", vbMethod, 3).Bind(TempCol, TempRow, ScreenSpriteX, ScreenSpriteY).FixArgs(True))
    Set CreateInput = Temp
End Function