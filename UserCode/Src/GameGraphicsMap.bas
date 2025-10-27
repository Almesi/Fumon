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

    Dim TempCol As VBGLCallable: Set TempCol = VBGLCallable.Create(MePlayerHuman.MoveBase.Column, "Value", vbGet, -1)
    Dim TempRow As VBGLCallable: Set TempRow = VBGLCallable.Create(MePlayerHuman.MoveBase.Row, "Value", vbGet, -1)

    Call Temp.AddKeyUp(27, ConvertCallable("RemoveRenderObject()"))
    Call Temp.AddKeyUp(27, VBGLCallable.Create(TileSet, "LookAt", vbMethod, 3, TempCol, TempRow, ScreenSpriteX, ScreenSpriteY))
    Set CreateInput = Temp
End Function