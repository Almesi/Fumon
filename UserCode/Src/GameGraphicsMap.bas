Attribute VB_Name = "GameGraphicsMap"


Option Explicit

Public Sub SetUpMapGraphics()
    Call MapRenderObject.AddDrawable(TileSet)
    Call MapRenderObject.AddDrawable(PlayerPositions)
    MapRenderObject.Inputt = CreateInput()
End Sub

Private Function CreateInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Call Temp.AddKeyUp(27, VBGLCallable.Create(Nothing  , "RemoveRenderObject" , vbMethod, -1))
    Call Temp.AddKeyUp(27, VBGLCallable.Create(TileSet  , "LookAt"             , vbMethod,  3, MePlayerHuman.Column.Value, MePlayerHuman.Row.Value, ScreenSpriteX, ScreenSpriteY))
    Set CreateInput = Temp
End Function