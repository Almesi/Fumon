Attribute VB_Name = "GameGraphicsFumon"


Option Explicit

Private FumonList As VBGLTextBox

Public Sub SetUpFumonGraphics()
    Set FumonList = CreateFumonList
    FumonRenderObject.Inputt = CreateInput()
    Call FumonRenderObject.AddDrawable(FumonList)
End Sub

Public Sub UpdateFumon(ByVal Index As Long)
    Dim i As Long
    Dim Color() As Single
    ReDim Color(3)
    MePlayer.Fumons.Selected = Index
    FumonList.Fonts = UpdateTextBox(UsedFont)
    For i = 0 To MePlayer.Fumons.FumonCount
        FumonList.Font(i).FontColor = Color
    Next i
    Color(1) = 1
    FumonList.Font(Index).FontColor = Color
    Call FumonList.UpdateData()
End Sub

Private Function CreateInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Dim SelectedFumon As VBGLCallable
    Set SelectedFumon = VBGLCallable.Create(MePlayer.Fumons, "Selected", vbGet, -1)

    Call Temp.AddKeyUp(Asc("1")     , VBGLCallable.Create(Nothing         , "UpdateFumon"        , vbMethod, 0, 0))
    Call Temp.AddKeyUp(Asc("2")     , VBGLCallable.Create(Nothing         , "UpdateFumon"        , vbMethod, 0, 1))
    Call Temp.AddKeyUp(Asc("3")     , VBGLCallable.Create(Nothing         , "UpdateFumon"        , vbMethod, 0, 2))
    Call Temp.AddKeyUp(Asc("4")     , VBGLCallable.Create(Nothing         , "UpdateFumon"        , vbMethod, 0, 3))
    Call Temp.AddKeyUp(Asc("5")     , VBGLCallable.Create(Nothing         , "UpdateFumon"        , vbMethod, 0, 4))
    Call Temp.AddKeyUp(Asc("6")     , VBGLCallable.Create(Nothing         , "UpdateFumon"        , vbMethod, 0, 5))
    Call Temp.AddKeyUp(Asc("7")     , VBGLCallable.Create(Nothing         , "UpdateFumon"        , vbMethod, 0, 6))
    Call Temp.AddKeyUp(Asc("8")     , VBGLCallable.Create(Nothing         , "UpdateFumon"        , vbMethod, 0, 7))
    Call Temp.AddKeyUp(Asc(" ")     , VBGLCallable.Create(MePlayer        , "LetCurrentMove"     , vbMethod, 0, FightMove.FightMoveChangeFumon))
    Call Temp.AddKeyUp(Asc(" ")     , VBGLCallable.Create(MePlayer        , "LetCurrentValue"    , vbMethod, 0, SelectedFumon))
    Call Temp.AddKeyUp(Asc("0")     , VBGLCallable.Create(Nothing         , "UpdateFumon"        , vbMethod, 0, SelectedFumon))
    Call Temp.AddKeyUp(Asc(" ")     , VBGLCallable.Create(Nothing         , "RemoveRenderObject" , vbMethod, -1))
    Call Temp.AddKeyUp(27           , VBGLCallable.Create(Nothing         , "RemoveRenderObject" , vbMethod, -1))
    Call Temp.AddKeyUp(Asc("a")     , VBGLCallable.Create(Nothing         , "AddRenderObject"    , vbMethod, 0, AttackRenderObject))

    Set CreateInput = Temp
End Function

Private Function CreateFumonList() As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +0.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +0.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)
    Set CreateFumonList = VBGLTextBox.Create(Temp, UpdateTextBox(UsedFont))
End Function

Private Function UpdateTextBox(FontLayout As VBGLFontLayout) As VBGLFont()
    Dim Fonts() As VBGLFont
    Dim Text As String
    With MePlayer.Fumons
        ReDim Fonts(.FumonCount)
        Dim i As Long
        For i = 0 To .FumonCount
            Text = .Fumon(i).Definition.Name & " Lvl: " & .Fumon(i).GetLevel & vbCrLf
            Set Fonts(i) = VBGLFont.Create(Text, FontLayout)
        Next i
    End With
    UpdateTextBox = Fonts
End Function