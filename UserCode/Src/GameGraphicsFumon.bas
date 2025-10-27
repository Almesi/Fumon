Attribute VB_Name = "GameGraphicsFumon"


Option Explicit

Private FumonList As VBGLTextBox

Public Function SetUpFumonGraphics() As VBGLRenderObject
    Set FumonList = CreateFumonList()
    Set SetUpFumonGraphics = VBGLRenderObject.Create(CreateInput(), CurrentContext.CurrentFrame())
    Call SetUpFumonGraphics.AddDrawable(FumonList)
End Function

Public Sub UpdateFumon(ByVal Index As Long)
    Dim i As Long
    Dim Color() As Single
    ReDim Color(3)
    MeFighter.FightBase.Fumons.Selected = Index
    FumonList.Fonts = UpdateTextBox(UsedFont)
    For i = 0 To MeFighter.FightBase.Fumons.FumonCount
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
    Set SelectedFumon = ConvertCallable("$0.Selected()", MeFighter.FightBase.Fumons)

    Call Temp.AddKeyUp(Asc("1")     , ConvertCallable("UpdateFumon(0)"))
    Call Temp.AddKeyUp(Asc("2")     , ConvertCallable("UpdateFumon(1)"))
    Call Temp.AddKeyUp(Asc("3")     , ConvertCallable("UpdateFumon(2)"))
    Call Temp.AddKeyUp(Asc("4")     , ConvertCallable("UpdateFumon(3)"))
    Call Temp.AddKeyUp(Asc("5")     , ConvertCallable("UpdateFumon(4)"))
    Call Temp.AddKeyUp(Asc("6")     , ConvertCallable("UpdateFumon(5)"))
    Call Temp.AddKeyUp(Asc("7")     , ConvertCallable("UpdateFumon(6)"))
    Call Temp.AddKeyUp(Asc("8")     , ConvertCallable("UpdateFumon(7)"))
    Call Temp.AddKeyUp(Asc(" ")     , ConvertCallable("$0.LetCurrentMove($1)"  , MeFighter.FightBase, FightMove.FightMoveChangeFumon))
    Call Temp.AddKeyUp(Asc(" ")     , ConvertCallable("$0.LetCurrentValue($1)" , MeFighter.FightBase, SelectedFumon))
    Call Temp.AddKeyUp(Asc("0")     , ConvertCallable("UpdateFumon($0)"        , SelectedFumon))
    Call Temp.AddKeyUp(Asc(" ")     , ConvertCallable("RemoveRenderObject()"))
    Call Temp.AddKeyUp(27           , ConvertCallable("RemoveRenderObject()"))
    Call Temp.AddKeyUp(Asc("a")     , ConvertCallable("AddRenderObject($0)"    , AttackRenderObject))

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
    Set CreateFumonList = FactoryTextBox.Create(Temp, UpdateTextBox(UsedFont))
End Function

Private Function UpdateTextBox(FontLayout As VBGLFontLayout) As VBGLFont()
    Dim Fonts() As VBGLFont
    Dim Text As String
    With MeFighter.FightBase.Fumons
        ReDim Fonts(.FumonCount)
        Dim i As Long
        For i = 0 To .FumonCount
            Text = .Fumon(i).Definition.Name & " Lvl: " & .Fumon(i).GetLevel & vbCrLf
            Set Fonts(i) = VBGLFont.Create(Text, FontLayout)
        Next i
    End With
    UpdateTextBox = Fonts
End Function