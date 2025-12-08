Attribute VB_Name = "GameGraphicsFumon"


Option Explicit

Private FumonList As VBGLList
Public UpdateFumon As std_Callable

Public Function SetUpFumonGraphics() As VBGLRenderObject
    Dim Base As Fumons
    Set Base = MePlayer.FightBase.Fumons
    Dim X              As Single       : Let X              = -1.0!
    Dim Texture        As VBGLTexture  : Set Texture        = GameTextures.ObjectByName("Fumons")
    Dim TextObject     As std_Callable : Set TextObject     = CreateUnFixedCallable("$0.Fumon($1).FullName()", Base, 0)
    Dim NameObject     As std_Callable : Set NameObject     = CreateUnFixedCallable("$0.Fumon($1).FrontName()", Base, 0)
    ' For name add front or back
    Dim ColorObject    As std_Callable : Set ColorObject    = Nothing
    Set FumonList = CreateList(Base, X, Texture, TextObject, NameObject, ColorObject)

    Dim UserInput As VBGLIInput
    Set UserInput = CreateInput(Base, Texture, TextObject, NameObject, ColorObject)

    Set SetUpFumonGraphics = VBGLRenderObject.Create(UserInput, LeftSideFrame)
    Call SetUpFumonGraphics.AddDrawable(FumonList)
End Function

Private Function CreateInput(ByVal Base As Object, _
                             ByVal Texture As VBGLTexture, _
                             ByVal TextObject As std_Callable, _
                             ByVal NameObject As std_Callable, _
                             ByVal ColorObject As std_Callable _
                            ) As VBGLIInput

    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Dim SelectedFumon As std_Callable
    Set SelectedFumon = CreateFixedCallable("$0.Selected()", Base)

    Set UpdateFumon = CreateFixedCallable("UpdateList($0, $1, $2, $3, $4, $5, $6)", FumonList, Base, Texture, TextObject.SetAutoExecute(False), NameObject.SetAutoExecute(False), ColorObject, SelectedFumon)

    With MePlayer.FightBase
        Call Temp.AddKeyUp(Asc("1") , CreateFixedCallable("$0.Selected(0)", .Fumons))
        Call Temp.AddKeyUp(Asc("2") , CreateFixedCallable("$0.Selected(1)", .Fumons))
        Call Temp.AddKeyUp(Asc("3") , CreateFixedCallable("$0.Selected(2)", .Fumons))
        Call Temp.AddKeyUp(Asc("4") , CreateFixedCallable("$0.Selected(3)", .Fumons))
        Call Temp.AddKeyUp(Asc("5") , CreateFixedCallable("$0.Selected(4)", .Fumons))
        Call Temp.AddKeyUp(Asc("6") , CreateFixedCallable("$0.Selected(5)", .Fumons))
        Call Temp.AddKeyUp(Asc("7") , CreateFixedCallable("$0.Selected(6)", .Fumons))
        Call Temp.AddKeyUp(Asc("8") , CreateFixedCallable("$0.Selected(7)", .Fumons))
    End With

    Call Temp.AddKeyUp(Asc("1")     , UpdateFumon)
    Call Temp.AddKeyUp(Asc("2")     , UpdateFumon)
    Call Temp.AddKeyUp(Asc("3")     , UpdateFumon)
    Call Temp.AddKeyUp(Asc("4")     , UpdateFumon)
    Call Temp.AddKeyUp(Asc("5")     , UpdateFumon)
    Call Temp.AddKeyUp(Asc("6")     , UpdateFumon)
    Call Temp.AddKeyUp(Asc("7")     , UpdateFumon)
    Call Temp.AddKeyUp(Asc("8")     , UpdateFumon)
    Call Temp.AddKeyUp(Asc("s")     , CreateFixedCallable("$0.Swap(0, $1)", Base, SelectedFumon))
    Call Temp.AddKeyUp(Asc("s")     , UpdateFumon)
    Call Temp.AddKeyUp(Asc("s")     , UpdateAttack)


    Call Temp.AddKeyUp(Asc(" ")     , std_Callable.Create(MePlayer.FightBase.CurrentMove , "Value", vbLet, 0).Bind(FightMove.FightMoveChangeFumon).FixArgs(True))
    Call Temp.AddKeyUp(Asc(" ")     , std_Callable.Create(MePlayer.FightBase.CurrentValue, "Value", vbLet, 0).Bind(SelectedFumon).FixArgs(True))
    Call Temp.AddKeyUp(Asc(" ")     , CreateFixedCallable("RemoveDrawableFromRenderObject()"))
    Call Temp.AddKeyUp(Asc(" ")     , UpdateAttack)
    Call Temp.AddKeyUp(27           , CreateFixedCallable("RemoveDrawableFromRenderObject()"))
    Call Temp.AddKeyUp(Asc("a")     , CreateFixedCallable("AddDrawableToRenderObject($0, $1)", AttackRenderObject, AttackRenderObject.UserInput))
    Call Temp.AddKeyUp(Asc("a")     , UpdateAttack)

    Set CreateInput = Temp
End Function