Attribute VB_Name = "GameGraphicsAttack"


Option Explicit

Private AttackList As VBGLList
Public UpdateAttack As std_Callable

Public Function SetUpAttackGraphics() As VBGLRenderObject
    Dim X              As Single       : Let X              = 0
    Dim Base           As std_Callable : Set Base           = CreateFixedCallable("$0.SelectedFumon.Attacks()", MePlayer.FightBase.Fumons)
    Dim Texture        As VBGLTexture  : Set Texture        = GameTextures.ObjectByName("Attacks")
    Dim TextObject     As std_Callable : Set TextObject     = CreateUnFixedCallable("$0.Attack($1).FullName()", Base, 0)
    Dim NameObject     As std_Callable : Set NameObject     = CreateUnFixedCallable("$0.Attack($1).Name()", Base, 0)
    Dim ColorObject    As std_Callable : Set ColorObject    = CreateUnFixedCallable("$0.Color($1.Attack($2).ElementType())", MeServer.ElementTypes, Base, 0)
    Set AttackList = CreateList(Base.Run, X, Texture, TextObject, NameObject, ColorObject)

    Dim UserInput As VBGLIInput
    Set UserInput = CreateInput(Base, Texture, TextObject, NameObject, ColorObject)

    Set SetUpAttackGraphics = VBGLRenderObject.Create(UserInput, LeftSideFrame)
    Call SetUpAttackGraphics.AddDrawable(AttackList)
End Function

Private Function CreateInput(ByVal Base As std_Callable, _
                             ByVal Texture As VBGLTexture, _
                             ByVal TextObject As std_Callable, _
                             ByVal NameObject As std_Callable, _
                             ByVal ColorObject As std_Callable _
                            ) As VBGLIInput
                             
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Dim GetSelectedAttack As std_Callable
    Set GetSelectedAttack = CreateFixedCallable("$0.Selected()", Base)


    Set UpdateAttack = CreateFixedCallable("UpdateList($0, $1, $2, $3, $4, $5, $6)", AttackList, Base, Texture, TextObject.SetAutoExecute(False), NameObject.SetAutoExecute(False), ColorObject.SetAutoExecute(False), GetSelectedAttack)

    Call Temp.AddKeyUp(Asc("1") , CreateFixedCallable("$0.Selected(0)", Base))
    Call Temp.AddKeyUp(Asc("2") , CreateFixedCallable("$0.Selected(1)", Base))
    Call Temp.AddKeyUp(Asc("3") , CreateFixedCallable("$0.Selected(2)", Base))
    Call Temp.AddKeyUp(Asc("4") , CreateFixedCallable("$0.Selected(3)", Base))
    Call Temp.AddKeyUp(Asc("1") , UpdateAttack)
    Call Temp.AddKeyUp(Asc("2") , UpdateAttack)
    Call Temp.AddKeyUp(Asc("3") , UpdateAttack)
    Call Temp.AddKeyUp(Asc("4") , UpdateAttack)

    Call Temp.AddKeyUp(Asc(" ") , std_Callable.Create(MePlayer.FightBase.CurrentMove , "Value", vbLet, 0).Bind(FightMove.FightMoveAttack).FixArgs(True))
    Call Temp.AddKeyUp(Asc(" ") , std_Callable.Create(MePlayer.FightBase.CurrentValue, "Value", vbLet, 0).Bind(GetSelectedAttack).FixArgs(True))
    Call Temp.AddKeyUp(Asc(" ") , CreateFixedCallable("RemoveDrawableFromRenderObject()"))
    Call Temp.AddKeyUp(27       , CreateFixedCallable("RemoveDrawableFromRenderObject()"))
    Set CreateInput = Temp
End Function