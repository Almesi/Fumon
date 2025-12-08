Attribute VB_Name = "GameGraphicsInventory"


Option Explicit

Private InventoryList As VBGLList
Public UpdateInventory As std_Callable

Public Function SetUpInventoryGraphics() As VBGLRenderObject
    Dim Base As Items
    Set Base = MePlayer.FightBase.Items
    Dim X              As Single       : Let X              = -1.0!
    Dim Texture        As VBGLTexture  : Set Texture        = GameTextures.ObjectByName("Items")
    Dim TextObject     As std_Callable : Set TextObject     = CreateUnFixedCallable("$0.Item($1).FullName()", Base, 0)
    Dim NameObject     As std_Callable : Set NameObject     = CreateUnFixedCallable("$0.Item($1).Definition.Name()", Base, 0)
    Dim ColorObject    As std_Callable : Set ColorObject    = Nothing
    Set InventoryList = CreateList(Base, X, Texture, TextObject, NameObject, ColorObject)

    Dim UserInput As VBGLIInput
    Set UserInput = CreateInput(Base, Texture, TextObject, NameObject, ColorObject)

    Set SetUpInventoryGraphics = VBGLRenderObject.Create(UserInput, LeftSideFrame)
    Call SetUpInventoryGraphics.AddDrawable(InventoryList)
End Function

Private Function CreateInput(ByVal Base As Object, _
                             ByVal Texture As VBGLTexture, _
                             ByVal TextObject As std_Callable, _
                             ByVal NameObject As std_Callable, _
                             ByVal ColorObject As std_Callable _
                            ) As VBGLIInput

    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Dim GetSelected As std_Callable
    Set GetSelected = CreateFixedCallable("$0.Selected()", MePlayer.FightBase.Items)
    
    Set UpdateInventory = CreateFixedCallable("UpdateList($0, $1, $2, $3, $4, $5, $6)", InventoryList, Base, Texture, TextObject.SetAutoExecute(False), NameObject.SetAutoExecute(False), ColorObject, GetSelected)
    Call Temp.AddKeyUp(Asc("w") , CreateFixedCallable("$0.Selected(+1)", Base))
    Call Temp.AddKeyUp(Asc("s") , CreateFixedCallable("$0.Selected(-1)", Base))
    Call Temp.AddKeyUp(Asc("w") , UpdateInventory)
    Call Temp.AddKeyUp(Asc("s") , UpdateInventory)
    
    Dim ItemCallback As std_Callable
    Set ItemCallback = CreateFixedCallable("$0.SelectedItem()", MePlayer.FightBase.Items)
    Call Temp.AddKeyUp(Asc(" ") , std_Callable.Create(MePlayer.FightBase.CurrentMove , "Value", vbLet, 0).Bind(FightMove.FightMoveItem).FixArgs(True))
    Call Temp.AddKeyUp(Asc(" ") , std_Callable.Create(MePlayer.FightBase.CurrentValue, "Value", vbLet, 0).Bind(ItemCallback).FixArgs(True))
    Call Temp.AddKeyUp(Asc(" ") , CreateFixedCallable("RemoveDrawableFromRenderObject()"))
    Call Temp.AddKeyUp(27       , CreateFixedCallable("RemoveDrawableFromRenderObject()"))

    Set CreateInput = Temp
End Function