Attribute VB_Name = "GameGraphicsInventory"


Option Explicit

Private InventoryList As VBGLTextBox

Public Function SetUpInventoryGraphics() As VBGLRenderObject
    Set InventoryList = CreateInventoryList()
    Set SetUpInventoryGraphics = VBGLRenderObject.Create(CreateInput(), CurrentContext.CurrentFrame())
    Call SetUpInventoryGraphics.AddDrawable(InventoryList)
End Function

Public Sub UpdateInventory(ByVal Offset As Long)
    Dim i As Long
    Dim Color() As Single
    Dim Index As Long
    Call MeFighter.FightBase.Items.Increment(Offset)
    Index = MeFighter.FightBase.Items.Selected
    ReDim Color(2)
    InventoryList.Fonts = UpdateTextBox(UsedFont)
    For i = 0 To MeFighter.FightBase.Items.Count
        InventoryList.Font(i).FontColor = Color
    Next i
    Color(1) = 1
    InventoryList.Font(Index).FontColor = Color
    Call InventoryList.UpdateData()
End Sub

Private Function CreateInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Dim ItemCallback As VBGLCallable
    Set ItemCallback = ConvertCallable("$0.SelectedItem()", MeFighter.FightBase.Items)

    Call Temp.AddKeyUp(Asc("w") , ConvertCallable("UpdateInventory(+1)"))
    Call Temp.AddKeyUp(Asc("s") , ConvertCallable("UpdateInventory(-1)"))
    Call Temp.AddKeyUp(Asc(" ") , ConvertCallable("$0.LetCurrentMove($1)"  , MeFighter.FightBase, FightMove.FightMoveItem))
    Call Temp.AddKeyUp(Asc(" ") , ConvertCallable("$0.LetCurrentValue($1)" , MeFighter.FightBase, ItemCallback))
    Call Temp.AddKeyUp(Asc(" ") , ConvertCallable("RemoveRenderObject()"))
    Call Temp.AddKeyUp(27       , ConvertCallable("RemoveRenderObject()"))
    Set CreateInput = Temp
End Function

Private Function CreateInventoryList() As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +0.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +0.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)
    Set CreateInventoryList = FactoryTextBox.Create(Temp, UpdateTextBox(UsedFont))
End Function

Private Function UpdateTextBox(FontLayout As VBGLFontLayout) As VBGLFont()
    Dim Fonts() As VBGLFont
    If MeFighter.FightBase.Items.Count = -1 Then
        ReDim Fonts(0)
        Set Fonts(0) = VBGLFont.Create("No items yet", FontLayout)
        UpdateTextBox = Fonts
        Exit Function
    End If
    ReDim Fonts(MeFighter.FightBase.Items.Count)
    Dim i As Long
    For i = 0 To MeFighter.FightBase.Items.Count
        With MeFighter.FightBase.Items.Item(i)
            Set Fonts(i) = VBGLFont.Create(.Definition.Name & " [" & .Amount & "]" & vbCrLf, FontLayout)
        End With
    Next i
    UpdateTextBox = Fonts
End Function