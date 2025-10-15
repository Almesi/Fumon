Attribute VB_Name = "GameGraphicsInventory"


Option Explicit

Private InventoryList As VBGLTextBox

Public Sub SetUpInventoryGraphics()
    Set InventoryList = CreateInventoryList
    InventoryRenderObject.Inputt = CreateInput()
    Call InventoryRenderObject.AddDrawable(InventoryList)
End Sub

Public Sub UpdateInventory(ByVal Offset As Long)
    Debug.Assert False
    Dim i As Long
    Dim Color() As Single
    Dim Index As Long
    Call MePlayer.Items.Increment(Offset)
    Index = MePlayer.Items.Selected
    ReDim Color(2)
    InventoryList.Fonts = UpdateTextBox(UsedFont)
    For i = 0 To MePlayer.Items.Count
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
    Set ItemCallback = VBGLCallable.Create(MePlayer.Items, "SelectedItem", vbGet, -1)

    Call Temp.AddKeyUp(Asc("w") , VBGLCallable.Create(Nothing              , "UpdateInventory"    , vbMethod, 0, +1))
    Call Temp.AddKeyUp(Asc("s") , VBGLCallable.Create(Nothing              , "UpdateInventory"    , vbMethod, 0, -1))
    Call Temp.AddKeyUp(Asc(" ") , VBGLCallable.Create(MePlayer             , "CurrentMove"        , vbLet   , 0, FightMove.FightMoveItem))
    Call Temp.AddKeyUp(Asc(" ") , VBGLCallable.Create(MePlayer             , "CurrentValue"       , vbLet   , 0, ItemCallback))
    Call Temp.AddKeyUp(Asc(" ") , VBGLCallable.Create(Nothing              , "RemoveRenderObject" , vbMethod, -1))
    Call Temp.AddKeyUp(27       , VBGLCallable.Create(Nothing              , "RemoveRenderObject" , vbMethod, -1))
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
    Set CreateInventoryList = VBGLTextBox.Create(Temp, UpdateTextBox(UsedFont))
End Function

Private Function UpdateTextBox(FontLayout As VBGLFontLayout) As VBGLFont()
    Dim Fonts() As VBGLFont
    If MePlayer.Items.Count = -1 Then
        ReDim Fonts(0)
        Set Fonts(0) = VBGLFont.Create("No items yet", FontLayout)
        UpdateTextBox = Fonts
        Exit Function
    End If
    ReDim Fonts(MePlayer.Items.Count)
    Dim i As Long
    For i = 0 To MePlayer.Items.Count
        With MePlayer.Items.Item(i)
            Set Fonts(i) = VBGLFont.Create(.Definition.Name & " [" & .Amount & "]" & vbCrLf, FontLayout)
        End With
    Next i
    UpdateTextBox = Fonts
End Function