Attribute VB_Name = "GameGraphicsOptions"


Option Explicit

Private OptionsList As VBGLTextBox

Public Function SetUpOptionsGraphics() As VBGLRenderObject
    Set OptionsList = CreateOptionsList()
    Set SetUpOptionsGraphics = VBGLRenderObject.Create(CreateInput(), CurrentContext.CurrentFrame())
    Call SetUpOptionsGraphics.AddDrawable(OptionsList)
End Function

Public Sub UpdateOptions(ByVal Index As Long)
    Call RemoveRenderObject()
End Sub

Private Function CreateInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Call Temp.AddKeyUp(27, ConvertCallable("RemoveRenderObject()"))
    Set CreateInput = Temp
End Function

Private Function CreateOptionsList() As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)
    Set CreateOptionsList = FactoryTextBox.Create(Temp, UpdateTextBox(UsedFont))
End Function

Private Function UpdateTextBox(FontLayout As VBGLFontLayout) As VBGLFont()
    Dim Fonts() As VBGLFont
    ReDim Fonts(37)
    Set Fonts(00) = VBGLFont.Create("Overworld: "                  & vbCrLf, FontLayout)
    Set Fonts(01) = VBGLFont.Create("w          = Move Up"         & vbCrLf, FontLayout)
    Set Fonts(02) = VBGLFont.Create("a          = Move Left"       & vbCrLf, FontLayout)
    Set Fonts(03) = VBGLFont.Create("s          = Move Down"       & vbCrLf, FontLayout)
    Set Fonts(04) = VBGLFont.Create("d          = Move Right"      & vbCrLf, FontLayout)
    Set Fonts(05) = VBGLFont.Create("W          = Look Up"         & vbCrLf, FontLayout)
    Set Fonts(06) = VBGLFont.Create("A          = Look Left"       & vbCrLf, FontLayout)
    Set Fonts(07) = VBGLFont.Create("S          = Look Down"       & vbCrLf, FontLayout)
    Set Fonts(08) = VBGLFont.Create("D          = Look Right"      & vbCrLf, FontLayout)
    Set Fonts(09) = VBGLFont.Create("WHITESPACE = Interact"        & vbCrLf, FontLayout)
    Set Fonts(10) = VBGLFont.Create("m          = Open Map"        & vbCrLf, FontLayout)
    Set Fonts(11) = VBGLFont.Create("i          = Open Inventory"  & vbCrLf, FontLayout)
    Set Fonts(12) = VBGLFont.Create("f          = Open Fumons"     & vbCrLf, FontLayout)
    Set Fonts(13) = VBGLFont.Create("ESC        = Exit Game"       & vbCrLf, FontLayout)

    Set Fonts(14) = VBGLFont.Create("Map: "                        & vbCrLf, FontLayout)
    Set Fonts(15) = VBGLFont.Create("ESC        = Exit Map"        & vbCrLf, FontLayout)

    Set Fonts(16) = VBGLFont.Create("Inventory: "                  & vbCrLf, FontLayout)
    Set Fonts(17) = VBGLFont.Create("w          = Go Up"           & vbCrLf, FontLayout)
    Set Fonts(18) = VBGLFont.Create("s          = Go Down"         & vbCrLf, FontLayout)
    Set Fonts(19) = VBGLFont.Create("WHITESPACE = Use"             & vbCrLf, FontLayout)
    Set Fonts(20) = VBGLFont.Create("ESC        = Exit Inventory"  & vbCrLf, FontLayout)


    Set Fonts(21) = VBGLFont.Create("Fumons: "                                    & vbCrLf, FontLayout)
    Set Fonts(22) = VBGLFont.Create("1          = Select First Fumon"             & vbCrLf, FontLayout)
    Set Fonts(23) = VBGLFont.Create("2          = Select Second Fumon"            & vbCrLf, FontLayout)
    Set Fonts(24) = VBGLFont.Create("3          = Select Third Fumon"             & vbCrLf, FontLayout)
    Set Fonts(25) = VBGLFont.Create("4          = Select Fourth Fumon"            & vbCrLf, FontLayout)
    Set Fonts(26) = VBGLFont.Create("5          = Select Fifth Fumon"             & vbCrLf, FontLayout)
    Set Fonts(27) = VBGLFont.Create("6          = Select Sixth Fumon"             & vbCrLf, FontLayout)
    Set Fonts(28) = VBGLFont.Create("7          = Select Seventh Fumon"           & vbCrLf, FontLayout)
    Set Fonts(29) = VBGLFont.Create("8          = Select Eigth Fumon"             & vbCrLf, FontLayout)
    Set Fonts(30) = VBGLFont.Create("WHITESPACE = Swap Selected with First Fumon" & vbCrLf, FontLayout)
    Set Fonts(31) = VBGLFont.Create("ESC        = Exit Fumons"                    & vbCrLf, FontLayout)

    Set Fonts(32) = VBGLFont.Create("Attacks: "          & vbCrLf, FontLayout)
    Set Fonts(33) = VBGLFont.Create("1   = Use Attack 1" & vbCrLf, FontLayout)
    Set Fonts(34) = VBGLFont.Create("2   = Use Attack 2" & vbCrLf, FontLayout)
    Set Fonts(35) = VBGLFont.Create("3   = Use Attack 3" & vbCrLf, FontLayout)
    Set Fonts(36) = VBGLFont.Create("4   = Use Attack 4" & vbCrLf, FontLayout)
    Set Fonts(37) = VBGLFont.Create("ESC = Exit Attacks" & vbCrLf, FontLayout)

    UpdateTextBox = Fonts
End Function