Attribute VB_Name = "GameGraphicsStart"


Option Explicit


Public Function SetUpStartGraphics() As VBGLRenderObject
    Set SetUpStartGraphics = VBGLRenderObject.Create(CreateInput(), CurrentContext.CurrentFrame())
    Call SetUpStartGraphics.AddDrawable(CreateTextBox())
End Function

Private Function CreateTextBox() As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)
    Set CreateTextBox = FactoryTextBox.CreateFromText(Temp, _
                                                   "Welcome to Fumon"            & vbCrLf & _
                                                   "To start the Game press [s]" & vbCrLf & _
                                                   "To View the Options [o]"     & vbCrLf & _
                                                   "To cancel Press [ESC]"       & vbCrLf, UsedFont)
End Function

Private Function CreateInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput
    Call Temp.AddKeyUp(Asc("s"), ConvertCallable("AddRenderObject($0)", OverWorldRenderObject))
    Call Temp.AddKeyUp(Asc("o"), ConvertCallable("AddRenderObject($0)", OptionsRenderObject))
    Call Temp.AddKeyUp(27      , ConvertCallable("$0.LeaveMainLoop()", CurrentContext))
    Set CreateInput = Temp
End Function