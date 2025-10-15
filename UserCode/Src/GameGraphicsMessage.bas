Attribute VB_Name = "GameGraphicsMessage"


Option Explicit

Public MessageBox As VBGLTextBox

Public Sub SetUpMessageGraphics()
    Set MessageBox = CreateMessageBox()
End Sub

Public Sub UpdateMessage(ByVal Name As String, ByVal Text As String)
    MessageBox.Font(0).Text = Name & vbCrLf
    MessageBox.Font(1).Text = Text
    Call MessageBox.UpdateData()
End Sub

Public Function MessageBoxInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Call Temp.AddKeyUp(27, VBGLCallable.Create(Nothing , "EscapeTextBox"      , vbMethod, 0, True))
    Set MessageBoxInput = Temp
End Function

Private Function CreateMessageBox() As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, -0.3!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +1.0!, -0.3!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +1.0!)
    Set CreateMessageBox = VBGLTextBox.Create(Temp, UpdateTextBox(UsedFont))
End Function

Private Function UpdateTextBox(ByVal FontLayout As VBGLFontLayout) As VBGLFont()
    Dim Fonts() As VBGLFont
    ReDim Fonts(1)
    Set Fonts(0) = VBGLFont.Create("Name" & vbCrLf, FontLayout)
    Set Fonts(1) = VBGLFont.Create("Text"         , FontLayout)
    UpdateTextBox = Fonts
End Function