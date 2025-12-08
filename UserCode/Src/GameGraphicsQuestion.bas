Attribute VB_Name = "GameGraphicsQuestion"


Option Explicit

Private QuestionBox As VBGLTextBox

Public Function SetUpQuestionGraphics() As VBGLRenderObject
    Set QuestionBox           = CreateQuestionBox()
    Set SetUpQuestionGraphics = VBGLRenderObject.Create(New VBGLGeneralInput, CurrentContext.CurrentFrame())
    Call SetUpQuestionGraphics.AddDrawable(QuestionBox)
End Function

Public Sub UpdateQuestion(ByVal Name As String, ByVal Text As String)
    QuestionBox.Font(1).Text = Name & vbCrLf
    QuestionBox.Font(2).Text = Text
    Call QuestionBox.UpdateData()
End Sub

Public Function QuestionBoxInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Call Temp.AddKeyUp(Asc("y"), CreateFixedCallable("EscapeQuestionBox(True)"))
    Call Temp.AddKeyUp(Asc("n"), CreateFixedCallable("EscapeQuestionBox(True)"))
    Set QuestionBoxInput = Temp
End Function

Public Function EscapeQuestionBox(Optional ByVal Setter As Boolean = False) As Boolean
    Static Value As Boolean
    If Setter Then Value = Value Xor True
    EscapeQuestionBox = Value
End Function

Private Function CreateQuestionBox() As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, -0.3!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +1.0!, -0.3!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +1.0!)
    Set CreateQuestionBox = FactoryTextBox.Create(Temp, UpdateTextBox(UsedFont))
End Function

Private Function UpdateTextBox(ByVal FontLayout As VBGLFontLayout) As VBGLFont()
    Dim Fonts() As VBGLFont
    ReDim Fonts(2)
    Set Fonts(0) = VBGLFont.Create("yes[y] no[n]" & vbCrLf, FontLayout)
    Set Fonts(1) = VBGLFont.Create("Name"         & vbCrLf, FontLayout)
    Set Fonts(2) = VBGLFont.Create("Text"                 , FontLayout)
    UpdateTextBox = Fonts
End Function