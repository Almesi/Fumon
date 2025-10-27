Attribute VB_Name = "GameGraphicsAttack"


Option Explicit

Private AttackList As VBGLTextBox

Public Function SetUpAttackGraphics() As VBGLRenderObject
    Set AttackList = CreateAttackList()
    Set SetUpAttackGraphics = VBGLRenderObject.Create(CreateInput(), CurrentContext.CurrentFrame())
    Call SetUpAttackGraphics.AddDrawable(AttackList)
End Function

Public Sub UpdateAttack(ByVal Index As Long)
    Dim i As Long
    Dim Color() As Single
    ReDim Color(3)
    MeFighter.FightBase.Fumons.FirstAlive.Attacks.Selected = Index
    AttackList.Fonts = UpdateTextBox(UsedFont)
    For i = 0 To MeFighter.FightBase.Fumons.FirstAlive.Attacks.AttackCount
        AttackList.Font(i).FontColor = Color
    Next i
    Color(1) = 1
    AttackList.Font(Index).FontColor = Color
    Call AttackList.UpdateData()
End Sub

Private Function CreateInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Dim GetSelected As VBGLCallable
    Set GetSelected = ConvertCallable("$0.Selected()", MeFighter.FightBase.Fumons.FirstAlive.Attacks)
    Call Temp.AddKeyUp(Asc("1") , ConvertCallable("UpdateAttack(0)"))
    Call Temp.AddKeyUp(Asc("2") , ConvertCallable("UpdateAttack(1)"))
    Call Temp.AddKeyUp(Asc("3") , ConvertCallable("UpdateAttack(2)"))
    Call Temp.AddKeyUp(Asc("4") , ConvertCallable("UpdateAttack(3)"))
    Call Temp.AddKeyUp(Asc(" ") , ConvertCallable("$0.LetCurrentMove()" , MeFighter.FightBase))
    Call Temp.AddKeyUp(Asc(" ") , ConvertCallable("$0.LetCurrentValue()", MeFighter.FightBase))
    Call Temp.AddKeyUp(Asc(" ") , ConvertCallable("RemoveRenderObject()"))
    Call Temp.AddKeyUp(27       , ConvertCallable("RemoveRenderObject()"))
    Set CreateInput = Temp
End Function

Private Function CreateAttackList() As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +0.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +0.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)
    Set CreateAttackList = FactoryTextBox.Create(Temp, UpdateTextBox(UsedFont))
End Function

Private Function UpdateTextBox(FontLayout As VBGLFontLayout) As VBGLFont()
    Dim Fonts() As VBGLFont
    Dim Text As String
    With MeFighter.FightBase.Fumons.FirstAlive.Attacks
        ReDim Fonts(.AttackCount)
        Dim i As Long
        For i = 0 To .AttackCount
            Text = .Attack(i).Name & " " & .Attack(i).GetTypeName() & ": ( " & .Attack(i).ElementType & " | " & .Attack(i).Func & " )" & vbCrLf
            Set Fonts(i) = VBGLFont.Create(Text , FontLayout)
        Next i
    End With
    UpdateTextBox = Fonts
End Function