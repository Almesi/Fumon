Attribute VB_Name = "GameGraphicsAttack"


Option Explicit

Private AttackList As VBGLTextBox

Public Sub SetUpAttackGraphics()
    Set AttackList = CreateAttackList
    AttackRenderObject.Inputt = CreateInput()
    Call AttackRenderObject.AddDrawable(AttackList)
End Sub

Public Sub UpdateAttack(ByVal Index As Long)
    Dim i As Long
    Dim Color() As Single
    ReDim Color(3)
    MePlayer.Fumons.FirstAlive.Attacks.Selected = Index
    AttackList.Fonts = UpdateTextBox(UsedFont)
    For i = 0 To MePlayer.Fumons.FirstAlive.Attacks.AttackCount
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
    Set GetSelected = VBGLCallable.Create(MePlayer.Fumons.FirstAlive.Attacks, "Selected", vbGet, -1)
    Call Temp.AddKeyUp(Asc("1") , VBGLCallable.Create(Nothing  , "UpdateAttack"       , vbMethod, 0, 0))
    Call Temp.AddKeyUp(Asc("2") , VBGLCallable.Create(Nothing  , "UpdateAttack"       , vbMethod, 0, 1))
    Call Temp.AddKeyUp(Asc("3") , VBGLCallable.Create(Nothing  , "UpdateAttack"       , vbMethod, 0, 2))
    Call Temp.AddKeyUp(Asc("4") , VBGLCallable.Create(Nothing  , "UpdateAttack"       , vbMethod, 0, 3))
    Call Temp.AddKeyUp(Asc(" ") , VBGLCallable.Create(MePlayer , "LetCurrentMove"     , vbMethod, 0, FightMove.FightMoveAttack))
    Call Temp.AddKeyUp(Asc(" ") , VBGLCallable.Create(MePlayer , "LetCurrentValue"    , vbMethod, 0, GetSelected))
    Call Temp.AddKeyUp(Asc(" ") , VBGLCallable.Create(Nothing  , "RemoveRenderObject" , vbMethod, -1))
    Call Temp.AddKeyUp(27       , VBGLCallable.Create(Nothing  , "RemoveRenderObject" , vbMethod, -1))
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
    Set CreateAttackList = VBGLTextBox.Create(Temp, UpdateTextBox(UsedFont))
End Function

Private Function UpdateTextBox(FontLayout As VBGLFontLayout) As VBGLFont()
    Dim Fonts() As VBGLFont
    Dim Text As String
    With MePlayer.Fumons.FirstAlive.Attacks
        ReDim Fonts(.AttackCount)
        Dim i As Long
        For i = 0 To .AttackCount
            Text = .Attack(i).Name & " " & .Attack(i).GetTypeName() & ": ( " & .Attack(i).ElementType & " | " & .Attack(i).Func & " )" & vbCrLf
            Set Fonts(i) = VBGLFont.Create(Text , FontLayout)
        Next i
    End With
    UpdateTextBox = Fonts
End Function