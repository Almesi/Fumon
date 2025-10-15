Attribute VB_Name = "GameGraphicsFight"


Option Explicit

Private FumonName1   As VBGLTextBox
Private FumonName2   As VBGLTextBox
Private Dialog       As VBGLTextBox
Private History      As VBGLTextBox
Private Buttons      As VBGLTextBox
Private FumonSprites As VBGLMesh
Private FumonHealths As VBGLMesh
Public  FumonTime    As VBGLMesh

Public Sub SetUpFightGraphics()
    FightRenderObject.Inputt = CreateInput()

    Set FumonName1   = GetFumonName1()
    Set FumonName2   = GetFumonName2()
    Set Dialog       = GetDialog()
    Set History      = GetHistory()
    Set Buttons      = GetButtons()
    Set FumonSprites = GetFumonSprites(MePlayer.Fumons.Fumon(0), MePlayer.Fumons.Fumon(0))
    Set FumonHealths = GetFumonHealths(MePlayer.Fumons.Fumon(0), MePlayer.Fumons.Fumon(0))
    Set FumonTime    = GetFumonTimer(1, 1)
    Call FightRenderObject.AddDrawable(Dialog)
    Call FightRenderObject.AddDrawable(History)
    Call FightRenderObject.AddDrawable(Buttons)
    Call FightRenderObject.AddDrawable(FumonSprites)
    Call FightRenderObject.AddDrawable(FumonName2)
    Call FightRenderObject.AddDrawable(FumonName1)
    Call FightRenderObject.AddDrawable(FumonHealths)
    Call FightRenderObject.AddDrawable(FumonTime)
End Sub

Public Sub UpdateFight(ByVal MyFight As Fight, ByVal Player1 As IPlayer, ByVal Player2 As IPlayer, ByVal Fumon1 As Fumon, ByVal Fumon2 As Fumon, Optional DialogText As String = Empty, Optional CurrentMove As String = Empty, Optional SelectedButton As Long = 0)
    FumonName1.Font(0).Text = Fumon1.Definition.Name
    FumonName2.Font(0).Text = Fumon2.Definition.Name
    Dialog.Font(0).Text     = DialogText
    Dim Text As String
    Text = HistoryText(MyFight, Player1) & vbCrLf & HistoryText(MyFight, Player2)
    History.Font(0).Text = Text

    Call History.UpdateData()
    Call FumonName1.UpdateData()
    Call FumonName2.UpdateData()

    Dim i As Long
    Dim Color() As Single
    ReDim Color(2)
    For i = 0 To 3
        Buttons.Font(i).FontColor = Color
    Next i
    Color(1) = 1
    Buttons.Font(SelectedButton).FontColor = Color
    
    Dim TempData() As Single
    TempData = UpdateSprites(Fumon1, Fumon2)
    Call FumonSprites.VAO.Buffer.Update(VBGLData.CreateSingle(TempData))
    TempData = UpdateHealthBars(Fumon1, Fumon2)
    Call FumonHealths.VAO.Buffer.Update(VBGLData.CreateSingle(TempData))
    TempData = UpdateFumonTimer(1, 1)
    Call FumonTime.VAO.Buffer.Update(VBGLData.CreateSingle(TempData))
End Sub

Private Function UpdateSprites(ByVal Fumon1 As Fumon, ByVal Fumon2 As Fumon) As Single()
    Dim Vertices() As Single
    Dim VertexSize  As Long: VertexSize  = 5
    Dim VertexCount As Long: VertexCount = 6
    Dim FumonsCount As Long: FumonsCount = 2
    ReDim Vertices(VertexSize * VertexCount  * FumonsCount - 1)
    ' xyz txty

    With MeServer.GetTexture("Fumons").SubTextureID(Fumon1.Definition.Name & "Back")
        Vertices(00) = -1: Vertices(01) = +0: Vertices(02) = 0: Vertices(03) = .GetX("TopLeft")     : Vertices(04) = .GetY("TopLeft")
        Vertices(05) = +0: Vertices(06) = +0: Vertices(07) = 0: Vertices(08) = .GetX("TopRight")    : Vertices(09) = .GetY("TopRight")
        Vertices(10) = -1: Vertices(11) = -1: Vertices(12) = 0: Vertices(13) = .GetX("BottomLeft")  : Vertices(14) = .GetY("BottomLeft")
        Vertices(15) = +0: Vertices(16) = +0: Vertices(17) = 0: Vertices(18) = .GetX("TopRight")    : Vertices(19) = .GetY("TopRight")
        Vertices(20) = +0: Vertices(21) = -1: Vertices(22) = 0: Vertices(23) = .GetX("BottomRight") : Vertices(24) = .GetY("BottomRight")
        Vertices(25) = -1: Vertices(26) = -1: Vertices(27) = 0: Vertices(28) = .GetX("BottomLeft")  : Vertices(29) = .GetY("BottomLeft")
    End With

    With MeServer.GetTexture("Fumons").SubTextureID(Fumon2.Definition.Name & "Front")
        Vertices(30) = +0: Vertices(31) = +1: Vertices(32) = 0: Vertices(33) = .GetX("TopLeft")     : Vertices(34) = .GetY("TopLeft")
        Vertices(35) = +1: Vertices(36) = +1: Vertices(37) = 0: Vertices(38) = .GetX("TopRight")    : Vertices(39) = .GetY("TopRight")
        Vertices(40) = +0: Vertices(41) = +0: Vertices(42) = 0: Vertices(43) = .GetX("BottomLeft")  : Vertices(44) = .GetY("BottomLeft")
        Vertices(45) = +1: Vertices(46) = +1: Vertices(47) = 0: Vertices(48) = .GetX("TopRight")    : Vertices(49) = .GetY("TopRight")
        Vertices(50) = +1: Vertices(51) = +0: Vertices(52) = 0: Vertices(53) = .GetX("BottomRight") : Vertices(54) = .GetY("BottomRight")
        Vertices(55) = +0: Vertices(56) = +0: Vertices(57) = 0: Vertices(58) = .GetX("BottomLeft")  : Vertices(59) = .GetY("BottomLeft")
    End With
    UpdateSprites = Vertices
End Function

Private Function UpdateHealthBars(ByVal Fumon1 As Fumon, ByVal Fumon2 As Fumon) As Single()
    Dim MaxHealth1    As Long  : MaxHealth1    = Fumon1.MaxHealth
    Dim MaxHealth2    As Long  : MaxHealth2    = Fumon2.MaxHealth
    Dim Fumon1Percent As Single: Fumon1Percent = Fumon1.CurrentHealth.Value / MaxHealth1
    Dim Fumon2Percent As Single: Fumon2Percent = Fumon2.CurrentHealth.Value / MaxHealth2

    Dim Vertices() As Single
    'xy rgb
    Dim VertexSize  As Long: VertexSize  = 5
    Dim VertexCount As Long: VertexCount = 6
    Dim FumonsCount As Long: FumonsCount = 2
    ReDim Vertices(VertexSize * VertexCount  * FumonsCount - 1)
    Vertices(00) = -1.0                 : Vertices(01) = -0.9: Vertices(02) = 1 - Fumon1Percent: Vertices(03) = Fumon1Percent : Vertices(04) = 0
    Vertices(05) = -1.0 + Fumon1Percent : Vertices(06) = -0.9: Vertices(07) = 1 - Fumon1Percent: Vertices(08) = Fumon1Percent : Vertices(09) = 0
    Vertices(10) = -1.0                 : Vertices(11) = -1.0: Vertices(12) = 1 - Fumon1Percent: Vertices(13) = Fumon1Percent : Vertices(14) = 0
    Vertices(15) = -1.0 + Fumon1Percent : Vertices(16) = -0.9: Vertices(17) = 1 - Fumon1Percent: Vertices(18) = Fumon1Percent : Vertices(19) = 0
    Vertices(20) = -1.0 + Fumon1Percent : Vertices(21) = -1.0: Vertices(22) = 1 - Fumon1Percent: Vertices(23) = Fumon1Percent : Vertices(24) = 0
    Vertices(25) = -1.0                 : Vertices(26) = -1.0: Vertices(27) = 1 - Fumon1Percent: Vertices(28) = Fumon1Percent : Vertices(29) = 0

    Vertices(30) = +0.0                 : Vertices(31) = +0.1: Vertices(32) = 1 - Fumon2Percent: Vertices(33) = Fumon2Percent : Vertices(34) = 0
    Vertices(35) = +0.0 + Fumon2Percent : Vertices(36) = +0.1: Vertices(37) = 1 - Fumon2Percent: Vertices(38) = Fumon2Percent : Vertices(39) = 0
    Vertices(40) = +0.0                 : Vertices(41) = +0.0: Vertices(42) = 1 - Fumon2Percent: Vertices(43) = Fumon2Percent : Vertices(44) = 0
    Vertices(45) = +0.0 + Fumon2Percent : Vertices(46) = +0.1: Vertices(47) = 1 - Fumon2Percent: Vertices(48) = Fumon2Percent : Vertices(49) = 0
    Vertices(50) = +0.0 + Fumon2Percent : Vertices(51) = +0.0: Vertices(52) = 1 - Fumon2Percent: Vertices(53) = Fumon2Percent : Vertices(54) = 0
    Vertices(55) = +0.0                 : Vertices(56) = +0.0: Vertices(57) = 1 - Fumon2Percent: Vertices(58) = Fumon2Percent : Vertices(59) = 0

    UpdateHealthBars = Vertices
End Function

Public Function UpdateFumonTimer(ByVal Time1 As Single, ByVal Time2 As Single) As Single()
    Dim Vertices() As Single
    'xy rgb
    Dim VertexSize  As Long: VertexSize  = 5
    Dim VertexCount As Long: VertexCount = 6
    Dim FumonsCount As Long: FumonsCount = 2
    ReDim Vertices(VertexSize * VertexCount  * FumonsCount - 1)
    Dim Time1Offset As Single : Time1Offset = Time1 * 0.9
    Dim Time2Offset As Single : Time2Offset = Time2 * 0.9
    Vertices(00) = -1.0 : Vertices(01) = -0.0 - Time1Offset: Vertices(02) = Time1: Vertices(03) = 1 - Time1 : Vertices(04) = 0
    Vertices(05) = -0.9 : Vertices(06) = -0.0 - Time1Offset: Vertices(07) = Time1: Vertices(08) = 1 - Time1 : Vertices(09) = 0
    Vertices(10) = -1.0 : Vertices(11) = -0.9              : Vertices(12) = Time1: Vertices(13) = 1 - Time1 : Vertices(14) = 0
    Vertices(15) = -0.9 : Vertices(16) = -0.0 - Time1Offset: Vertices(17) = Time1: Vertices(18) = 1 - Time1 : Vertices(19) = 0
    Vertices(20) = -0.9 : Vertices(21) = -0.9              : Vertices(22) = Time1: Vertices(23) = 1 - Time1 : Vertices(24) = 0
    Vertices(25) = -1.0 : Vertices(26) = -0.9              : Vertices(27) = Time1: Vertices(28) = 1 - Time1 : Vertices(29) = 0

    Vertices(30) = +0.0 : Vertices(31) = +1.0 - Time2Offset: Vertices(32) = Time2: Vertices(33) = 1 - Time2 : Vertices(34) = 0
    Vertices(35) = +0.1 : Vertices(36) = +1.0 - Time2Offset: Vertices(37) = Time2: Vertices(38) = 1 - Time2 : Vertices(39) = 0
    Vertices(40) = +0.0 : Vertices(41) = +0.1              : Vertices(42) = Time2: Vertices(43) = 1 - Time2 : Vertices(44) = 0
    Vertices(45) = +0.1 : Vertices(46) = +1.0 - Time2Offset: Vertices(47) = Time2: Vertices(48) = 1 - Time2 : Vertices(49) = 0
    Vertices(50) = +0.1 : Vertices(51) = +0.1              : Vertices(52) = Time2: Vertices(53) = 1 - Time2 : Vertices(54) = 0
    Vertices(55) = +0.0 : Vertices(56) = +0.1              : Vertices(57) = Time2: Vertices(58) = 1 - Time2 : Vertices(59) = 0

    UpdateFumonTimer = Vertices
End Function

Private Function CreateInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput

    Call Temp.AddKeyUp(Asc("f") , VBGLCallable.Create(Nothing              , "AddRenderObject"    , vbMethod, +0 , FumonRenderObject))
    Call Temp.AddKeyUp(Asc("i") , VBGLCallable.Create(Nothing              , "AddRenderObject"    , vbMethod, +0 , InventoryRenderObject))
    Call Temp.AddKeyUp(Asc("r") , VBGLCallable.Create(MePlayer             , "CurrentMove"        , vbLet   , +0 , FightMove.FightMoveFlee))
    Call Temp.AddKeyUp(Asc("a") , VBGLCallable.Create(Nothing              , "AddRenderObject"    , vbMethod, +0 , AttackRenderObject))
    Set CreateInput = Temp
End Function

Private Function GetFumonName1() As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, +0.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +0.0!, +0.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, -0.2!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +0.0!, -0.2!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)
    Set GetFumonName1 = VBGLTextBox.CreateFromText(Temp, "FUMON2", UsedFont)
End Function
Private Function GetFumonName2() As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , +0.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , +0.0!, +0.8!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +1.0!, +0.8!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)
    Set GetFumonName2 = VBGLTextBox.CreateFromText(Temp, "FUMON1", UsedFont)
End Function
Private Function GetDialog() As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, +0.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +0.0!, +0.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +0.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)
    Set GetDialog = VBGLTextBox.CreateFromText(Temp, "DIALOG", UsedFont)
End Function
Private Function GetHistory() As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , +0.0!, +0.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +1.0!, +0.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , +0.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)
    Set GetHistory = VBGLTextBox.CreateFromText(Temp, " ", UsedFont)
End Function
Private Function GetButtons() As VBGLTextBox

    Dim Fonts() As VBGLFont
    ReDim Fonts(3)
    Dim i As Long
    For i = 0 To 3
        Set Fonts(i) = New VBGLFont
        Fonts(i).FontLayout = UsedFont
    Next i
    Fonts(0).Text = "Attacks"   & vbCrLf
    Fonts(1).Text = "Fumons"    & vbCrLf
    Fonts(2).Text = "Inventory" & vbCrLf
    Fonts(3).Text = "Flee"
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , -0.8!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , -0.8!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)
    Set GetButtons = VBGLTextBox.Create(Temp, Fonts)
End Function
Private Function GetFumonSprites(ByVal Fumon1 As Fumon, ByVal Fumon2 As Fumon) As VBGLMesh
    Dim Data As IDataSingle
    Set Data = VBGLData.CreateSingle(UpdateSprites(Fumon1, Fumon2))

    Set GetFumonSprites = VBGLMesh.Create(VBGLPrCoShaderXYZTxTy, VBGLPrCoLayoutXYZTxTy, Data)
    Call GetFumonSprites.AddTexture(MeServer.GetTexture("Fumons"))
End Function
Private Function GetFumonHealths(ByVal Fumon1 As Fumon, ByVal Fumon2 As Fumon) As VBGLMesh
    Dim Data As IDataSingle
    Set Data = VBGLData.CreateSingle(UpdateHealthBars(Fumon1, Fumon2))

    Set GetFumonHealths = VBGLMesh.Create(VBGLPrCoShaderXYRGB, VBGLPrCoLayoutXYRGB, Data)
End Function
Private Function GetFumonTimer(ByVal Time1 As Single, ByVal Time2 As Single) As VBGLMesh
    Dim Data As IDataSingle
    Set Data = VBGLData.CreateSingle(UpdateFumonTimer(Time1, Time2))

    Set GetFumonTimer = VBGLMesh.Create(VBGLPrCoShaderXYRGB, VBGLPrCoLayoutXYRGB, Data)
End Function
Private Function HistoryText(ByVal MyFight As Fight, ByVal Player As IPlayer) As String
    Dim Obj As Object
    Set Obj = Player.GetCurrentValue(MyFight)
    Dim Value As String
    Select Case TypeName(Obj)
        Case "Attack" : Value = Obj.Name
        Case "Fumon"  : Value = Obj.Definition.Name
        Case "Item"   : Value = Obj.Name
    End Select
    Dim PlayerName As String
    PlayerName = Player.Name
    Select Case Player.GetCurrentMove
        Case FightMove.FightMoveAttack      : HistoryText = PlayerName & " used Attack " & Value
        Case FightMove.FightMoveFlee        : HistoryText = PlayerName & " tried to flee"
        Case FightMove.FightMoveChangeFumon : HistoryText = PlayerName & " changed to Fumon " & Value
        Case FightMove.FightMoveNothing     : HistoryText = PlayerName & " skipped a turn"
        Case FightMove.FightMoveItem        : HistoryText = PlayerName & " used Item " & Value
    End Select
End Function