Attribute VB_Name = "InitializeGameProcedures"


Option Explicit
Public MeServer       As GameServer
Public MePlayer       As IPlayer
Public MeOwner        As ServerOwner
Public AllCallables   As std_AllCallables


Public Sub InitializeGame()
    Dim StartTime As Single
    StartTime = Timer
    Dim CurrentBuild As std_VBProject
    Dim PlayerSettings As IConfig

    Dim Shower As IDestination: Set Shower = Nothing
    Dim Logger As IDestination: Set Logger = std_ImmiedeateDestination.Create()
    
    Dim NewErrorHandler As std_ErrorHandler
    Set NewErrorHandler = std_ErrorHandler.Create(Shower, Logger)

    Debug.Print "Start: " & Timer - StartTime

    Set PlayerSettings = std_ConfigRange.Create(ThisWorkbook.Sheets("Settings"))
    Set CurrentBuild = std_VBProject.Create(ThisWorkbook.VBProject, NewErrorHandler)
    Set AllCallables = std_AllCallables.CreateFromProject(ThisWorkbook.VBProject)
    Debug.Print "AllCallables: " & Timer - StartTime

    Call CreateContextAndWindow(Logger, Shower)
    If IsNothing(CurrentContext) Then Exit Sub
    Debug.Print "WindowCreation: " & Timer - StartTime

    Set MeServer = ConnectToServer(PlayerSettings)
    If IsNothing(MeServer) Then Exit Sub
    Debug.Print "ServerConnection: " & Timer - StartTime
    
    Set MePlayer = MeServer.GetPlayer(PlayerSettings.Setting("Username"))
    If IsNothing(MePlayer) Then Exit Sub
    MePlayer.MoveBase.DoInteract = True
    Debug.Print "PlayerCreation: " & Timer - StartTime

    Set FactoryTextBox = New VBGLTextBox
    With FactoryTextBox
        .CharsPerLine   = 128
        .LinesPerPage   = 128
        .Pages          = 1
        .LineOffset     = 0.1!
    End With
    Debug.Print "FactoryCreation: " & Timer - StartTime

    Set FactoryTextBoxProperties = FactoryTextBox.CreateProperties(2, 3)
    Set MeOwner = ServerOwner.Create(MeServer)
    Debug.Print "ServerOwnerCreation: " & Timer - StartTime
    Call SetUpGameGraphics()
    Debug.Print "GraphicsSetUp: " & Timer - StartTime
    Call MeServer.Workbook.Close
    Debug.Print "ClosingBook: " & Timer - StartTime
End Sub

Private Sub CreateContextAndWindow(ByVal Logger As IDestination, ByVal Shower As IDestination)
    If IsSomething(CurrentContext) Then Exit Sub
    Set CurrentContext = VBGLContext.Create("C:\Users\deallulic\Documents\GitHub\VBGL\Code\Src\Graphics\Core", GLUT_CORE_PROFILE, GLUT_DEBUG, Logger, Shower)
    If IsNothing(CurrentContext) Then Exit Sub
    Call VBGLWindow.Create(1600, 900, GLUT_RGBA, "Fumon", "4_6", True)
    CurrentContext.BlendTest = True 
    CurrentContext.DepthTest = True
    CurrentContext.CullTest = True
    Call CurrentContext.DepthFunc(GL_LEQUAL)
    Call CurrentContext.BlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)
    Call CurrentContext.CullFace(GL_BACK)
End Sub

Private Function ConnectToServer(ByVal Settings As IConfig) As GameServer
    
    Call Workbooks.Open(Settings.Setting("Server"))
    Call InitializeAllRanges(Workbooks(Workbooks.Count))
    Set ConnectToServer = GameServer.Create(Workbooks(Workbooks.Count))
    Exit Function

    Error:
    Call CurrentContext.ErrorHandler.Raise(std_Error.Create(Err.Source, "Severe", "Couldnt connect to Server", Err.Description, Empty))
    If IsSomething(CurrentContext) Then
        Call glutDestroyWindow(CurrentContext.CurrentWindow.ID)
        CurrentContext.CurrentWindow = Nothing
    End If
End Function

Public Function SetUpGameGraphics() As Boolean
    Dim FreetypePath As String: FreetypePath = "C:\Users\deallulic\Documents\GitHub\VBGL\Code\Src\Graphics\Extensions\Drawables\TextRendering"
    Dim FontPath     As String: FontPath     = "C:\Users\deallulic\Documents\GitHub\VBGL\Code\Res\Fonts\Consolas.ttf"

    Set UsedFont = VBGLFontLayout.Create(FreetypePath, FontPath, 48)
    Set RenderStack = New std_Stack
    Set InputStack  = New std_Stack
    Call AddRenderObject(StartRenderObject)

    Application.EnableCancelKey = xlDisabled
    Call SetForegroundWindow(CurrentContext.CurrentWindow.ID)

    ' In Reversed order, so that addrenderobject will not recieve nothing, as the next renderobject is not created yet
    Set FarLeftSideFrame = VBGLFrame.CreateFromWindow(0   , 0, 0.25, 1, 0   , 0, 0.25, 1, CurrentContext.CurrentWindow)
    Set LeftSideFrame    = VBGLFrame.CreateFromWindow(0.25, 0, 0.25, 1, 0.25, 0, 0.25, 1, CurrentContext.CurrentWindow)
    Dim StartTime As Single
    StartTime = Timer
    Debug.Print "Frames: " & Timer - StartTime


    Set GameTextures = InitTextures(MeServer.WorkBook)
    Debug.Print "Textures: " & Timer - StartTime

    Set QuestionRenderObject  = SetUpMessageGraphics()   : Call QuestionRenderObject.AssignColor(1, 1, 1, 1)
    Debug.Print "Questions: " & Timer - StartTime
    Set MessageRenderObject   = SetUpMessageGraphics()   : Call MessageRenderObject.AssignColor(1, 1, 1, 1)
    Debug.Print "Messages: " & Timer - StartTime
    Set OptionsRenderObject   = SetUpOptionsGraphics()   : Call OptionsRenderObject.AssignColor(1, 1, 1, 1)
    Debug.Print "Options: " & Timer - StartTime
    Set AttackRenderObject    = SetUpAttackGraphics()    : Call AttackRenderObject.AssignColor(1, 1, 1, 1)
    Debug.Print "Attacks: " & Timer - StartTime
    Set FumonRenderObject     = SetUpFumonGraphics()     : Call FumonRenderObject.AssignColor(1, 1, 1, 1)
    Debug.Print "Fumons: " & Timer - StartTime
    Set InventoryRenderObject = SetUpInventoryGraphics() : Call InventoryRenderObject.AssignColor(1, 1, 1, 1)
    Debug.Print "Inventory: " & Timer - StartTime
    Set FightRenderObject     = SetUpFightGraphics()     : Call FightRenderObject.AssignColor(1, 1, 1, 1)
    Debug.Print "Fight: " & Timer - StartTime
    Call SetUpTileSet()
    Debug.Print "TileSet: " & Timer - StartTime
    Set MapRenderObject       = SetUpMapGraphics()       : Call MapRenderObject.AssignColor(0, 0, 0, 1)
    Debug.Print "Map: " & Timer - StartTime
    Set OverWorldRenderObject = SetUpOverWorldGraphics() : Call OverWorldRenderObject.AssignColor(0, 0, 0, 1)
    Debug.Print "OverWorld: " & Timer - StartTime
    Set StartRenderObject     = SetUpStartGraphics()     : Call StartRenderObject.AssignColor(1, 1, 1, 1)
    Debug.Print "StartRenderObject: " & Timer - StartTime

    Call VBGLCallbackFunc("DisplayFunc")
    Call CurrentContext.SetIdleFunc(AddressOf GameIdleFunc)
    Call VBGLCallbackFunc("KeyboardFunc")
    Call VBGLCallbackFunc("KeyboardUpFunc")
    Call VBGLCallbackFunc("PassiveMotionFunc")
    Call VBGLCallbackFunc("MouseWheelFunc")
    Call AddRenderObject(StartRenderObject)
    Debug.Assert False
    Call CurrentContext.MainLoop()
    MeServer.MapData.ServerStarter.Formula = Empty
    SetUpGameGraphics = True
End Function

' Usually i would put this under RenderMethods.bas,
' But since the calling function in SetUpGameGraphics has to be on the same module and i dont want to move that, it will stay here
Private Sub GameIdleFunc()
    Debug.Print CurrentContext.CurrentWindow.LimitFPS 'DEBUG remove for prod
    Call MeOwner.UpdateServer()
    Call MeOwner.UpdatePlayer()
    Call VBGLCallbackIdleFunc()
End Sub

'===========================
'=========Textures==========
'===========================
Private Function InitTextures(ByVal WB As Workbook) As PropCollection
    Dim SpriteFolder As String
    SpriteFolder = MeServer.MapData.Folder.Value & "\Sprites"

    Dim ToLoad() As String
    ReDim ToLoad(3)
    ToLoad(0) = "Attacks"
    ToLoad(1) = "Fumons"
    ToLoad(2) = "Items"
    ToLoad(3) = "Players"
    
    
    Dim ReturnArr() As VBGLTexture
    Dim i As Long

    For i = 0 To USize(ToLoad)
        Dim Manager As VBGLTextureManager  : Set Manager = VBGLTextureManager.Create(VBGLTextureMergerGrid.Create(False))
        Dim Sheet   As Worksheet           : Set Sheet = WB.Worksheets(ToLoad(i))
        Dim Tex     As VBGLTexture         : Call CreateSpriteTexture(Manager, Sheet, SpriteFolder & "\" & ToLoad(i))

        Manager.Transpose = True
        Set Tex = Manager.CreateTexture(TextureFactory(Manager), Sheet.Name, Empty)
        If IsNothing(Tex) Then
            Debug.Print "Error, SubFolder does not exist for sprite creation: " & ToLoad(i)
        Else
            Call VBGLAdd(ReturnArr, Tex)
        End If
    Next i
    Set InitTextures = PropCollection.Create(ReturnArr)
End Function


Private Sub CreateSpriteTexture(ByVal Manager As VBGLTextureManager, ByVal Sheet As Worksheet, ByVal FolderPath As String)

    Dim Names() As String
    Select Case Sheet.Name
        Case "Attacks"  : Names = ArrayString()
        Case "Items"    : Names = ArrayString()
        Case "Fumons"   : Names = ArrayString("Front", "Back")
        Case "Players"  : Names = ArrayString("Up", "Left", "Down", "Right", "Front", "Back")
        Case Else       : Exit Sub
    End Select
    Dim ColumnCount As Long
    ColumnCount = Usize(Names) + 1
    If ColumnCount = 0 Then ColumnCount = 1

    Dim File     As Object
    Dim fso      As Object : Set fso    = CreateObject("Scripting.FileSystemObject")
    Dim Folder   As Object : Set Folder = fso.GetFolder(FolderPath)
    For Each File In Folder.Files
        Call Manager.LoadFromFileArr(File.Path, 1, ColumnCount, VBGLTextureManagerHelperSetUp.VBGLTextureManagerHelperSetUpGrid, Names)
    Next
    Dim SubFolder As Object
    For Each SubFolder In Folder.SubFolders
        Call CreateSpriteTexture(Manager, Sheet, SubFolder.Path)
    Next SubFolder
End Sub

Private Function TextureFactory(ByVal Manager As VBGLTextureManager) As VBGLTexture
    Set TextureFactory = New VBGLTexture
    With TextureFactory
        .Width           = Manager.MaxWidth
        .Height          = Manager.MaxHeight
        .BPP             = 4
        .InternalFormat  = GL_RGBA
        .Format          = GL_RGBA
        .GLTextureMin    = GL_NEAREST
        .GLTextureMag    = GL_NEAREST
        .GLTextureWrapS  = GL_CLAMP_TO_EDGE
        .GLTextureWrapT  = GL_CLAMP_TO_EDGE
    End With
End Function

'===========================
'===========Misc============
'===========================
Public Function CreateFixedCallable(ByVal Definition As String, ParamArray Args() As Variant) As std_Callable
    Dim ArgsV() As Variant
    ArgsV = Args
    Set CreateFixedCallable = AllCallables.CreateCallableArr(Definition, ArgsV).FixArgs(True)
End Function

Public Function CreateUnFixedCallable(ByVal Definition As String, ParamArray Args() As Variant) As std_Callable
    Dim ArgsV() As Variant
    ArgsV = Args
    Set CreateUnFixedCallable = AllCallables.CreateCallableArr(Definition, ArgsV).FixArgs(False)
End Function


'===========================
'=Repeated Renderfunctions==
'===========================
Public Function CreateList(ByVal Base As Object, _
                           ByVal X As Single, _
                           ByVal Texture As VBGLTexture, _
                           ByVal TextObject As std_Callable, _
                           ByVal NameObject As std_Callable, _
                           ByVal ColorObject As std_Callable) As VBGLList

    Dim Temp As VBGLProperties
    Set Temp = VBGLProperties.Create()
    Temp.Value("X") = X
    Temp.Value("Y") = 1.0!
    Temp.Value("Z") = 0.0!

    Dim WhiteBackground(3) As Single
    WhiteBackground(0) = 1
    WhiteBackground(1) = 1
    WhiteBackground(2) = 1
    WhiteBackground(3) = 1

    Set CreateList = VBGLList.Create(Temp)
    Call CreateList.AddRows(Base.Count)
    Dim i As Long
    For i = 0 To Base.Count
        Dim Text           As String : Let Text           = TextObject.Run(i)
        Dim SubTextureName As String : Let SubTextureName = NameObject.Run(i)
        
        Dim Color()        As Single
        If IsSomething(ColorObject) Then
            Let Color = ColorObject.Run(i)
        Else
            ReDim Color(3)
        End If

        Call CreateList.AddElement(i, CreateSprite(Texture, SubTextureName) , WhiteBackground)
        Call CreateList.AddElement(i, CreateTextBox(Text, Color)            , WhiteBackground)
    Next i

    If Base.Count = -1 Then
        Call CreateList.AddRows(0)
        Call CreateList.AddElement(0, CreateTextBox("No " & TypeName(Base) & " yet"), WhiteBackground)
    End If
End Function

Public Sub UpdateList(ByVal List As VBGLList, _
                      ByVal Base As Object, _
                      ByVal Texture As VBGLTexture, _
                      ByVal TextObject As std_Callable, _
                      ByVal NameObject As std_Callable, _
                      ByVal ColorObject As std_Callable, _
                      ByVal Index As Long)

    Dim SelectedColor() As Single
    ReDim SelectedColor(3)
    SelectedColor(1) = 1
    Dim i As Long
    For i = 0 To Base.Count
        Dim Text           As String : Let Text           = TextObject.Run(i)
        Dim SubTextureName As String : Let SubTextureName = NameObject.Run(i)

        Dim Color()        As Single
        If IsSomething(ColorObject) Then
            Color = ColorObject.Run(Missing, i)
        Else
            ReDim Color(3)
        End If

        Dim Sprite  As VBGLMesh
        Set Sprite  = List.GetElement(i, 0)
        Call Sprite.VAO.Buffer.Update(UpdateSprite(Texture, SubTextureName))

        Dim TextBox As VBGLTextBox
        Set TextBox = List.GetElement(i, 1)
        If i = Index Then
            TextBox.Font(0) = UpdateTextBoxText(TextBox.Font(0), Text, SelectedColor)
        Else
            TextBox.Font(0) = UpdateTextBoxText(TextBox.Font(0), Text, Color)
        End If

        Call TextBox.UpdateData()
    Next i

    If Base.Count = -1 Then
        Set TextBox = List.GetElement(0, 0)
        TextBox.Font(0).Text = "No " & TypeName(Base) & " yet"
    End If
End Sub

Private Function CreateTextBox(ByVal Text As String, Optional ByVal Color As Variant) As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = VBGLTextBox.CreateProperties(2, 3)
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)

    Dim Fonts() As VBGLFont
    ReDim Fonts(0)
    Set Fonts(0) = VBGLFont.Create("PLACEHOLDER", UsedFont)
    Set Fonts(0) = UpdateTextBoxText(Fonts(0), Text, Color)

    Set CreateTextBox = FactoryTextBox.Create(Temp, Fonts)
End Function

Private Function CreateSprite(ByVal Texture As VBGLTexture, ByVal SubTextureName As String) As VBGLMesh
    Dim Shader As VBGLShader  : Set Shader = VBGLPrCoShaderXYTxTy
    Dim Layout As VBGLLayout  : Set Layout = VBGLPrCoLayoutXYTxTy
    Dim Data   As VBGLData    : Set Data   = UpdateSprite(Texture, SubTextureName)

    Set CreateSprite = VBGLMesh.Create(Shader, Layout, Data)
    Call CreateSprite.AddTexture(Texture)
End Function

Private Function UpdateTextBoxText(ByVal Font As VBGLFont, ByVal Text As String, Optional ByVal Color As Variant) As VBGLFont
    Dim BlackColor(3) As Single
    BlackColor(3) = 1
    With Item
        Font.Scalee    = 10
        Font.Text      = Text
        Font.FontColor = IIF(IsMissing(Color), BlackColor, Color)
    End With
    Set UpdateTextBoxText = Font
End Function

Private Function UpdateSprite(ByVal Texture As VBGLTexture, ByVal SubTextureName As String) As VBGLData
    Dim Pos()  As Single      : Let Pos     = VBGLBaFoRectangleXY
    Dim Tex()  As Single      : Let Tex     = Texture.SubTextureID(SubTextureName).GetRectangle()
    Dim Arr()  As Single      : Call VBGLArrayInsert(Arr, Pos, Tex, 2, 2)

    Set UpdateSprite   = VBGLData.CreateSingle(Arr)
End Function

Private Function Missing(Optional ByVal DontPopulateThisVariable As Variant) As Variant
    Missing = DontPopulateThisVariable
End Function