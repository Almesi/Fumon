Attribute VB_Name = "InitializeGameProcedures"


Option Explicit
'===========================
'=========Graphics==========
'===========================
Private RenderStack          As std_Stack

Public StartRenderObject     As VBGLRenderObject
Public OverWorldRenderObject As VBGLRenderObject
Public MapRenderObject       As VBGLRenderObject
Public InventoryRenderObject As VBGLRenderObject
Public FightRenderObject     As VBGLRenderObject
Public FumonRenderObject     As VBGLRenderObject
Public AttackRenderObject    As VBGLRenderObject
Public OptionsRenderObject   As VBGLRenderObject
Public MessageRenderObject   As VBGLRenderObject
Public UsedFont              As VBGLFontLayout
Public GameTextures          As PropCollection


Public MeServer       As GameServer
Public MePlayerHuman  As HumanPlayer
Public MePlayer       As IPlayer
Public MeFighter      As IFighter
Public MeOwner        As ServerOwner
Public AllCallables   As std_AllCallables

Public FactoryTextBoxProperties As VBGLProperties
Public FactoryTextBox           As VBGLTextBox

Public Sub InitializeGame()
    Dim CurrentBuild As std_VBProject
    Dim PlayerSettings As IConfig

    Dim Shower As IDestination: Set Shower = Nothing
    Dim Logger As IDestination: Set Logger = std_ImmiedeateDestination.Create()
    
    Dim NewErrorHandler As std_ErrorHandler
    Set NewErrorHandler = std_ErrorHandler.Create(Shower, Logger)

    Set PlayerSettings = std_ConfigRange.Create(ThisWorkbook.Sheets("Settings"))
    Set CurrentBuild = std_VBProject.Create(ThisWorkbook.VBProject, NewErrorHandler)
    Set AllCallables = std_AllCallables.CreateFromProject(ThisWorkbook.VBProject)

    Call CreateContextAndWindow(Logger, Shower)
    If IsNothing(CurrentContext) Then Exit Sub
    Set MeServer = ConnectToServer(PlayerSettings)
    If IsNothing(MeServer) Then Exit Sub
    
    Set MePlayerHuman = MeServer.Players.ObjectByName(PlayerSettings.Setting("Username"))
    If IsNothing(MePlayerHuman) Then Exit Sub
    Set MePlayer  = MePlayerHuman
    Set MeFighter = MePlayerHuman

    Set FactoryTextBox = New VBGLTextBox
    With FactoryTextBox
        .CharsPerLine   = 128
        .LinesPerPage   = 128
        .Pages          = 1
        .LineOffset     = 0.1!
    End With
    Set FactoryTextBoxProperties = FactoryTextBox.CreateProperties(2, 3)
    Set GameTextures = InitTextures(MeServer.WorkBook)
    Set MeOwner = ServerOwner.Create(MeServer)
    Call SetUpGameGraphics()
    Call MeServer.Workbook.Close
End Sub

Private Sub CreateContextAndWindow(ByVal Logger As IDestination, ByVal Shower As IDestination)
    If IsSomething(CurrentContext) Then Exit Sub
    Set CurrentContext = VBGLContext.Create("C:\Users\deallulic\Documents\GitHub\VBGL\Code\Src\Externals", GLUT_CORE_PROFILE, GLUT_DEBUG, Logger, Shower)
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
    Set ConnectToServer = GameServer.Create(Workbooks(Workbooks.Count))

    ThisWorkbook.Activate
    Exit Function

    Error:
    Call CurrentContext.ErrorHandler.Raise(std_Error.Create(Err.Source, "Severe", "Couldnt connect to Server", Err.Description, Empty))
    If IsSomething(CurrentContext) Then
        Call glutDestroyWindow(CurrentContext.CurrentWindow.ID)
        CurrentContext.CurrentWindow = Nothing
    End If
End Function

Public Function SetUpGameGraphics() As Boolean
    Dim FreetypePath As String: FreetypePath = "C:\Users\deallulic\Documents\GitHub\VBGL\Code\Src\Externals"
    Dim FontPath     As String: FontPath     = "C:\Users\deallulic\Documents\GitHub\VBGL\Code\Res\Fonts\Consolas.ttf"

    Set UsedFont = VBGLFontLayout.Create(FreetypePath, FontPath, 48)
    Set RenderStack = New std_Stack
    Call RenderStack.Add(StartRenderObject)

    Application.EnableCancelKey = xlDisabled
    Call SetForegroundWindow(CurrentContext.CurrentWindow.ID)

    ' In Reversed order, so that addrenderobject will not recieve nothing, as the next renderobject is not created yet
    Set MessageRenderObject   = SetUpMessageGraphics()   : Call MessageRenderObject.AssignColor(1, 1, 1, 1)
    Set OptionsRenderObject   = SetUpOptionsGraphics()   : Call OptionsRenderObject.AssignColor(1, 1, 1, 1)
    Set AttackRenderObject    = SetUpAttackGraphics()    : Call AttackRenderObject.AssignColor(1, 1, 1, 1)
    Set FumonRenderObject     = SetUpFumonGraphics()     : Call FumonRenderObject.AssignColor(1, 1, 1, 1)
    Set FightRenderObject     = SetUpFightGraphics()     : Call FightRenderObject.AssignColor(1, 1, 1, 1)
    Set InventoryRenderObject = SetUpInventoryGraphics() : Call InventoryRenderObject.AssignColor(1, 1, 1, 1)
    Call SetUpTileSet()
    Set MapRenderObject       = SetUpMapGraphics()       : Call MapRenderObject.AssignColor(0, 0, 0, 1)
    Set OverWorldRenderObject = SetUpOverWorldGraphics() : Call OverWorldRenderObject.AssignColor(0, 0, 0, 1)
    Set StartRenderObject     = SetUpStartGraphics()     : Call StartRenderObject.AssignColor(1, 1, 1, 1)

    Set GameTextures = InitTextures(MeServer.WorkBook)
    Call VBGLCallbackFunc("DisplayFunc")
    Call CurrentContext.SetIdleFunc(AddressOf GameIdleFunc)
    Call VBGLCallbackFunc("KeyboardFunc")
    Call VBGLCallbackFunc("KeyboardUpFunc")
    Call VBGLCallbackFunc("PassiveMotionFunc")
    Call VBGLCallbackFunc("MouseWheelFunc")
    Call AddRenderObject(StartRenderObject)
    Call CurrentContext.MainLoop()
    MeServer.MapData.ServerStarter.Formula = Empty
    SetUpGameGraphics = True
End Function

'===========================
'=========Textures==========
'===========================
Private Function InitTextures(ByVal WB As Workbook) As PropCollection
    Dim SpriteFolder As String
    SpriteFolder = MeServer.MapData.Folder.Value & "\Sprites"
    Dim fso As Object
    Dim Folder As Object
    Dim File As Object
    Dim Tex As VBGLTexture
    Dim FileName As String
    
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set Folder = fso.GetFolder(SpriteFolder)
    
    Dim ReturnArr() As VBGLTexture
    ' Loop through each file
    For Each File In Folder.Files
        FileName = Mid(File.Name, 1, InStr(1, File.Name, ".") - 1)
        Set Tex = VBGLTexture.Create(File.Path, FileName)
        Call VBGLAdd(ReturnArr, GetTexturePositions(WB, Tex))
    Next File
    Set InitTextures = PropCollection.Create(ReturnArr)
End Function

Private Function GetTexturePositions(ByVal WB As WorkBook, ByVal Tex As VBGLTexture) As VBGLTexture
    Dim RngI As Range
    Dim RngX As Range
    Dim RngY As Range
    Dim WS As WorkSheet

    Set WS = WB.Worksheets(Tex.Name)
    Tex.SubTextures = GetSubTextures(WS, Tex)
    Set GetTexturePositions = Tex
End Function

Private Function GetSubTextures(ByVal WS As Worksheet, ByVal Tex As VBGLTexture) As VBGLSubTexture()
    Dim i As Long, j As Long
    Dim SpritesPerPlayer As Long
    Select Case WS.Name
        Case "Fumons" : SpritesPerPlayer = 6
        Case "Items"  : SpritesPerPlayer = 1
        Case "Attacks": SpritesPerPlayer = 1
        Case "Tiles"  : SpritesPerPlayer = 15
        Case Else     : SpritesPerPlayer = 6
    End Select

    Dim Pos As Range
    If IsSomething(WS.Cells.Find("TextureIndex")) Then
        Set Pos = WS.Cells.Find("TextureIndex").Offset(1, 0)
    Else
        Set Pos = WS.Cells.Find("Index").Offset(1, 0)
    End If
    Dim NamePos     As Range: Set NamePos   = Ws.Cells.Find("Name")
    Dim Uniques()   As Long : Uniques       = Unique(Pos)
    Dim UniqueCount As Long : UniqueCount   = USize(Uniques) + 1
    Dim Count       As Long : Count         = RangeCount(Pos, xlDown) + 1

    Dim SpriteWidth  As Long : SpriteWidth  = 64
    Dim SpriteHeight As Long : SpriteHeight = Tex.Height / UniqueCount ' Faulty when not all have the same height

    Dim Identifier() As String
    Dim X1()         As Long
    Dim X2()         As Long
    Dim Y1()         As Long
    Dim Y2()         As Long

    Dim Size As Long
    Size = (Count * SpritesPerPlayer) - 1
    ReDim Identifier(Size)
    ReDim X1(Size)
    ReDim X2(Size)
    ReDim Y1(Size)
    ReDim Y2(Size)

    For i = 0 To Count - 1
        For j = 0 To SpritesPerPlayer - 1
            Dim Index As Long
            Index = i * SpritesPerPlayer + j
            Identifier(Index) = NamePos.Offset(i, 0).Value & ExtraName(WS, j)
            X1(Index)         = (j * SpriteWidth)
            X2(Index)         = (j * SpriteWidth) + SpriteWidth
            Y1(Index)         = (Pos.Offset(i, 0).Value * SpriteHeight)
            Y2(Index)         = (Pos.Offset(i, 0).Value * SpriteHeight) + SpriteHeight
        Next j
    Next i
    Dim Factory As VBGLSubTexture
    Set Factory = VBGLSubTexture.CreateFactory(Tex.Width, Tex.Height, True)
    GetSubTextures = Factory.CreateFromArray(Identifier, X1, Y1, X2, Y2)
End Function

Private Function Unique(ByVal Rng As Range) As Long()
    Dim Limit As Long
    Limit = RangeCount(Rng, xlDown)
    Dim Arr() As Long
    ReDim Arr(0)
    Dim i As Long, j As Long
    Arr(0) = Rng.Value
    For i = 1 To Limit
        For j = 0 To USize(Arr)
            If Arr(j) = Rng.Offset(i, 0).Value Then
                GoTo Skip
            End If
        Next j
        Call VBGLAdd(Arr, Rng.Offset(i, 0).Value)
        Skip:
    Next i
    Unique = Arr
End Function

Private Function ExtraName(ByVal WS As Worksheet, ByVal Index As Long) As String
    Select Case WS.Name
        Case "Items", "Attacks"
            ExtraName = Empty
        Case "Tiles"
            ExtraName = CStr(Index)
        Case "Players", "Fumons"
            Select Case Index
                Case 0 : ExtraName = "Up"
                Case 1 : ExtraName = "Left"
                Case 2 : ExtraName = "Down"
                Case 3 : ExtraName = "Right"
                Case 4 : ExtraName = "Front"
                Case 5 : ExtraName = "Back"
            End Select
    End Select
End Function

'===========================
'=========DrawStack=========
'===========================
Public Sub AddRenderObject(ByVal Obj As VBGLRenderObject)
    Call RenderStack.Add(Obj)
    Set CurrentRenderObject = Obj
End Sub

Public Sub RemoveRenderObject()
    Call RenderStack.Delete()
    Set CurrentRenderObject = RenderStack.Value
End Sub

Private Sub GameIdleFunc()
    Debug.Print CurrentContext.CurrentWindow.LimitFPS 'DEBUG remove for prod
    Call MeOwner.UpdateServer()
    Call MeOwner.UpdatePlayer()
    Call VBGLCallbackIdleFunc()
End Sub

'===========================
'===========Misc============
'===========================
Public Function ConvertCallable(ByVal Definition As String, ParamArray Args() As Variant) As VBGLCallable
    Dim ArgsV() As Variant
    ArgsV = Args
    Dim Temp As std_Callable
    Set Temp = AllCallables.CreateCallableArr(Definition, ArgsV)
    With Temp
        Set ConvertCallable = VBGLCallable.CreateArr(.Object, .Name, .CallType, .ArgCount, ArgsV)
    End With
End Function