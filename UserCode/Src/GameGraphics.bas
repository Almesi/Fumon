Attribute VB_Name = "GameGraphics"


Option Explicit

Private GameRunning As Boolean
Private DeltaTime As Long

Private Window          As std_Window

Private TileTexture     As std_Texture
Private NPCTexture      As std_Texture
Private ItemTexture     As std_Texture
Private FumonTexture    As std_Texture

Private TileBuffer      As std_Buffer
Private NPCBuffer       As std_Buffer
Private ItemBuffer      As std_Buffer

Private TileVertexArray As std_Buffer
Private NPCVertexArray  As std_Buffer
Private ItemVertexArray As std_Buffer

Private TileIndexBuffer As std_Buffer
Private NPCIndexBuffer  As std_Buffer
Private ItemIndexBuffer As std_Buffer

Private Renderer        As std_Renderer
Private Shader          As std_Shader


Private Const FPS           As Long = 60
Private Const ScreenSpriteX As Long = 16
Private Const ScreenSpriteY As Long = 09
Private Const SpriteWidth   As Long = 32
Private Const SpriteHeight  As Long = 32

Public CurrentScreenType As ScreenTypes
Public Enum ScreenTypes
    OverWorld = 0
    Inventory = 1
    Fight     = 2
    WorldMap  = 3
End Enum

Public Function StartGame() As Boolean
    Dim FilePath As String
    FilePath = FumonGame.MeServer.GameMap.Folder.Value
    If LoadLibrary(FilePath & "\Freeglut64.dll") = False Then
        MsgBox("Couldnt load freeglut")
        Exit Function
    End If

    Call glutInit(0&, "")
    Set Window = New std_Window
    Set Window = std_Window.Create(SpriteHeight * ScreenSpriteY, _
                                   SpriteWidth  * ScreenSpriteX, _
                                   "Fumon", _
                                   "4_6", _
                                   GLUT_CORE_PROFILE, _
                                   GLUT_DEBUG)

    Call GLStartDebug()

    Call glEnable(GL_BLEND)
    Call glEnable(GL_DEPTH_TEST)
    Call glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)

    Call glFrontFace(GL_CW)

    Set Renderer = New std_Renderer
    Set Shader = std_Shader.CreateFromFile(FilePath & "\Vertex.Shader", FilePath & "\Fragment.Shader")

    Set TileTexture  = std_Texture.Create(FilePath & "\Tile.png")
    Set NPCTexture   = std_Texture.Create(FilePath & "\NPC.png")
    Set ItemTexture  = std_Texture.Create(FilePath & "\Item.png")
    Set FumonTexture = std_Texture.Create(FilePath & "\Fumon.png")

    With std_BufferLayoutType
        Call UpdateVertexArray(TileVertexArray, GetMapData("Tile") , GL_ARRAY_BUFFER, TileBuffer, .XYZ, .RedGreenBlue, .TextureXTextureY)
        Call UpdateVertexArray(NPCVertexArray , GetMapData("NPC")  , GL_ARRAY_BUFFER, NPCBuffer , .XYZ, .RedGreenBlue, .TextureXTextureY)
        Call UpdateVertexArray(ItemVertexArray, GetItemData()      , GL_ARRAY_BUFFER, ItemBuffer, .XYZ, .RedGreenBlue, .TextureXTextureY)
    End With

    Call glutDisplayFunc(AddressOf GameLoop)
    Call glutIdleFunc(AddressOf GameLoop)
    Call glutKeyboardFunc(AddressOf KeyPressed)

    Call glutMainLoop
End Function

Public Sub GameLoop()
    Dim ScreenDataType As ScreenTypes
    Dim ScreenData() As Variant
    GameRunning = True
    Do While GameRunning
        Call UpdateScreen()
        DeltaTime = DeltaTime + 1
        If DeltaTime > FPS Then DeltaTime = 0
        Call WaitFor(1 / FPS)
    Loop
End Sub

Public Sub UpdateScreen()
    Call Renderer.Clear()
    Select Case CurrentScreenType
        Case ScreenTypes.OverWorld
            Call UpdateVertexArray(TileVertexArray, GetMapData("Tile") , GL_ARRAY_BUFFER, TileBuffer, 3, 3, 2)
            Call Renderer.Draw(TileVertexArray, Nothing, Shader)
            Call UpdateVertexArray(NPCVertexArray , GetMapData("NPC")  , GL_ARRAY_BUFFER, NPCBuffer , 3, 3, 2)
            Call Renderer.Draw(NPCVertexArray, Nothing, Shader)
        Case ScreenTypes.Inventory
            Call UpdateVertexArray(ItemVertexArray, GetItemData()      , GL_ARRAY_BUFFER, ItemBuffer, 3, 3, 2)
            Call Renderer.Draw(ItemVertexArray, Nothing, Shader)
        Case ScreenTypes.Fight

        Case ScreenTypes.WorldMap

    End Select
End Sub

Private Sub UpdateVertexArray(VA As std_VertexArray, Mesh As std_Mesh, BufferType As Long, Buffer As std_Buffer, ParamArray VertexAttributes As Variant)
    Dim i As Long
    Dim VertexSize As Long

    If VA Is Nothing Then Set VA = New std_VertexArray
    Dim VBLayout As New std_BufferLayout
    For i = 0 To ArraySize(VertexAttributes)
        Call VBLayout.AddFloat(VertexAttributes(i))
        VertexSize = VertexSize + VertexAttributes(i)
    Next i

    Set Buffer = std_Buffer.Create(BufferType, Mesh)
    Call VA.AddBuffer(Buffer, VBLayout)
End Sub

Private Function GetMapData(GetWhat As Long) As std_Mesh

    Dim MaxY       As Long  : MaxY        = FumonGame.MeServer.GameMap.Rows.Value
    Dim MaxX       As Long  : MaxX        = FumonGame.MeServer.GameMap.Columns.Value
    Dim MapData()  As Long  : MapData     = GetAreaData(0, 0, MaxX, MaxY, GetWhat)
    Dim VertexSize As Long  : VertexSize  = 8
    Dim NewSize    As Long  : NewSize     = ((ArraySize(MapData) + 1) * VertexSize) - 1
    Dim Vertices() As Single: ReDim Vertices(NewSize)
    Dim i As Long, x As Long, y As Long
    
    For i = 0 To ArraySize(MapData) Step +VertexSize
        x = (i Mod MaxX)
        y = (Int(i / MaxX))
        With FumonGame.MeServer.GameMap
            Vertices(i + 0) = x
            Vertices(i + 1) = y
            Vertices(i + 2) = 0
            Vertices(i + 3) = 1
            Vertices(i + 4) = 1
            Vertices(i + 5) = 1
            Vertices(i + 6) = .GetTile(y, x).TextureDefinition.X
            Vertices(i + 7) = .GetTile(y, x).TextureDefinition.Y
        End With
    Next i

    Dim Indices() As Single
    Indices = GetIndexBufferFrom2DArray(MaxX, MaxY)

    Dim TempVertex As New std_Mesh: Call TempVertex.AssignData(VarPtr(Vertices(0)), VarType(Vertices(0)), LenB(Vertices(0)), ArraySize(Vertices), VertexSize)
    Dim TempIndex  As New std_Mesh: Call TempIndex.AssignData(VarPtr(Indices(0))  , VarType(Indices(0)) , LenB(Indices(0)) , ArraySize(Indices) , 3)
    Dim ReturnMesh As New std_Mesh
    Set GetMapData = ReturnMesh.CreateMeshFromIndex(TempVertex, TempIndex)
End Function

Private Function GetItemData() As std_Mesh
    Dim VertexSize As Long : VertexSize  = 8 'x,y,z,r,g,b,tx,ty
    Dim i As Long

    Dim NewSize    As Long
    NewSize     = ((ArraySize(MePlayer.Items.Items) + 1) * VertexSize * 2) - 1

    Dim Vertices() As Single
    ReDim Vertices(NewSize)
    For i = 0 To ArraySize(MePlayer.Items.Items)
        Vertices(i * VertexSize + 00) = -1
        Vertices(i * VertexSize + 01) = 1 / (ArraySize(MapData) * i + 1) ' +1 to avoid division by 0
        Vertices(i * VertexSize + 02) = 0
        Vertices(i * VertexSize + 03) = 1
        Vertices(i * VertexSize + 04) = 1
        Vertices(i * VertexSize + 05) = 1
        Vertices(i * VertexSize + 06) = MePlayer.Items.Item(i).ItemDefinition.TextureDefinition.X
        Vertices(i * VertexSize + 07) = MePlayer.Items.Item(i).ItemDefinition.TextureDefinition.Y

        Vertices(i * VertexSize + 10) = 0
        Vertices(i * VertexSize + 11) = 1 / (ArraySize(MapData) * i + 1) ' +1 to avoid division by 0
        Vertices(i * VertexSize + 12) = 0
        Vertices(i * VertexSize + 13) = 1
        Vertices(i * VertexSize + 14) = 1
        Vertices(i * VertexSize + 15) = 1
        Vertices(i * VertexSize + 16) = 0 ' Font Texture here TODO
        Vertices(i * VertexSize + 17) = 0 ' Font Texture here TODO
    Next i

    Dim Indices()  As Single
    Indices = GetIndexBufferFrom2DArray(1, ArraySize(PlayerItems))

    Dim TempVertex As New std_Mesh: Call TempVertex.AssignData(VarPtr(Vertices(0)), VarType(Vertices(0)), LenB(Vertices(0)), ArraySize(Vertices), VertexSize)
    Dim TempIndex  As New std_Mesh: Call TempIndex.AssignData(VarPtr(Indices(0))  , VarType(Indices(0)) , LenB(Indices(0)) , ArraySize(Indices) , 3)
    Dim ReturnMesh As New std_Mesh
    Set GetItemData = ReturnMesh.CreateMeshFromIndex(TempVertex, TempIndex)
End Function

Private Function GetIndexSizeFromSheet(MaxX As Long, MaxY As Long) As Long
    Dim Twos   As Long :Twos   = 4                           * 2' 4 Corners
    Dim Threes As Long :Threes = ((MaxX - 1)) + ((MaxY - 1)) * 3' 4 Side
    Dim Fours  As Long :Fours  = ((MaxX - 1)) * ((MaxY - 1)) * 4' The Rest in the Middle
    GetIndexSizeFromSheet = Twos + Threes + Fours
End Function

Private Function GetIndexBufferFrom2DArray(MaxX As Long, MaxY As Long) As Single()
    Dim Indices() As Single
    Dim i As Long
    Dim x As Long, y As Long
    Dim Size As Long
    
    Size = MaxX * MaxY - 1
    ReDim Indices(GetIndexSizeFromSheet(MaxX, MaxY) - 1)
    For i = 0 To Size
        x = (i Mod MaxX) + 1
        y = (Int(i / MaxX)) + 1
        Indices(i * 3 + 0) = x
        Indices(i * 3 + 1) = x + 1
        Indices(i * 3 + 2) = y * MaxX + i + 1
    Next i
End Function

Private Function GetAreaData(EndX As Long, EndY As Long, GetWhat As String) As Long()
    Dim Size As Long
    Dim y As Long, x As Long
    Dim CurrentIndex As Long
    Dim Data() As Long

    Size = (EndX + 1) * (EndY + 1) - 1
    ReDim Data(Size)
    With FumonGame.MeServer.GameMap
        For y = 0 To EndX
            For x = 0 To EndY
                Data(CurrentIndex) = CLng(.ExtractPoint(MapPointer.Offset(y, x).Value, GetWhat))
                CurrentIndex = CurrentIndex + 1
            Next x
        Next y
    End With
    GetAreaData = Data
End Function

Private Sub WaitFor(NumOfSeconds As Double)
    Dim SngSec As Double
    SngSec = Timer + NumOfSeconds

    Do While Timer < SngSec
        DoEvents
   Loop
End Sub