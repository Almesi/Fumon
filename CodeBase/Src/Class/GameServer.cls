VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Workbook           As WorkBook
Private p_GameMap            As GameMap
Private p_ElementTypes       As ElementTypes
Private p_Updates            As UpdateQueue
Private p_FumonDefinitions   As PropCollection
Private p_ItemDefinitions    As PropCollection
Private p_Players            As PropCollection
Private p_Quests             As PropCollection
Private p_Scripts            As PropCollection
Private p_Attacks            As PropCollection
Private p_Fights             As PropCollection
Private p_Tiles              As PropCollection
Private p_Textures           As PropCollection

Public Property Let WorkBook(n_WorkBook                   As WorkBook)       : Set p_WorkBook         = n_WorkBook                     : End Property
Public Property Let GameMap(n_GameMap                     As GameMap)        : Set p_GameMap          = n_GameMap                      : End Property
Public Property Let ElementTypes(n_ElementTypes           As ElementTypes)   : Set p_ElementTypes     = n_ElementTypes                 : End Property
Public Property Let Updates(n_Updates                     As UpdateQueue)    : Set p_Updates          = n_Updates                      : End Property
Public Property Let FumonDefinitions(n_FumonDefinitions   As PropCollection) : Set p_FumonDefinitions = n_FumonDefinitions             : End Property
Public Property Let ItemDefinitions(n_ItemDefinitions     As PropCollection) : Set p_ItemDefinitions  = n_ItemDefinitions              : End Property
Public Property Let Players(n_Players                     As PropCollection) : Set p_Players          = n_Players                      : End Property
Public Property Let Quests(n_Quests                       As PropCollection) : Set p_Quests           = n_Quests                       : End Property
Public Property Let Scripts(n_Scripts                     As PropCollection) : Set p_Scripts          = n_Scripts                      : End Property
Public Property Let Attacks(n_Attacks                     As PropCollection) : Set p_Attacks          = n_Attacks                      : End Property
Public Property Let Fights(n_Fight                        As PropCollection) : Set p_Fights           = n_Fights                       : End Property
Public Property Let Tiles(n_Tiles                         As PropCollection) : Set p_Tiles            = n_Tiles                        : End Property
Public Property Let Textures(n_Textures                   As PropCollection) : Set p_Textures         = n_Textures                     : End Property

Public Property Get Workbook()                            As WorkBook        : Set WorkBook           = p_WorkBook                     : End Property
Public Property Get GameMap()                             As GameMap         : Set GameMap            = p_GameMap                      : End Property
Public Property Get ElementTypes()                        As ElementTypes    : Set ElementTypes       = p_ElementTypes                 : End Property
Public Property Get Updates()                             As UpdateQueue     : Set Updates            = p_Updates                      : End Property
Public Property Get FumonDefinitions()                    As PropCollection  : Set FumonDefinitions   = p_FumonDefinitions             : End Property
Public Property Get ItemDefinitions()                     As PropCollection  : Set ItemDefinitions    = p_ItemDefinitions              : End Property
Public Property Get Players()                             As PropCollection  : Set Players            = p_Players                      : End Property
Public Property Get Quests()                              As PropCollection  : Set Quests             = p_Quests                       : End Property
Public Property Get Scripts()                             As PropCollection  : Set Scripts            = p_Scripts                      : End Property
Public Property Get Attacks()                             As PropCollection  : Set Attacks            = p_Attacks                      : End Property
Public Property Get Fights()                              As PropCollection  : Set Fights             = p_Fights                       : End Property
Public Property Get Tiles()                               As PropCollection  : Set Tiles              = p_Tiles                        : End Property
Public Property Get Textures()                            As PropCollection  : Set Textures           = p_Textures                     : End Property

Public Property Get FumonDefinition(Index        As Long) As FumonDefinition : Set FumonDefinition    = FumonDefinitions.Object(Index) : End Property
Public Property Get ItemDefinition(Index         As Long) As ItemDefinition  : Set ItemDefinition     = ItemDefinitions.Object(Index)  : End Property
Public Property Get Quest(Index                  As Long) As Quest           : Set Quests             = Quests.Object(Index)           : End Property
Public Property Get Player(Index                 As Long) As IPlayer         : Set Player             = Players.Object(Index)          : End Property
Public Property Get Script(Index                 As Long) As Script          : Set Script             = Scripts.Object(Index)          : End Property
Public Property Get Attack(Index                 As Long) As Attack          : Set Attack             = Attacks.Object(Index)          : End Property
Public Property Get Fight(Index                  As Long) As Fight           : Set Fight              = Fights.Object(Index)           : End Property
Public Property Get Tile(Index                   As Long) As TileDefinition  : Set Tile               = Tiles.Object(Index)            : End Property
Public Property Get Texture(Index                As Long) As VBGLTexture     : Set Texture            = Textures.Object(Index)         : End Property

Public Function InitTextures(ByVal WB As WorkBook, ByVal FolderPath As String) As PropCollection
    Dim SpriteFolder As String
    SpriteFolder = FolderPath & "\Sprites"
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

Public Function InitPlayers(ByVal Rng As Range) As PropCollection
    Dim i As Long
    Dim Count As Long
    Dim WS As WorkSheet
    Dim TempArr() As IPlayer

    With Rng
        Count = RangeCount(.Offset(0, 0), xlDown)
        ReDim TempArr(Count)
        For i = 0 To Count
            If .Offset(i, 2) = -1 Then
                Set TempArr(i) = HumanPlayer.Create(.Offset(i, 0))
            Else
                Set TempArr(i) = ComPlayer.Create(.Offset(i, 0))
            End If
        Next i
    End With
    Set InitPlayers = PropCollection.Create(TempArr)
End Function

Public Function InitUpdates(ByVal Rng As Range) As UpdateQueue
    Dim i As Long
    Dim Once() As VBGLCallable

    Set InitUpdates = UpdateQueue.Create()

    If GameMap.ServerStarter.Formula <> Empty Then Exit Function

    Dim Temp() As IPlayer
    Call Players.GetObjects(Temp)
    Call GameMap.ClearPlayers()

    With Rng
        Do While .Offset(i, 0).Formula <> Empty
            Call InitUpdates.AddThread(CreateCallables(CreateCallable(.Offset(i, 2))), Once, .Offset(i, 1))
            i = i + 1
        Loop
    End With
    Call GameMap.SetSpawnPoints(Temp)
End Function

Public Function InitGroup(ByVal Rng As Range, ByVal Definition As Object) As PropCollection
    Dim i As Long, Count As Long
    Dim TempArr() As Object

    Count = RangeCount(Rng, xlDown)
    ReDim TempArr(Count)
    For i = 0 To Count
        Set TempArr(i) = Definition.Create(Rng.Offset(i, 0))
    Next i
    Set InitGroup = PropCollection.Create(TempArr)
End Function

Public Function GetFight(ByVal Player1 As IFighter, ByVal Player2 As IFighter) As Fight
    Dim i As Long
    Dim p1Num As Long : p1Num = Player1.PlayerBase.Number.Value
    Dim p2Num As Long : p2Num = Player2.PlayerBase.Number.Value
    For i = 0 To USize(Fights)
        If IsSomething(Fight(i)) Then
            If Fight(i).p1.Value = p1Num And Fight(i).p2.Value = p2Num Then
                Set GetFight = Fight(i)
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub Update()
    If GameMap.ServerStarter.Formula <> Empty Then
        If GameMap.ServerStarter.Value = MePlayer.PlayerBase.Number.Value Then
            Call Updates.Update(GameMap.Minute.Value)
            Call GameMap.UpdateTime()
        End If
    Else
        GameMap.ServerStarter.Value = MePlayer.PlayerBase.Number.Value
    End If
End Sub



Private Function GetTexturePositions(ByVal WB As WorkBook, ByVal Tex As VBGLTexture) As VBGLTexture
    Dim RngI As Range
    Dim RngX As Range
    Dim RngY As Range
    Dim WS As WorkSheet

    Set WS = WB.Worksheets(Tex.Name)
    Tex.SubTextures = GetSubTextures(WS, Tex)
    Set GetTexturePositions = Tex
End Function

Private Function GetSubTextures(WS As Worksheet, Tex As VBGLTexture) As VBGLSubTexture()
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
    Dim NamePos As Range
    If IsSomething(WS.Cells.Find("TextureIndex")) Then
        Set Pos = WS.Cells.Find("TextureIndex").Offset(1, 0)
    Else
        Set Pos = WS.Cells.Find("Index").Offset(1, 0)
    End If
    Set NamePos = WS.Cells.Find("Name").Offset(1, 0)
    Dim Count As Long
    Count = RangeCount(NamePos, xlDown) + 1



    Dim SpriteWidth  As Long : SpriteWidth  = 64
    Dim SpriteHeight As Long : SpriteHeight = Tex.Height / Count ' Faulty when not all have the same height

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
            Identifier(i * SpritesPerPlayer + j) = NamePos.Offset(i, 0).Value & ExtraName(WS, j)
            X1(i * SpritesPerPlayer + j)         = (j * SpriteWidth)
            X2(i * SpritesPerPlayer + j)         = (j * SpriteWidth) + SpriteWidth
            Y1(i * SpritesPerPlayer + j)         = (Pos.Offset(i, 0).Value * SpriteHeight)
            Y2(i * SpritesPerPlayer + j)         = (Pos.Offset(i, 0).Value * SpriteHeight) + SpriteHeight
        Next j
    Next i
    Dim Factory As VBGLSubTexture
    Set Factory = VBGLSubTexture.CreateFactory(Tex.Width, Tex.Height, True)
    GetSubTextures = Factory.CreateFromArray(Identifier, X1, Y1, X2, Y2)
End Function

Private Function RangeCount(ByVal Rng As Range, Direction As XlDirection) As Long
    Dim Result As Long
    Result = -1
    Select Case Direction
        Case xlUp    : If Rng.Offset(-1, 0).Value = Empty Then Result = 0
        Case xlLeft  : If Rng.Offset(0, -1).Value = Empty Then Result = 0
        Case xlDown  : If Rng.Offset(+1, 0).Value = Empty Then Result = 0
        Case xlRight : If Rng.Offset(0, +1).Value = Empty Then Result = 0
    End Select
    If Result <> 0 Then
        Select Case Direction
            Case xlUp  , xlDown  : Result = Rng.End(Direction).Row    - Rng.Row
            Case xlLeft, xlRight : Result = Rng.End(Direction).Column - Rng.Column
        End Select
    End If
    RangeCount = Result
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

Private Function CreateCallables(ByVal Callable As VBGLCallable) As VBGLCallable()
    Dim Arr() As VBGLCallable
    Arr = Callable.Handle
    CreateCallables = Arr
End Function

Private Function CreateCallable(ByVal Rng As Range) As VBGLCallable
    Dim Obj As Object
    Dim Met As String
    Dim Typee As vbCallType
    Dim ArgCount As Long
    Dim Args() As Variant

    Set Obj  = GetObj(Rng.Offset(0, 0).Value)
    Met      = Rng.Offset(0, 1).Value
    Typee    = GetCallType(Rng.Offset(0, 2).Value)
    Args     = GetArgs(Rng.Offset(0, 3))
    ArgCount = USize(Args)
    Set CreateCallable = VBGLCallable.CreateArr(Obj, Met, Typee, ArgCount, Args)
End Function

Private Function GetObj(ByVal Text As String) As Object
    If Text = "Nothing" Then Set GetObj = Nothing
End Function

Private Function GetCallType(ByVal Text As String) As vbCallType
    Select Case Text
        Case "vbMethod" : GetCallType = vbMethod
        Case "vbGet"    : GetCallType = vbGet
        Case "vbLet"    : GetCallType = vbLet
        Case "vbSet"    : GetCallType = vbSet
    End Select
End Function

Private Function GetArgs(ByVal Rng As Range) As Variant()
    Dim i As Long
    Dim Arr() As Variant
    Dim Text As String
    Do Until Rng.Offset(0, i).Formula = Empty
        Text = Rng.Offset(0, i).Value
        If Text Like "*(*)" Then
            Dim LeftBracket As Long    : LeftBracket = Instr(1, Text, "(")
            Dim ScriptName  As String  : ScriptName = VBGLMidP(Text, 1, LeftBracket - 1) 
            Dim Found       As Script  : Set Found = Scripts.ObjectByName(ScriptName)
            Dim Arg         As Variant : Arg = CVar(VBGLMidP(Text, LeftBracket + 1, Len(Text) - 1))
            Call VBGLAdd(Arr, Found.Run(Arg))
        Else
            Call VBGLAdd(Arr, Text)
        End If
        i = i + 1
    Loop
    GetArgs = Arr
End Function