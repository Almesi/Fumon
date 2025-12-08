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
Private p_MapData            As MapData
Private p_ElementTypes       As ElementTypes
Private p_FumonDefinitions   As PropCollection
Private p_ItemDefinitions    As PropCollection
Private p_Players            As PropCollection
Private p_Quests             As PropCollection
Private p_Scripts            As PropCollection
Private p_Attacks            As PropCollection
Private p_Fights             As PropCollection
Private p_Tiles              As PropCollection
Private p_FumonSpawners      As PropCollection

Public Property Let WorkBook(n_WorkBook                   As WorkBook)       : Set p_WorkBook         = n_WorkBook                     : End Property
Public Property Let MapData(n_MapData                     As MapData)        : Set p_MapData          = n_MapData                      : End Property
Public Property Let ElementTypes(n_ElementTypes           As ElementTypes)   : Set p_ElementTypes     = n_ElementTypes                 : End Property
Public Property Let FumonDefinitions(n_FumonDefinitions   As PropCollection) : Set p_FumonDefinitions = n_FumonDefinitions             : End Property
Public Property Let ItemDefinitions(n_ItemDefinitions     As PropCollection) : Set p_ItemDefinitions  = n_ItemDefinitions              : End Property
Public Property Let Players(n_Players                     As PropCollection) : Set p_Players          = n_Players                      : End Property
Public Property Let Quests(n_Quests                       As PropCollection) : Set p_Quests           = n_Quests                       : End Property
Public Property Let Scripts(n_Scripts                     As PropCollection) : Set p_Scripts          = n_Scripts                      : End Property
Public Property Let Attacks(n_Attacks                     As PropCollection) : Set p_Attacks          = n_Attacks                      : End Property
Public Property Let Fights(n_Fight                        As PropCollection) : Set p_Fights           = n_Fights                       : End Property
Public Property Let Tiles(n_Tiles                         As PropCollection) : Set p_Tiles            = n_Tiles                        : End Property
Public Property Let FumonSpawners(n_FumonSpawners         As PropCollection) : Set p_FumonSpawners    = n_FumonSpawners                : End Property

Public Property Get Workbook()                            As WorkBook        : Set WorkBook           = p_WorkBook                     : End Property
Public Property Get MapData()                             As MapData         : Set MapData            = p_MapData                      : End Property
Public Property Get ElementTypes()                        As ElementTypes    : Set ElementTypes       = p_ElementTypes                 : End Property
Public Property Get FumonDefinitions()                    As PropCollection  : Set FumonDefinitions   = p_FumonDefinitions             : End Property
Public Property Get ItemDefinitions()                     As PropCollection  : Set ItemDefinitions    = p_ItemDefinitions              : End Property
Public Property Get Players()                             As PropCollection  : Set Players            = p_Players                      : End Property
Public Property Get Quests()                              As PropCollection  : Set Quests             = p_Quests                       : End Property
Public Property Get Scripts()                             As PropCollection  : Set Scripts            = p_Scripts                      : End Property
Public Property Get Attacks()                             As PropCollection  : Set Attacks            = p_Attacks                      : End Property
Public Property Get Fights()                              As PropCollection  : Set Fights             = p_Fights                       : End Property
Public Property Get Tiles()                               As PropCollection  : Set Tiles              = p_Tiles                        : End Property
Public Property Get FumonSpawners()                       As PropCollection  : Set FumonSpawners      = p_FumonSpawners                : End Property

Public Property Get FumonDefinition(Index        As Long) As FumonDefinition : Set FumonDefinition    = FumonDefinitions.Object(Index) : End Property
Public Property Get ItemDefinition(Index         As Long) As ItemDefinition  : Set ItemDefinition     = ItemDefinitions.Object(Index)  : End Property
Public Property Get Quest(Index                  As Long) As Quest           : Set Quests             = Quests.Object(Index)           : End Property
Public Property Get Player(Index                 As Long) As IPlayer         : Set Player             = Players.Object(Index)          : End Property
Public Property Get Script(Index                 As Long) As Script          : Set Script             = Scripts.Object(Index)          : End Property
Public Property Get Attack(Index                 As Long) As Attack          : Set Attack             = Attacks.Object(Index)          : End Property
Public Property Get Fight(Index                  As Long) As Fight           : Set Fight              = Fights.Object(Index)           : End Property
Public Property Get Tile(Index                   As Long) As TileDefinition  : Set Tile               = Tiles.Object(Index)            : End Property
Public Property Get FumonSpawner(Index           As Long) As FumonSpawner    : Set FumonSpawner       = FumonSpawners.Object(Index)    : End Property

Public Function Create(ByVal WB As WorkBook) As GameServer
    Set Create = New GameServer
    With Create
        .WorkBook         = WB
        .ElementTypes     = InitElementTypes(WB)
        .Quests           = InitQuests(WB)
        .Scripts          = InitScripts(WB)
        .Attacks          = InitAttacks(WB)
        .FumonDefinitions = InitFumonDefinitions(WB, .Attacks)
        .ItemDefinitions  = InitItemDefinitions(WB, .Scripts)
        .FumonSpawners    = InitFumonSpawners(WB, .FumonDefinitions)
        .Tiles            = InitTiles(WB)
        .MapData          = InitMapData(WB, .Tiles)
        .Players          = InitPlayers(WB, .FumonDefinitions, .ItemDefinitions, .MapData.Maps, .Scripts)
    End With
End Function

Public Function GetPlayer(ByVal Name As String) As IPlayer
    Dim i As Long
    For i = 0 To Players.Count
        If Player(i).PlayerBase.Name.Value = Name Then
            Set GetPlayer = Player(i)
            Exit Function
        End If
    Next i
End Function

Public Function GetFight(ByVal Player1 As IPlayer, ByVal Player2 As IPlayer) As Fight
    Dim i As Long
    Dim p1Num As Long : p1Num = Player1.Number.Value
    Dim p2Num As Long : p2Num = Player2.Number.Value
    For i = 0 To USize(Fights)
        If IsSomething(Fight(i)) Then
            If Fight(i).p1.Value = p1Num And Fight(i).p2.Value = p2Num Then
                Set GetFight = Fight(i)
                Exit Function
            End If
        End If
    Next i
End Function


Private Function InitElementTypes(ByVal WB As Workbook) As ElementTypes
    Dim Temp As ElementTypes : Set Temp = New ElementTypes
    Set InitElementTypes = Temp.Create(ElementTypesStart)
End Function
Private Function InitQuests(ByVal WB As Workbook) As PropCollection
    Dim Rng As IRange : Set Rng = QuestsStart
    Set InitQuests = InitGroup(Rng, New Quest)
End Function
Private Function InitScripts(ByVal WB As Workbook) As PropCollection
    Dim Rng As IRange : Set Rng = ScriptsStart
    Set InitScripts = InitGroup(Rng, New Script)
End Function
Private Function InitAttacks(ByVal WB As Workbook) As PropCollection
    Dim Rng As IRange : Set Rng = AttacksStart
    Set InitAttacks = InitGroup(Rng, New Attack)
End Function
Private Function InitFumonDefinitions(ByVal WB As Workbook, ByVal nAttacks As PropCollection) As PropCollection
    Dim Rng As IRange : Set Rng = FumonsStart
    Set InitFumonDefinitions = InitGroup(Rng, New FumonDefinition, nAttacks)
End Function
Private Function InitItemDefinitions(ByVal WB As Workbook, ByVal nScripts As PropCollection) As PropCollection
    Dim Rng As IRange : Set Rng = ItemsStart
    Set InitItemDefinitions = InitGroup(Rng, New ItemDefinition, nScripts)
End Function
Private Function InitFumonSpawners(ByVal WB As Workbook, ByVal nFumonDefinitions As PropCollection) As PropCollection
    Dim Rng As IRange : Set Rng = FumonSpawnersStart
    Set InitFumonSpawners = InitGroup(Rng, New FumonSpawner, nFumonDefinitions)
End Function
Private Function InitPlayers(ByVal WB As Workbook, ByVal FumonDefinitions As PropCollection, ByVal ItemDefinitions As PropCollection, ByVal GameMaps As PropCollection, ByVal Scripts As PropCollection) As PropCollection
    Dim i As Long
    Dim Count As Long
    Dim WS As WorkSheet
    Dim TempArr() As IPlayer

    With PlayersStart
        '2 because -1 for vba offset and the other -1 because Range includes titles, which are not part of the item
        Count = .GetSelf.Parent.RowCount - 2
        ReDim TempArr(Count)
        For i = 0 To Count
            Set TempArr(i) = IPlayer.Create(.Offset(i, 0), FumonDefinitions, ItemDefinitions, GameMaps, Scripts)
        Next i
    End With
    Set InitPlayers = PropCollection.Create(TempArr)
End Function
Private Function InitTiles(ByVal WB As Workbook) As PropCollection
    Dim Rng As IRange : Set Rng = TilesStart
    Set InitTiles = InitGroup(Rng, New TileDefinition)
End Function
Private Function InitMapData(ByVal WB As Workbook, ByVal TileDefinitions As PropCollection) As MapData
    Dim Temp As MapData : Set Temp = New MapData
    Set InitMapData = Temp.Create(WB, MapDataStart, TileDefinitions)
End Function
Private Function InitGroup(ByVal Rng As IRange, ByVal Definition As Object, Optional ByVal Coll As PropCollection = Nothing) As PropCollection
    Dim i As Long

    '2 because -1 for vba offset and the other -1 because Range includes titles, which are not part of the item
    Dim Count     As Long   : Count = Rng.GetSelf.Parent.RowCount - 2
    If Count <> -1 Then
        Dim TempArr() As Object : ReDim TempArr(Count)
        For i = 0 To Count
            If IsSomething(Coll) Then
                Set TempArr(i) = Definition.Create(Rng.Offset(i, 0), Coll)
            Else
                Set TempArr(i) = Definition.Create(Rng.Offset(i, 0))
            End If
        Next i
    End If
    Set InitGroup = PropCollection.Create(TempArr)
End Function


Private Function Missing(Optional ByVal DontPopulateThisVariable) As Variant
    Missing = DontPopulateThisVariable
End Function