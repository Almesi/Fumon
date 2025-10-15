Attribute VB_Name = "PublicVariables"


Option Explicit

Public MeServer As GameServer
Public FactoryAttack            As Attack
Public FactoryAttacks           As Attacks
Public FactoryElementTypes      As ElementTypes
Public FactoryFight             As Fight
Public FactoryFumon             As Fumon
Public FactoryFumonDefinition   As FumonDefinition
Public FactoryFumons            As Fumons
Public FactoryItem              As Item
Public FactoryItemDefinition    As ItemDefinition
Public FactoryItems             As Items
Public FactoryHumanPlayer       As HumanPlayer
Public FactoryComPlayer         As ComPlayer
Public FactoryWildPlayer        As WildPlayer
Public FactoryGameMap           As GameMap
Public FactoryServer            As GameServer
Public FactoryQuest             As Quest
Public FactoryScript            As Script
Public FactoryTile              As Tile
Public FactoryTiles             As Tiles
Public FactoryTileDefinition    As TileDefinition
Public FactoryUpdateQueue       As UpdateQueue

Public Sub SetupGameServer(ByVal WB As WorkBook)
    If MeServer Is Nothing Then
        Set FactoryAttack          = New Attack
        Set FactoryAttacks         = New Attacks
        Set FactoryElementTypes    = New ElementTypes
        Set FactoryFight           = New Fight
        Set FactoryFumon           = New Fumon
        Set FactoryFumonDefinition = New FumonDefinition
        Set FactoryFumons          = New Fumons
        Set FactoryItem            = New Item
        Set FactoryItemDefinition  = New ItemDefinition
        Set FactoryItems           = New Items
        Set FactoryHumanPlayer     = New HumanPlayer
        Set FactoryComPlayer       = New ComPlayer
        Set FactoryWildPlayer      = New WildPlayer
        Set FactoryGameMap         = New GameMap
        Set FactoryQuest           = New Quest
        Set FactoryScript          = New Script
        Set FactoryTile            = New Tile
        Set FactoryTiles           = New Tiles
        Set FactoryTileDefinition  = New TileDefinition
        Set FactoryServer          = New GameServer
        Set FactoryUpdateQueue     = New UpdateQueue

        Set MeServer                   = FactoryServer
        FactoryServer.WorkBook         = WB
        FactoryServer.Textures         = FactoryServer.InitTextures(WB, WB.Sheets("MapData").Range("B2").Offset(08, 0).Value)
        FactoryServer.ElementTypes     = ElementTypes.Create(WB.Sheets("Fumons").Range("W1"))
        FactoryServer.Quests           = FactoryServer.InitGroup(WB.Sheets("Quests").Range("A2") , Quest)
        FactoryServer.Scripts          = FactoryServer.InitGroup(WB.Sheets("Scripts").Range("A2"), Script)
        FactoryServer.Attacks          = FactoryServer.InitGroup(WB.Sheets("Attacks").Range("A2"), Attack)
        FactoryServer.FumonDefinitions = FactoryServer.InitGroup(WB.Sheets("Fumons").Range("A2") , FumonDefinition)
        FactoryServer.ItemDefinitions  = FactoryServer.InitGroup(WB.Sheets("Items").Range("A2")  , ItemDefinition)
        FactoryServer.Players          = FactoryServer.InitPlayers(WB)
        FactoryServer.Tiles            = FactoryServer.InitGroup(WB.Sheets("Tiles").Range("A2"), TileDefinition)
        FactoryServer.GameMap          = FactoryGameMap.Create(WB.Sheets("Map").Range("A1"), WB.Sheets("MapData").Range("B2"))
        FactoryServer.Updates          = FactoryServer.InitUpdates(WB)
    End If
End Sub