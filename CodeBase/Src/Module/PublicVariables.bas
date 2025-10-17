Attribute VB_Name = "PublicVariables"


Option Explicit

Public MeServer      As GameServer
Public FactoryServer As GameServer

Public Sub SetupGameServer(ByVal WB As WorkBook)
    If IsNothing(MeServer) Then
        Set FactoryServer          = New GameServer

        Set MeServer                   = FactoryServer
        FactoryServer.WorkBook         = WB
        FactoryServer.Textures         = FactoryServer.InitTextures(WB, WB.Sheets("MapData").Range("B2").Offset(08, 0).Value)
        FactoryServer.ElementTypes     = ElementTypes.Create(WB.Sheets("Fumons").Range("W1"))
        FactoryServer.Quests           = FactoryServer.InitGroup(WB.Sheets("Quests").Range("A2") , Quest)
        FactoryServer.Scripts          = FactoryServer.InitGroup(WB.Sheets("Scripts").Range("A2"), Script)
        FactoryServer.Attacks          = FactoryServer.InitGroup(WB.Sheets("Attacks").Range("A2"), Attack)
        FactoryServer.FumonDefinitions = FactoryServer.InitGroup(WB.Sheets("Fumons").Range("A2") , FumonDefinition)
        FactoryServer.ItemDefinitions  = FactoryServer.InitGroup(WB.Sheets("Items").Range("A2")  , ItemDefinition)
        FactoryServer.Players          = FactoryServer.InitPlayers(WB.Sheets("Players").Range("A2"))
        FactoryServer.Tiles            = FactoryServer.InitGroup(WB.Sheets("Tiles").Range("A2"), TileDefinition)
        FactoryServer.GameMap          = GameMap.Create(WB.Sheets("Map").Range("A1"), WB.Sheets("MapData").Range("B2"))
        FactoryServer.Updates          = FactoryServer.InitUpdates(WB.Sheets("ScriptInit").Range("A2"))
    End If
End Sub