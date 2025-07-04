Attribute VB_Name = "PublicVariables"


Option Explicit

Public MeServer As GameServer
Public FactoryServer As GameServer

Public Sub SetupGameServer()
    If MeServer Is Nothing Then
        Set FactoryServer = GameServer.SetupFactory(ThisWorkbook)
        Set FactoryServer = FactoryServer.SetupFactory2(ThisWorkbook)
        Set MeServer      = FactoryServer.Create(ThisWorkbook)
    End If
End Sub