Attribute VB_Name = "InitializeGameProcedures"


Option Explicit

#Const TemporaryBuild = 0

Public CurrentGameServer As GameServer
Public MePlayer As IPlayer

Public Sub InitializeGame()
    Dim CurrentBuild As std_VBProject
    Dim PlayerSettings As IConfig

    Set PlayerSettings = std_ConfigRange.Create(ThisWorkbook.Sheets("Settings"))
    Set CurrentBuild = std_VBProject.Create(ThisWorkbook.VBProject)

    Set CurrentGameServer = ConnectToServer(PlayerSettings)
    If CurrentGameServer Is Nothing Then Exit Sub
    
    Set MePlayer = GetMyPlayer(CurrentGameServer, PlayerSettings)
    If MePlayer Is Nothing Then Exit Sub

    #If TemporaryBuild = 1 Then
        If CurrentBuild.IncludeFolder(PlayerSettings.Setting("Source Code"), 0, True, False, True) = CurrentBuild.IS_ERROR Then Exit Sub
    #Else
        'Nothing, as the user build it manually
    #End If
    Call StartGame()
End Sub

Public Function ConnectToServer(Settings As IConfig) As GameServer
    On Error GoTo Error
    
    Workbooks.Open(Settings.Setting("Server"))
    Call SetupGameServer()
    Set ConnectToServer = FumonGame.MeServer

    ThisWorkbook.Activate
    Exit Function

    Error:
    Debug.Print "Couldnt connect to Server"
    Debug.Print Err.Description
    Debug.Print Err.Source
End Function

Public Function GetMyPlayer(Server As GameServer, Settings As IConfig) As IPlayer
    Set GetMyPlayer = Server.GetPlayer(Settings.Setting("Username"))
    If GetMyPlayer Is Nothing Then MsgBox("Player not registered on Server")
End Function