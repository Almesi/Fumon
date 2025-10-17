Attribute VB_Name = "InitializeGameProcedures"


Option Explicit

#Const TemporaryBuild = 0

Public MyServer       As GameServer
Public MePlayerHuman  As HumanPlayer
Public MePlayer       As IPlayer
Public MeFighter      As IFighter

Public FactoryTextBoxProperties As VBGLProperties
Public FactoryTextBox As VBGLTextBox

Public Sub InitializeGame()
    Dim CurrentBuild As std_VBProject
    Dim PlayerSettings As IConfig

    Dim Shower As IDestination: Set Shower = Nothing
    Dim Logger As IDestination: Set Logger = std_ImmiedeateDestination.Create()
    
    Set NewErrorHandler = std_ErrorHandler.Create(Shower, Logger)

    Set PlayerSettings = std_ConfigRange.Create(ThisWorkbook.Sheets("Settings"))
    Set CurrentBuild = std_VBProject.Create(ThisWorkbook.VBProject, NewErrorHandler)

    Call CreateContextAndWindow(Logger, Shower)
    If IsNothing(CurrentContext) Then Exit Sub
    Set MyServer = ConnectToServer(PlayerSettings)
    If IsNothing(MyServer) Then Exit Sub
    
    Set MePlayerHuman = GetMyPlayer(MyServer, PlayerSettings)
    If IsNothing(MePlayerHuman) Then Exit Sub
    Set MePlayer  = MePlayerHuman
    Set MeFighter = MePlayerHuman

    #If TemporaryBuild = 1 Then
        If CurrentBuild.IncludeFolder(PlayerSettings.Setting("Source Code"), 0, True, False, True) = CurrentBuild.IS_ERROR Then Exit Sub
    #Else
        'Nothing, as the user build it manually
    #End If

    With VBGLTextBox
        .CharsPerLine   = 128
        .LinesPerPage   = 128
        .Pages          = 1
        .LineOffset     = 0.1!
    End With
    Set FactoryTextBox = VBGLTextBox.Factory()
    Set FactoryTextBoxProperties = VBGLTextBox.CreateProperties(2, 3)
    Call StartGame()
    Call MyServer.Workbook.Close
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
    
    Workbooks.Open(Settings.Setting("Server"))
    Call SetupGameServer(Workbooks(Workbooks.Count))
    Set ConnectToServer = MeServer

    ThisWorkbook.Activate
    Exit Function

    Error:
    Call CurrentContext.ErrorHandler.Raise(std_Error.Create(Err.Source, "Severe", "Couldnt connect to Server", Err.Description, Empty))
    Debug.Print "Couldnt connect to Server"
    Debug.Print Err.Description
    Debug.Print Err.Source
    If IsSomething(CurrentContext) Then
        Call glutDestroyWindow(CurrentContext.CurrentWindow.ID)
        CurrentContext.CurrentWindow = Nothing
    End If
End Function

Private Function GetMyPlayer(ByVal Server As GameServer, ByVal Settings As IConfig) As HumanPlayer
    Dim i As Long
    Dim Name As String
    Name = Settings.Setting("Username")
    For i = 0 To Server.Players.Count
        If Server.Players.Object(i).PlayerBase.Name.Value = Name Then
            Set GetMyPlayer = Server.Players.Object(i)
            Exit For
        End If
    Next i
    If IsNothing(GetMyPlayer) Then MsgBox("Player not registered on Server")
End Function