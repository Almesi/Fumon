VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServerOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private p_UpdatesPlayer  As UpdateQueue
Private p_UpdatesServer  As UpdateQueue
Private p_Workbook       As Workbook
Private p_MapData        As MapData
Private p_Players        As PropCollection
Private p_Scripts        As PropCollection


Public Property Let UpdatesPlayer(n_UpdatesPlayer   As UpdateQueue)      : Set p_UpdatesPlayer  = n_UpdatesPlayer  : End Property
Public Property Let UpdatesServer(n_UpdatesServer   As UpdateQueue)      : Set p_UpdatesServer  = n_UpdatesServer  : End Property
Public Property Let WorkBook(n_WB                   As Workbook)         : Set p_Workbook       = n_WB             : End Property
Public Property Let MapData(n_MapData               As MapData)          : Set p_MapData        = n_MapData        : End Property
Public Property Let Players(n_Players               As PropCollection)   : Set p_Players        = n_Players        : End Property
Public Property Let Scripts(n_Scripts               As PropCollection)   : Set p_Scripts        = n_Scripts        : End Property

Public Property Get UpdatesPlayer()                 As UpdateQueue       : Set UpdatesPlayer    = p_UpdatesPlayer  : End Property
Public Property Get UpdatesServer()                 As UpdateQueue       : Set UpdatesServer    = p_UpdatesServer  : End Property
Public Property Get WorkBook()                      As Workbook          : Set WorkBook         = p_Workbook       : End Property
Public Property Get MapData()                       As MapData           : Set MapData          = p_MapData        : End Property
Public Property Get Players()                       As PropCollection    : Set Players          = p_Players        : End Property
Public Property Get Scripts()                       As PropCollection    : Set Scripts          = p_Scripts        : End Property

Public Function Create(ByVal Server As GameServer) As ServerOwner
    Set Create = New ServerOwner
    With Create
        .WorkBook      = Server.WorkBook
        .MapData       = Server.MapData
        .Players       = Server.Players
        .Scripts       = Server.Scripts
        .UpdatesServer = InitUpdatesServer(.WorkBook, .MapData, .Players, .Scripts)
        .UpdatesPlayer = InitUpdatesPlayer(.WorkBook, .MapData, .Players, .Scripts)
    End With
End Function

Public Sub UpdateServer()
    If MapData.ServerStarter.Formula <> Empty Then
        If MapData.ServerStarter.Value = MePlayer.Number.Value Then
            Call UpdatesServer.Update(MapData.Minute.Value)
            Call MapData.UpdateTime()
        End If
    Else
        MapData.ServerStarter.Value = MePlayer.Number.Value
    End If
End Sub

Public Sub UpdatePlayer()
    Call UpdatesPlayer.Update(MapData.Minute.Value)
End Sub


Private Function InitUpdatesServer(ByVal WB As Workbook, ByVal MapData As MapData, ByVal Players As PropCollection, ByVal Scripts As PropCollection) As UpdateQueue
    Dim i As Long
    Dim Once() As std_Callable

    Set InitUpdatesServer = UpdateQueue.Create()

    MapData.ServerStarter.Formula = Empty 'DEBUG Remove for prod
    If MapData.ServerStarter.Formula <> Empty Then Exit Function

    Dim Temp() As IPlayer
    Call Players.GetObjects(Temp)
    Call MapData.ClearPlayers()

    With WB.Sheets("ServerUpdates").Range("A2")
        Do While .Offset(i, 0).Formula <> Empty
            Call InitUpdatesServer.AddThread(CreateCallables(CreateCallable(.Offset(i, 2), Scripts)), Once, .Offset(i, 1))
            i = i + 1
        Loop
    End With
    Call MapData.SetSpawnPoints(Temp)
End Function

Private Function InitUpdatesPlayer(ByVal WB As Workbook, ByVal MapData As MapData, ByVal Players As PropCollection, ByVal Scripts As PropCollection) As UpdateQueue
    Dim i As Long
    Dim Once() As std_Callable

    Set InitUpdatesPlayer = UpdateQueue.Create()
    With WB.Sheets("PlayerUpdates").Range("A2")
        Do While .Offset(i, 0).Formula <> Empty
            Call InitUpdatesPlayer.AddThread(CreateCallables(CreateCallable(.Offset(i, 2), Scripts)), Once, .Offset(i, 1))
            i = i + 1
        Loop
    End With
End Function

Private Function CreateCallables(ByVal Callable As std_Callable) As std_Callable()
    Dim Arr() As std_Callable
    If Callable.ReturnType = "std_Callable()" Then
        Arr = Callable.Run()
    Else
        ReDim Arr(0)
        Set Arr(0) = Callable
    End If
    CreateCallables = Arr
End Function

Private Function CreateCallable(ByVal Rng As Range, ByVal Scripts As PropCollection) As std_Callable
    Set CreateCallable = AllCallables.CreateCallable(Rng.Value)
End Function

Private Function GetCallType(ByVal Text As String) As vbCallType
    Select Case Text
        Case "vbMethod" : GetCallType = vbMethod
        Case "vbGet"    : GetCallType = vbGet
        Case "vbLet"    : GetCallType = vbLet
        Case "vbSet"    : GetCallType = vbSet
    End Select
End Function