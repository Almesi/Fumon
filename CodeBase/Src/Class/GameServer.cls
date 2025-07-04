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
Private p_FumonDefinitions() As FumonDefinition
Private p_ItemDefinitions()  As ItemDefinition
Private p_Players()          As IPlayer
Private p_Quests()           As Quest
Private p_Scripts()          As Script
Private p_Attacks()          As Attack
Private p_Fights()           As Fight

Public Property Let WorkBook(n_WorkBook                   As WorkBook)               : Set p_WorkBook         = n_WorkBook               : End Property
Public Property Let GameMap(n_GameMap                     As GameMap)                : Set p_GameMap          = n_GameMap                : End Property
Public Property Let ElementTypes(n_ElementTypes           As ElementTypes)           : Set p_ElementTypes     = n_ElementTypes           : End Property
Public Property Let FumonDefinitions(n_FumonDefinitions() As FumonDefinition)        :     p_FumonDefinitions = n_FumonDefinitions       : End Property
Public Property Let ItemDefinitions(n_ItemDefinitions()   As ItemDefinition)         :     p_ItemDefinitions  = n_ItemDefinitions        : End Property
Public Property Let Players(n_Players()                   As IPlayer)                :     p_Players          = n_Players                : End Property
Public Property Let Quests(n_Quests()                     As Quest)                  :     p_Quests           = n_Quests                 : End Property
Public Property Let Scripts(n_Scripts()                   As Script)                 :     p_Scripts          = n_Scripts                : End Property
Public Property Let Attacks(n_Attacks()                   As Attack)                 :     p_Attacks          = n_Attacks                : End Property
Public Property Let Fights(n_Fight()                      As Fight)                  :     p_Fight            = n_Fight                  : End Property

Public Property Get Workbook()                          As WorkBook                  : Set WorkBook           = p_WorkBook               : End Property
Public Property Get GameMap()                           As GameMap                   : Set GameMap            = p_GameMap                : End Property
Public Property Get ElementTypes()                      As ElementTypes              : Set ElementTypes       = p_ElementTypes           : End Property
Public Property Get FumonDefinitions()                  As FumonDefinition()         : FumonDefinitions       = p_FumonDefinitions       : End Property
Public Property Get ItemDefinitions()                   As ItemDefinition()          : ItemDefinitions        = p_ItemDefinitions        : End Property
Public Property Get Players()                           As IPlayer()                 : Players                = p_Players                : End Property
Public Property Get Quests()                            As Quest()                   : Quests                 = p_Quests                 : End Property
Public Property Get Scripts()                           As Script()                  : Scripts                = p_Scripts                : End Property
Public Property Get Attacks()                           As Attack()                  : Attacks                = p_Attacks                : End Property
Public Property Get Fights()                            As Fight()                   : Fight                  = p_Fight                  : End Property

Public Property Get FumonDefinition(Index               As Long) As FumonDefinition  : Set FumonDefinition    = p_FumonDefinitions(Index): End Property
Public Property Get ItemDefinition(Index                As Long) As ItemDefinition   : Set ItemDefinition     = p_ItemDefinitions(Index) : End Property
Public Property Get Player(Index                        As Long) As IPlayer          : Set Player             = p_Players(Index)         : End Property
Public Property Get Quest(Index                         As Long) As Quest            : Set Quest              = p_Quests(Index)          : End Property
Public Property Get Script(Index                        As Long) As Script           : Set Script             = p_Scripts(Index)         : End Property
Public Property Get Attack(Index                        As Long) As Attack           : Set Attack             = p_Attacks(Index)         : End Property
Public Property Get Fight(Index                         As Long) As Fight            : Set Fight              = p_Fight(Index)           : End Property

Public Function Create(WB As WorkBook) As GameServer
    Set Create = SetupFactory(WB).SetupFactory2(WB)
    With Create
        .Players = .InitPlayers(WB)
    End With
End Function

Public Function SetupFactory(WB As WorkBook) As GameServer
    Set SetupFactory = New GameServer
    With SetupFactory
        .WorkBook         = WB
        .GameMap          = FumonGame.GameMap.Create(WB.Sheets("Map"), WB.Sheets("MapData").Range("B2"))
        .ElementTypes     = FumonGame.ElementTypes.Create(WB.Sheets("Fumons").Range("W1"))
        .Quests           = InitQuests(WB.Sheets("Quests").Range("A2"), FumonGame.Quest)
        .Scripts          = InitScripts(WB.Sheets("Scripts").Range("A2"), FumonGame.Script)
    End With
End Function

Public Function SetupFactory2(WB As WorkBook) As GameServer
    Me.FumonDefinitions = InitFumonDefinitions(WB.Sheets("Fumons").Range("A2") , FumonGame.FumonDefinition)
    Me.ItemDefinitions  = InitItemDefinitions(WB.Sheets("Items").Range("A2")   , FumonGame.ItemDefinition)
    Me.Attacks          = InitAttacks(WB.Sheets("Attacks").Range("A2"), FumonGame.Attack)
    Set SetupFactory2 = Me
End Function

Public Function InitPlayers(WB As WorkBook) As IPlayer()
    Dim ReservedNames As Variant
    Dim i As Long, Count As Long
    Dim WS As WorkSheet
    Dim TempPlayer As IPlayer
    Dim TempArr() As Iplayer

    Set TempPlayer = New ComPlayer
    With WB.Sheets("NPCs").Range("A2")
        Count = GetCount(.Offset(0, 0))
        ReDim TempArr(Count)
        For i = 0 To Count
            Set TempArr(i) = TempPlayer.Create(.Offset(i, 0))
        Next i
    End With

    Set TempPlayer = New HumanPlayer
    ReservedNames = Array("Map", "MapData", "Tiles", "Items", "Quests", "NPCs", "Attacks", "Fights", "Scripts", "Fumons")
    For Each WS in WB.WorkSheets
        For i = 0 To Ubound(ReservedNames)
            If WS.Name = ReservedNames(i) Then GoTo SkipWS
        Next i
        ReDim Preserve TempArr(ArraySize(TempArr) + 1)
        Set TempArr(ArraySize(TempArr)) = TempPlayer.Create(WS.Range("A1"))
        SkipWS:
    Next WS
    InitPlayers = TempArr
End Function

Public Function InitQuests(Rng As Range, Definition As Quest) As Quest()
    Dim i As Long, Count As Long
    Dim TempArr() As Quest

    Count = GetCount(Rng)
    ReDim TempArr(Count)
    For i = 0 To Count
        Set TempArr(i) = Definition.Create(Rng.Offset(i, 0))
    Next i
    InitQuests = TempArr
End Function

Public Function InitScripts(Rng As Range, Definition As Script) As Script()
    Dim i As Long, Count As Long
    Dim TempArr() As Script

    Count = GetCount(Rng)
    ReDim TempArr(Count)
    For i = 0 To Count
        Set TempArr(i) = Definition.Create(Rng.Offset(i, 0))
    Next i
    InitScripts = TempArr
End Function

Public Function InitAttacks(Rng As Range, Definition As Attack) As Attack()
    Dim i As Long, Count As Long
    Dim TempArr() As Attack

    Count = GetCount(Rng)
    ReDim TempArr(Count)
    For i = 0 To Count
        Set TempArr(i) = Definition.Create(Rng.Offset(i, 0))
    Next i
    InitAttacks = TempArr
End Function

Public Function InitFumonDefinitions(Rng As Range, Definition As FumonDefinition) As FumonDefinition()
    Dim i As Long, Count As Long
    Dim TempArr() As FumonDefinition

    Count = GetCount(Rng)
    ReDim TempArr(Count)
    For i = 0 To Count
        Set TempArr(i) = Definition.Create(Rng.Offset(i, 0))
    Next i
    InitFumonDefinitions = TempArr
End Function

Public Function InitItemDefinitions(Rng As Range, Definition As ItemDefinition) As ItemDefinition()
    Dim i As Long, Count As Long
    Dim TempArr() As ItemDefinition

    Count = GetCount(Rng)
    ReDim TempArr(Count)
    For i = 0 To Count
        Set TempArr(i) = Definition.Create(Rng.Offset(i, 0))
    Next i
    InitItemDefinitions = TempArr
End Function

Public Function GetPlayer(Name As String) As IPlayer
    Dim i As Long
    If UCase(Name) = "ME" Then
        'Set GetPlayer = MePlayer
    Else
        Set GetPlayer = Player(GetPlayerIndex(Name))
    End If
End Function

Public Function GetIndexByName(ObjectName As String ,Name As String) As Long
    Select Case ObjectName
        Case "FumonDefinitions" : MeServer.GetIndex(FumonDefinitions , Name)
        Case "ItemDefinitions"  : MeServer.GetIndex(ItemDefinitions  , Name)
        Case "Players"          : MeServer.GetIndex(Players          , Name)
        Case "Quests"           : MeServer.GetIndex(Quests           , Name)
        Case "Scripts"          : MeServer.GetIndex(Scripts          , Name)
        Case "Attacks"          : MeServer.GetIndex(Attacks          , Name)
        Case "Fight"            : MeServer.GetIndex(Fights           , Name)
    End Select
End Function

Public Function AddFight(Player1 As IPlayer, Player2 As IPlayer)
    ReDim Preserve Fights(ArraySize(Fights) + 1)
    Fights(ArraySize(Fights)) = Fight.Create(Player1, Player2)
End Function

Private Function GetIndex(ObjArr As Variant, Name As String) As Long
    Dim i As Long
    For i = 0 To ArraySize(ObjArr)
        If ObjArr(i).Name = Name Then
            GetIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function GetCount(Rng As Range) As Long
    GetCount = -1
    Do While Rng.Offset(GetCount + 1, 0).Formula <> Empty
        GetCount = GetCount + 1
    Loop
End Function