Attribute VB_Name = "Fighting"


Option Explicit

Public Function SpawnWildPlayer(ByVal Spawner As FumonSpawner) As IPlayer
    Dim StartRange         As IRange : Set StartRange         = WildPlayersStart
    'Dim Rng                As IRange : Set Rng                = StartRange.Offset(StartRange.GetSelf.Parent.RowCount, 0)
    Dim Rng                As IRange : Set Rng                = StartRange.Offset(0, 0)
    Dim OffsetToFirstFumon As Long   : Let OffsetToFirstFumon = 4
    Dim Offset             As Long   : Let Offset             = PlayerBase.Offset + SpecificBase.Offset + MoveBase.Offset + OffsetToFirstFumon
    Dim Fumon              As Fumon  : Set Fumon              = Spawner.Spawn(Rng.Offset(0, Offset))

    If IsSomething(Fumon) Then
        With MeServer
            Rng.Offset(0, 0).Value  = .Players.Count + 1
            Rng.Offset(0, 1).Value  = Fumon.Definition.Name
            Rng.Offset(0, 15).Value = AIType.WildAIType
            Set SpawnWildPlayer = IPlayer.Create(Rng, .FumonDefinitions, .ItemDefinitions, .MapData.Maps, .Scripts)
        End With
    End If
End Function

Private Sub DestroyWildPlayer(ByVal Player As IPlayer)
    Dim Rng As IRange
    Set Rng = Player.PlayerBase.Number
    Call Rng.GetSelf.Parent.Rng.Rows(Rng.GetSelf.Row).EntireRow.ClearContents
End Sub

Public Sub StartFightWithWildPlayer(ByVal Player As IPlayer, ByVal Spawner As FumonSpawner)
    Dim WildPlayer As IPlayer
    Set WildPlayer = SpawnWildPlayer(Spawner)
    If IsSomething(WildPlayer) Then
        Dim Index  As Long    : Let Index  = MeServer.Players.Add(WildPlayer)
        Dim Winner As IPlayer : Set Winner = FightMoneyExp(Player, WildPlayer)
        Call MeServer.Players.Remove(Index)
        Call DestroyWildPlayer(WildPlayer)
    End If
End Sub

Public Sub StartFightFromDistance(ByVal Player1 As IPlayer, ByVal Offset As Long, ByVal MessageStart As String, ByVal MessageWin As String, ByVal MessageLose As String)
    Dim Player2 As IPlayer
    Set Player2 = HumanPlayerInFront(Player1, Offset)
    If IsSomething(Player2) Then
        Call DoFightAndWalkBack(Player1, Player2, MessageStart, MessageWin, MessageLose)
    End If
End Sub

Public Function HumanPlayerInFront(ByVal Player As IPlayer, ByVal Offset As Long) As IPlayer
    Dim i As Long
    Dim Human As IPlayer
    For i = 1 To Offset
        Set Human = Player.MoveBase.InFront(i).Player
        If IsSomething(Human) Then
            If TypeName(Human.FightBase.AI) = "HumanAI" Then
                Set HumanPlayerInFront = Human
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub DoFightAndWalkBack(ByVal Player1 As IPlayer, ByVal Player2 As IPlayer, ByVal MessageStart As String, ByVal MessageWin As String, ByVal MessageLose As String)
    Dim x() As Long
    Dim y() As Long
    
    Dim p1Number      As Long        : p1Number = Player1.PlayerBase.Number.Value
    Dim p2Number      As Long        : p2Number = Player2.PlayerBase.Number.Value
    Dim MoveBase      As MoveBase    : Set MoveBase = Player1.MoveBase
    Dim PrevLook      As xlDirection : PrevLook = MoveBase.LookDirection.Value

    Call FindPath(p1Number, p2Number, x, y)
    Call MoveBase.MovePath(x, y)
    Call DoFight(Player1, Player2, MessageStart, MessageWin, MessageLose)

    Dim ReversedX() As Long
    Dim ReversedY() As Long
    Call ReversePath(x, y, ReversedX, ReversedY)
    Call MoveBase.MovePath(ReversedX, ReversedY)
    Call Player1.MoveBase.Look(PrevLook)
End Sub

Public Sub DoFight(ByVal Player1 As IPlayer, ByVal Player2 As IPlayer, ByVal MessageStart As String, ByVal MessageWin As String, ByVal MessageLose As String)
    Dim p1Name      As String: p1Name = Player1.PlayerBase.Name.Value
    Dim p2Name      As String: p2Name = Player2.PlayerBase.Name.Value
    Dim p1AI        As IAI   : Set p1AI   = Player1.FightBase.AI
    Dim p2AI        As IAI   : Set p2AI   = Player2.FightBase.AI

    If TypeName(p1AI) <> "HumanAI" Then Call Say(p1Name, MessageStart) Else Call Say(p2Name, MessageStart)

    If TypeName(p1AI) = "HumanAI" And TypeName(p2AI) <> "HumanAI" Then
        If VBGLFind(Player1.SpecificBase.Beaten, Player2) <> -1 Then
            Exit Sub
        End If
    End If

    Dim Winner As IPlayer
    Set Winner = FightMoneyExp(Player1, Player2)
    If TypeName(p1AI) <> "HumanAI" Then
        If Player1 Is Winner Then Call Say(p1Name, MessageWin) Else Call Say(p1Name, MessageLose)
    Else
        If Player2 Is Winner Then Call Say(p2Name, MessageWin) Else Call Say(p2Name, MessageLose)
    End If
End Sub

Public Function FightMoneyExp(ByVal Player1 As IPlayer, ByVal Player2 As IPlayer) As IPlayer
    Dim CurrentFight As Fight

    Dim Winner As IPlayer
    Dim Loser As IPlayer

    Dim Temp As Fight : Set Temp = New Fight
    Set CurrentFight = Temp.Create(FightsStart, Player1, Player2)
    Call Fight(CurrentFight, Winner, Loser)
    Set CurrentFight = Nothing

    Dim WinnerAI As IAI   : Set WinnerAI = Winner.FightBase.AI
    Dim LoserAI  As IAI   : Set LoserAI  = Loser.FightBase.AI


    If TypeName(LoserAI) <> "HumanAI" Then
        Call IncrementMoney(Winner, Loser.MoveBase.Money.Value)
    End If
    
    If TypeName(WinnerAI) = "HumanAI" And TypeName(LoserAI) <> "HumanAI" Then
        Call Winner.SpecificBase.Beaten.AddUnique(Loser)
    End If
    Set FightMoneyExp = Winner
End Function

Public Sub FightHuman(ByVal Player1 As IPlayer, ByVal Player2 As IPlayer, ByVal MessageStart As String, ByVal MessageWin As String, ByVal MessageLose As String)

End Sub

Public Function AskHumanToFight(ByVal Name As String, ByVal Message As String) As IPlayer
    Dim PreviousInput As VBGLIInput
    Call UpdateQuestion(Name, Message)
    Set PreviousInput = CurrentRenderObject.UserInput
    CurrentRenderObject.UserInput = QuestionBoxInput()
    Call CurrentRenderObject.AddDrawable(QuestionRenderObject)
    Do Until EscapeMsgBox = True
        Call glutMainLoopEvent()
        Call CurrentRenderObject.Loopp
    Loop
    CurrentRenderObject.UserInput = PreviousInput
    Call CurrentRenderObject.RemoveDrawable()
    Call EscapeMsgBox(True)
End Function

Public Sub IncrementMoney(ByVal Player As IPlayer, Value As Long)
    Dim Money As IRange: Set Money = Player.MoveBase.Money
    Money.Value = Money.Value + Value
End Sub

Private Sub Fight(ByVal MyFight As Fight, ByRef Winner As IPlayer, ByRef Loser As IPlayer)
    Dim p1Moves As MoveDecision
    Dim p2Moves As MoveDecision
    Call AddRenderObject(FightRenderObject)
    Do Until FightFinished(MyFight, Winner, Loser)
        Call UpdateFight(MyFight, p1Moves, p2Moves)
        Call MyFight.HandleTurn(Winner, Loser, p1Moves, p2Moves)
    Loop
    Call RemoveRenderObject()
End Sub

Private Function FightFinished(ByVal MyFight As Fight, ByRef Winner As IPlayer, ByRef Loser As IPlayer) As Boolean
    FightFinished = MyFight.Winner.Value <> Empty                             Or _
                    (MyFight.p1.Value      = Empty And MyFight.p1.Value <> 0) Or _
                    IsSomething(Winner)                                       Or _ 
                    IsSomething(Loser)
End Function