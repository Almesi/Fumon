VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Public Enum FightMove
    FightMoveNothing     = 0
    FightMoveAttack      = 1
    FightMoveFlee        = 2
    FightMoveItem        = 3
    FightMoveChangeFumon = 4
End Enum


Private p_p1            As IRange
Private p_p2            As IRange
Private p_mon1          As IRange
Private p_mon2          As IRange
Private p_p1Finished    As IRange
Private p_p2Finished    As IRange
Private p_Winner        As IRange

Public Property Let p1(n_p1                       As IRange): Set p_p1            = n_p1            : End Property
Public Property Let p2(n_p2                       As IRange): Set p_p2            = n_p2            : End Property
Public Property Let mon1(n_mon1                   As IRange): Set p_mon1          = n_mon1          : End Property
Public Property Let mon2(n_mon2                   As IRange): Set p_mon2          = n_mon2          : End Property
Public Property Let p1Finished(n_p1Finished       As IRange): Set p_p1Finished    = n_p1Finished    : End Property
Public Property Let p2Finished(n_p2Finished       As IRange): Set p_p2Finished    = n_p2Finished    : End Property
Public Property Let Winner(n_Winner               As IRange): Set p_Winner        = n_Winner        : End Property

Public Property Get p1()                          As IRange : Set p1              = p_p1            : End Property
Public Property Get p2()                          As IRange : Set p2              = p_p2            : End Property
Public Property Get mon1()                        As IRange : Set mon1            = p_mon1          : End Property
Public Property Get mon2()                        As IRange : Set mon2            = p_mon2          : End Property
Public Property Get p1Finished()                  As IRange : Set p1Finished      = p_p1Finished    : End Property
Public Property Get p2Finished()                  As IRange : Set p2Finished      = p_p2Finished    : End Property
Public Property Get Winner()                      As IRange : Set Winner          = p_Winner        : End Property

Public Function p1Fighter() As IPlayer
    Set p1Fighter = MeServer.Player(p1.Value)
End Function
Public Function p2Fighter() As IPlayer
    Set p2Fighter = MeServer.Player(p2.Value)
End Function
Public Function p1Fumon() As Fumon
    Set p1Fumon = p1FightBase.Fumons.Fumon(mon1.Value)
End Function
Public Function p2Fumon() As Fumon
    Set p2Fumon = p2FightBase.Fumons.Fumon(mon2.Value)
End Function
Private Function p1FightBase() As FightBase
    Set p1FightBase = p1Fighter.FightBase
End Function
Private Function p2FightBase() As FightBase
    Set p2FightBase = p2Fighter.FightBase
End Function

Public Function Create(ByVal Start As IRange, ByVal Player1 As IPlayer, ByVal Player2 As IPlayer) As Fight
    Dim p1Player As Long  : Let p1Player = Player1.PlayerBase.Number.Value
    Dim p2Player As Long  : Let p2Player = Player2.PlayerBase.Number.Value
    Dim Rng      As IRange : Set Rng      = GetActiveFight(Start, p1Player, p2Player)

    Set Create = New Fight
    With Create
        .p1                  = Rng.Offset(0, 0)
        .p2                  = Rng.Offset(0, 1)
        .mon1                = Rng.Offset(0, 2)
        .mon2                = Rng.Offset(0, 3)
        .p1Finished          = Rng.Offset(0, 4)
        .p2Finished          = Rng.Offset(0, 5)
        .Winner              = Rng.Offset(0, 6)
        .p1.Value            = p1Player
        .p2.Value            = p2Player
        .mon1.Value          = Player1.FightBase.Fumons.FirstAliveIndex()
        .mon2.Value          = Player2.FightBase.Fumons.FirstAliveIndex()
        .p1Finished.Value    = 0
        .p2Finished.Value    = 0
        Call .UpdateTimer(Player1.FightBase, 1)
        Call .UpdateTimer(Player2.FightBase, 1)
    End With
End Function

Public Sub HandleTurn(ByRef WinnerPlayer As IPlayer, ByRef LoserPlayer As IPlayer, ByRef p1Moves As MoveDecision, ByRef p2Moves As MoveDecision)
    Call GetMoves(p1Moves, p2Moves)

    Dim Player1First As Boolean
    If p1Moves.Priority > p2Moves.Priority Then
        Player1First = True
    ElseIf p1Moves.Priority < p2Moves.Priority Then
        Player1First = False
    Else
        Dim Speed1 As Long : Speed1 = p1Fumon.CurrentInitiative.Value
        Dim Speed2 As Long : Speed2 = p2Fumon.CurrentInitiative.Value
        Do Until Speed1 <> Speed2
            Speed1 = RandomNumber
        Loop
        Player1First = Speed1 > Speed2
    End If

    If Player1First Then
        Call DoRound(p1Fighter, p2Fighter, p1Moves, p2Moves)
    Else
        Call DoRound(p2Fighter, p1Fighter, p2Moves, p1Moves)
    End If
    p1Finished.Value = 0
    p2Finished.Value = 0
    Call AddExp(p1Fighter, p1Fumon, p2Fumon)
    Call AddExp(p2Fighter, p2Fumon, p1Fumon)
    Call GetNextFumon(p1FightBase, mon1)
    Call GetNextFumon(p2FightBase, mon2)
    Call CheckWinner()
    If Winner.Value <> Empty Then
        Set WinnerPlayer = GetPlayer(Winner.Value)
        Set LoserPlayer  = Loser
    End If
End Sub

Public Function ImFinished(ByVal Caller As FightBase) As Boolean
    If Caller Is p1FightBase Then
        ImFinished = p1Finished.Value > 0
    Else
        ImFinished = p2Finished.Value > 0
    End If
End Function

Public Function MyFumon(ByVal Player As FightBase) As Fumon
    If Player IS p1FightBase Then
        Set MyFumon = p1Fumon
    Else
        Set MyFumon = p2Fumon
    End If
End Function

Public Sub UpdateTimer(ByVal Player As FightBase, ByVal Timer1 As Single)
    Static RealTimer1 As Single
    Static RealTimer2 As Single
    If Player Is p1FightBase Then
        RealTimer1 = Timer1
    Else
        RealTimer2 = Timer1
    End If
    Dim Temp() As Single
    Temp = UpdateFumonTimer(Me, Player, RealTimer1, RealTimer2)
    Call FumonTime.VAO.Buffer.Update(VBGLData.CreateSingle(Temp))
End Sub

' Do your move
    Private Sub DoRound(ByVal First As IPlayer, ByVal Second As IPlayer, ByVal FirstMoves As MoveDecision, ByVal SecondMoves As MoveDecision)
        Call DoAction(First, Second, FirstMoves)
        If IsNothing(First.FightBase.Fumons.FirstAlive) Then Exit Sub
        Call DoAction(Second, First, SecondMoves)
    End Sub
    
    Private Sub DoAction(ByVal User As IPlayer, ByVal Victim As IPlayer, ByVal CurrentMove As MoveDecision)
        Dim Success As Boolean
        If User Is p1Fighter Then
            p1Finished.Value = 1
        Else
            p2Finished.Value = 1
        End If
        Select Case CurrentMove.Move
            Case FightMove.FightMoveAttack      : Success = Attack(User, Victim, CurrentMove.Value)
            Case FightMove.FightMoveFlee        : Success = Flee(User, Victim)
            Case FightMove.FightMoveItem        : Success = UseItem(User, Victim, CurrentMove.Value)
            Case FightMove.FightMoveChangeFumon : Success = ChangeFumon(User, Victim, CurrentMove.Value)
        End Select
        CurrentMove.Success = Success
    End Sub
    
    Private Function ChangeFumon(ByVal User As IPlayer, ByVal Victim As IPlayer, ByVal NewMon As Fumon)
        Dim Index As Long
        Index = User.FightBase.Fumons.GetFumonIndex(NewMon)

        If User Is p1Fighter Then
            mon1.Value = Index
        Else
            mon2.Value = Index
        End If
        ChangeFumon = True
    End Function
    
    Private Function Attack(ByVal User As IPlayer, ByVal Victim As IPlayer, ByVal UsedAttack As Attack)
        Attack = UsedAttack.Use(GetFumon(User), GetFumon(Victim))
    End Function
    
    Private Function UseItem(ByVal User As IPlayer, ByVal Victim As IPlayer, ByVal UsedItem As Item)
        UseItem = UsedItem.Use()
    End Function
    
    Private Function Flee(ByVal User As IPlayer, ByVal Victim As IPlayer) As Boolean
        If TypeName(User.FightBase.AI) = "HumanAI" And TypeName(Victim.FightBase.AI) = "WildAI" Then
            Winner.Value = Victim.PlayerBase.Number.Value
            Flee = True
        Else
            Call Say("GAME", "Cannot flee from a trainer battle")
        End If
    End Function
'

' Check if there already exists a fight.
' Function is used for the other player to create the object
Private Function GetActiveFight(ByVal Rng As IRange, ByVal Player1Index As Long, ByVal Player2Index As Long) As IRange
    Dim i As Long
    Do Until Rng.Offset(i, 0).Value = Empty 
        If Rng.Offset(i, 0).Value = Player1Index And Rng.Offset(i, 1).Value = Player2Index Then
            Exit Do
        End If
        i = i + 1
    Loop
    Set GetActiveFight = Rng.Offset(i, 0)
End Function

Private Function Loser() As IPlayer
    If Winner.Value <> Empty Then
        If p1.Value = Winner.Value Then
            Set Loser = p2Fighter
        Else
            Set Loser = p1Fighter
        End If
    End IF
End Function

Private Sub CheckWinner()
    If Winner.Value <> Empty Then Exit Sub
    If IsNothing(p1FightBase.Fumons.FirstAlive) Then
        Winner.Value = p2Fighter.PlayerBase.Number.Value
    End If
    If IsNothing(p2FightBase.Fumons.FirstAlive) Then
        Winner.Value = p1Fighter.PlayerBase.Number.Value
    End If
End Sub

Private Sub GetNextFumon(ByVal Player As FightBase, ByVal Goal As IRange)
    Dim CurrentMon As Fumon
    Set CurrentMon = Player.Fumons.Fumon(Goal.Value)
    If CurrentMon.Alive Then Exit Sub

    Dim FirstAlive As Fumon
    Set FirstAlive = Player.Fumons.FirstAlive
    If IsSomething(FirstAlive) Then
        Goal.Value = FirstAlive.Definition.Number
    End If
End Sub

Private Sub AddExp(ByVal Player As IPlayer, ByVal Reciever As Fumon, ByVal BeatenFumon As Fumon)
    If TypeName(Player.FightBase.AI) = "HumanAI" Then
        If BeatenFumon.Alive = False Then
            Call Reciever.FinishFight(BeatenFumon)
        End If
    End If
End Sub



Private Function GetPlayer(ByVal Index As Long) As IPlayer
    If p1Fighter.PlayerBase.Number.Value = Index Then
        Set GetPlayer = p1Fighter
    Else
        Set GetPlayer = p2Fighter
    End If
End Function

Private Function GetFumon(ByVal Player As IPlayer)
    If Player Is p1Fighter Then
        Set GetFumon = p1Fumon
    Else
        Set GetFumon = p2Fumon
    End If
End Function

Private Function RandomNumber() As Long
    Randomize Timer
    RandomNumber = Int((100 - 0 + 1) * Rnd + 0)
End Function


Private Sub GetMoves(ByRef p1Moves As MoveDecision, ByRef p2Moves As MoveDecision)
    Set p1Moves = p1FightBase.DecideMove(Me, p2FightBase)
    If TypeName(p1FightBase.AI) = "HumanAI" And TypeName(p2FightBase.AI) = "HumanAI" Then
        Set p2Moves = p2FightBase.AwaitTurn(Me, p1FightBase)
    Else
        Set p2Moves = p2FightBase.DecideMove(Me, p1FightBase)
    End If
End Sub

Private Sub Class_Terminate()
    If IsSomething(p_p1)       Then p_p1.Value = Empty
    If IsSomething(p_p2)       Then p_p2.Value = Empty
    If IsSomething(p_mon1)     Then p_mon1.Value = Empty
    If IsSomething(p_mon2)     Then p_mon2.Value = Empty
    If IsSomething(p_Winner)   Then p_Winner.Value = Empty
    If IsSomething(p1Finished) Then p1Finished.Value = Empty
    If IsSomething(p2Finished) Then p2Finished.Value = Empty
End Sub