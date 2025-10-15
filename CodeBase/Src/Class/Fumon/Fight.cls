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


Private p_p1            As Range
Private p_p2            As Range
Private p_mon1          As Range
Private p_mon2          As Range
Private p_Winner        As Range
Private p_p1Finished    As Range
Private p_p2Finished    As Range

Public Property Let p1(n_p1                       As Range): Set p_p1            = n_p1            : End Property
Public Property Let p2(n_p2                       As Range): Set p_p2            = n_p2            : End Property
Public Property Let mon1(n_mon1                   As Range): Set p_mon1          = n_mon1          : End Property
Public Property Let mon2(n_mon2                   As Range): Set p_mon2          = n_mon2          : End Property
Public Property Let Winner(n_Winner               As Range): Set p_Winner        = n_Winner        : End Property
Public Property Let p1Finished(n_p1Finished       As Range): Set p_p1Finished    = n_p1Finished    : End Property
Public Property Let p2Finished(n_p2Finished       As Range): Set p_p2Finished    = n_p2Finished    : End Property

Public Property Get p1()                          As Range : Set p1              = p_p1            : End Property
Public Property Get p2()                          As Range : Set p2              = p_p2            : End Property
Public Property Get mon1()                        As Range : Set mon1            = p_mon1          : End Property
Public Property Get mon2()                        As Range : Set mon2            = p_mon2          : End Property
Public Property Get Winner()                      As Range : Set Winner          = p_Winner        : End Property
Public Property Get p1Finished()                  As Range : Set p1Finished      = p_p1Finished    : End Property
Public Property Get p2Finished()                  As Range : Set p2Finished      = p_p2Finished    : End Property

Public Function Create(ByVal Player1 As IPlayer, ByVal Player2 As IPlayer) As Fight
    Dim i As Long
    Dim WS As Worksheet
    Set WS = FactoryServer.Workbook.Worksheets("Fights")
    Dim Rng As Range
    
    Set Rng = CheckActiveFight(WS.Range("A1"), Player1.Number, Player2.Number)
    If IsNothing(Rng) Then Set Rng = WS.Range("A1")

    Do Until Rng.Offset(i, 0).Formula = Empty
        i = i + 1
    Loop
    Set Create = New Fight
    With Create
        .p1                  = Rng.Offset(i, 0)
        .p2                  = Rng.Offset(i, 1)
        .mon1                = Rng.Offset(i, 2)
        .mon2                = Rng.Offset(i, 3)
        .Winner              = Rng.Offset(i, 4)
        .p1Finished          = Rng.Offset(i, 5)
        .p2Finished          = Rng.Offset(i, 6)
        .p1.Value            = Player1.Number.Value
        .p2.Value            = Player2.Number.Value
        .mon1.Value          = Player1.Fumons.FirstAlive.Definition.Number
        .mon2.Value          = Player2.Fumons.FirstAlive.Definition.Number
        .p1Finished.Value    = 0
        .p2Finished.Value    = 0
        Call .UpdateTimer(Player1, 1)
        Call .UpdateTimer(Player2, 1)
    End With
End Function

Public Sub Fight(ByVal FightCaller As IPlayer, ByRef WinnerPlayer As IPlayer, ByRef LoserPlayer As IPlayer)
    Call AddRenderObject(FightRenderObject)
    Do Until (Winner.Formula <> Empty) Or (p1.Formula = Empty)
        Call UpdateFight(Me, p1Player, p2Player, p1Fumon, p2Fumon)
        Call GetMoves(FightCaller)

        If p1Fumon.CurrentInitiative.Value > p2Fumon.CurrentInitiative.Value Then
            Call DoRound(p1Player, p2Player)
        ElseIf p1Fumon.CurrentInitiative.Value < p2Fumon.CurrentInitiative.Value Then
            Call DoRound(p2Player, p1Player)
        Else
            If RandomNumber() > 50 Then
                Call DoRound(p1Player, p2Player)
            Else
                Call DoRound(p2Player, p1Player)
            End If
        End If
        p1Finished.Value = 0
        p2Finished.Value = 0
    Loop
    Call RemoveRenderObject()
    If Winner.Formula <> Empty Then
        Set WinnerPlayer = MeServer.Player(Winner.Value)
        Set LoserPlayer  = Loser
    Else
        'Set Fight = 'FightCaller lost
    End If
End Sub

Public Function ImFinished(ByVal Caller As IPlayer) As Boolean
    If Caller Is p1Player Then
        ImFinished = p1Finished > 0
    Else
        ImFinished = p2Finished > 0
    End If
End Function

Public Function MyFumon(ByVal Player As IPlayer) As Fumon
    If Player IS p1Player Then
        Set MyFumon = p1Fumon
    Else
        Set MyFumon = p2Fumon
    End If
End Function

Public Sub UpdateTimer(ByVal Player As IPlayer, ByVal Timer1 As Single)
    Static RealTimer1 As Single
    Static RealTimer2 As Single
    If Player Is p1Player Then
        RealTimer1 = Timer1
    Else
        RealTimer2 = Timer1
    End If
    Dim Temp() As Single
    Temp = UpdateFumonTimer(RealTimer1, RealTimer2)
    Call FumonTime.VAO.Buffer.Update(VBGLData.CreateSingle(Temp))
End Sub

Private Sub DoRound(ByVal First As IPlayer, Second As IPlayer)
    Dim p1Move  As FightMove : Let p1Move  = First.GetCurrentMove
    Dim p2Move  As FightMove : Let p2Move  = Second.GetCurrentMove
    Dim p1Value As Variant   : Set p1Value = First.GetCurrentValue(Me)
    Dim p2Value As Variant   : Set p2Value = Second.GetCurrentValue(Me)

    Call DoAction(First, Second, p1Value, p1Move)
    Call DoAction(Second, First, p2Value, p2Move)
End Sub

Private Sub DoAction(ByVal User As IPlayer, ByVal Victim As IPlayer, ByVal CurrentValue As Variant, ByVal CurrentMove As Long)
    If User Is p1Player Then
        p1Finished.Value = 1
    Else
        p2Finished.Value = 1
    End If
    Select Case CurrentMove
        Case FightMove.FightMoveAttack
            Dim TempAttack As Attack
            Set TempAttack = CurrentValue
            Call Attack(User, Victim, TempAttack)
        Case FightMove.FightMoveFlee
            Call Flee()
        Case FightMove.FightMoveItem
            Dim TempUseItem As Item
            Set TempUseItem = CurrentValue
            Call UseItem(User, Victim, TempUseItem)
        Case FightMove.FightMoveChangeFumon
            Dim TempChangeFumon As Fumon
            Set TempChangeFumon = CurrentValue
            Call ChangeFumon(User, Victim, TempChangeFumon)
    End Select
    Call CheckWinner()
End Sub

' Check if there already exists a fight.
' Function is used for the other player to create the object
Private Function CheckActiveFight(ByVal Rng As Range, ByVal Player1Index As Long, ByVal Player2Index As Long) As Range
    Dim i As Long
    Do Until Rng.Offset(i, 0).Formula = Empty
        If Rng.Offset(i, 1).Value = Player1Index And Rng.Offset(i, 2).Value = Player2Index Then
            Set CheckActiveFight = Rng.Offset(i, 0)
        End If
        i = i + 1
    Loop
End Function

Private Function Loser() As IPlayer
    If Winner.Formula <> Empty Then
        If p1.Value = Winner.Value Then
            Set Loser = p2Player
        Else
            Set Loser = p1Player
        End If
    End IF
End Function

Private Sub CheckWinner()
    If Winner.Value <> Empty Then Exit Sub
    If IsNothing(p1Player.Fumons.FirstAlive) Then
        Winner.Value = p2Player.Number.Value
    End If
    If IsNothing(p2Player.Fumons.FirstAlive) Then
        Winner.Value = p1Player.Number.Value
    End If
End Sub

Private Sub ChangeFumon(ByVal User As IPlayer, ByVal Victim As IPlayer, ByVal NewMon As Fumon)
    Dim Index As Long
    Index = User.Fumons.GetFumonIndex(NewMon)
    If User Is p1Player Then
        Call User.Fumons.Swap(mon1.Value, Index)
        mon1.Value = Index
    Else
        Call User.Fumons.Swap(mon2.Value, Index)
        mon2.Value = Index
    End If
End Sub

Private Sub Attack(ByVal User As IPlayer, ByVal Victim As IPlayer, ByVal UsedAttack As Attack)
    Call UsedAttack.Use(GetFumon(User), GetFumon(Victim))
End Sub

Private Sub UseItem(ByVal User As IPlayer, ByVal Victim As IPlayer, ByVal UsedItem As Item)
    Call UsedItem.Use()
End Sub

Private Function Flee() As IPlayer
    If TypeName(OtherPlayerObject()) = "ComPlayer" Or TypeName(OtherPlayerObject()) = "HumanPlayer" Then
        MsgBox("Cannot flee from a trainer battle")
        Exit Sub
    End If
    Set Flee = OtherPlayerObject()
End Function




Private Sub Class_Terminate()
    If Not p_p1            Is Nothing Then p_p1.Clear
    If Not p_p2            Is Nothing Then p_p2.Clear
    If Not p_mon1          Is Nothing Then p_mon1.Clear
    If Not p_mon2          Is Nothing Then p_mon2.Clear
    If Not p_Winner        Is Nothing Then p_Winner.Clear
End Sub

Private Function p1Player() As IPlayer
    Set p1Player = MeServer.Player(p1.Value)
End Function
Private Function p2Player() As IPlayer
    Set p2Player = MeServer.Player(p2.Value)
End Function
Private Function p1Fumon() As Fumon
    Set p1Fumon = p1Player.Fumons.Fumon(mon1.Value)
End Function
Private Function p2Fumon() As Fumon
    Set p2Fumon = p2Player.Fumons.Fumon(mon2.Value)
End Function

Private Function GetFumon(ByVal Player As IPlayer)
    If Player Is p1Player Then
        Set GetFumon = p1Fumon
    Else
        Set GetFumon = p2Fumon
    End If
End Function

Private Function RandomNumber() As Long
    RandomNumber = Int((100 - 0 + 1) * Rnd + 0)
End Function

Private Sub GetMoves(ByVal FightCaller As IPlayer)
    If FightCaller Is p1Player Then
        Call p1Player.DoAI(Me, p2Player)
        Call GetMovesHelp(p1Player, p2Player)
    Else
        Call GetMovesHelp(p2Player, p1Player)
        Call p2Player.DoAI(Me, p1Player)
    End If
End Sub

Private Function GetFinished(ByVal Player As IPlayer) As Range
    If Player = p1Player Then
        Set GetFinished = p1Finished
    Else
        Set GetFinished = p2Finished
    End If
End Function

Private Sub GetMovesHelp(ByVal MyPlayer As IPlayer, ByVal OtherPlayer As IPlayer)
    If TypeName(OtherPlayer) = "HumanPlayer" Then
        Dim StartTime As Single
        StartTime = Timer
        Dim Diff As Single
        Do Until Diff > 60
            Diff = Timer - StartTime
            If OtherPlayer.GetCurrentMove <> FightMove.FightMoveNothing Then
                Exit Do
            End If
            Call UpdateTimer(Me, Diff / 60)
            Call CurrentRenderObject.Loopp()
        Loop
        GetFinished(OtherPlayer).Value = 1
    Else
        Call OtherPlayer.DoAI(Me, MyPlayer)
    End If
End Sub