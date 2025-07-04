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
    Attack      = 0
    Flee        = 1
    UseItem     = 2
    ChangeFumon = 3
End Enum


Private p_p1            As Range
Private p_p2            As Range
Private p_mon1          As Range
Private p_mon2          As Range
Private p_CurrentPlayer As Range
Private p_Winner        As Range

Public Property Let p1(n_p1                       As Range): Set p_p1            = n_p1            : End Property
Public Property Let p2(n_p2                       As Range): Set p_p2            = n_p2            : End Property
Public Property Let mon1(n_mon1                   As Range): Set p_mon1          = n_mon1          : End Property
Public Property Let mon2(n_mon2                   As Range): Set p_mon2          = n_mon2          : End Property
Public Property Let CurrentPlayer(n_CurrentPlayer As Range): Set p_CurrentPlayer = n_CurrentPlayer : End Property
Public Property Let Winner(n_Winner               As Range): Set p_Winner        = n_Winner        : End Property

Public Property Get p1()                          As Range : Set p1              = p_p1            : End Property
Public Property Get p2()                          As Range : Set p2              = p_p2            : End Property
Public Property Get mon1()                        As Range : Set mon1            = p_mon1          : End Property
Public Property Get mon2()                        As Range : Set mon2            = p_mon2          : End Property
Public Property Get CurrentPlayer()               As Range : Set CurrentPlayer   = p_CurrentPlayer : End Property
Public Property Get Winner()                      As Range : Set Winner          = p_Winner        : End Property

Public Function Create(Player1 As IPlayer, Player2 As IPlayer) As Fight
    Dim i As Long
    Dim Rng As Range

    Set Rng = CheckActiveFight(WS.Range("A1"), Player1.Number, Player2.Number)
    If Rng Is Nothing Then Set Rng = WS.Range("A1")

    Do Until Rng.Offset(i, 0).Formula = Empty
        i = i + 1
    Loop
    Set Create = New Fight
    With Create
        .p1                  = Rng.Offset(i, 0)
        .p2                  = Rng.Offset(i, 1)
        .mon1                = Rng.Offset(i, 2)
        .mon2                = Rng.Offset(i, 3)
        .CurrentPlayer       = Rng.Offset(i, 4)
        .Winner              = Rng.Offset(i, 5)
        .p1.Value            = Player1.Number
        .p2.Value            = Player2.Number
        .mon1.Value          = Player1.Fumons.FirstFumon
        .mon2.Value          = Player2.Fumons.FirstFumon
        .CurrentPlayer.Value = p1.Value
    End With
End Function

' Check if there already exists a fight.
' Function is used for the other player to create the object
Public Function CheckActiveFight(Rng As Range, Player1Index As Long, Player2Index As Long) As Range
    Dim i As Long
    Do Until Rng.Offset(i, 0).Formula = Empty
        If Rng.Offset(i, 1).Value = Player1Index And Rng.Offset(i, 2).Value = Player2Index Then
            Set CheckActiveFight = Rng.Offset(i, 0)
        End If
        i = i + 1
    Loop
End Function

Public Function Fight(FightCaller As IPlayer) As IPlayer
    Do While (Winner.Formula = Empty) Or (Not CurrentPlayer Is Nothing)
        If FightCaller.Number.Value <> CurrentPlayer.Value And TypeName(MeServer.Player(CurrentPlayer.Value)) = "CompPlayer" Then 'Do ComPlayer turn
            Call MeServer.Player(CurrentPlayer.Value).DoAI(OtherPlayerObject())
        End If 
        If WaitTime(FightCaller, 60) Then Call ChangePlayer() ' If the other players time runs out
        Call FightCaller.DoAI(OtherPlayerObject())
    Loop
    If Not CurrentPlayer Is Nothing Then
        Set Fight = MeServer.Player(Winner.Value)
    End If
End Function

Public Sub DoAction(CurrentValue As Variant, CurrentMove As Long)
    Select Case CurrentMove
        Case FightMove.Attack      : Call Attack(CurrentValue)
        Case FightMove.Flee        : Call Flee()
        Case FightMove.UseItem     : Call UseItem(CurrentValue)
        Case FightMove.ChangeFumon : Call ChangeFumon(CurrentValue)
    End Select
End Sub

Public Function WaitTime(FightCaller As IPlayer, Value As Double) As Boolean
    Dim StartTime As Double
    StartTime = Timer
    Do While Timer - StartTime < Value
        DoEvents
        If CurrentPlayer.Value = FightCaller.Number.Value Then Exit Function
    Loop
    WaitTime = True
End Function

Public Sub ChangePlayer()
    If CurrentPlayer.Value = p1.Value Then
        CurrentPlayer.Value = p2.Value
    Else
        CurrentPlayer.Value = p1.Value
    End If
End Sub

Public Sub Loser() As IPlayer
    If Winner.Value <> Empty Then
        If CurrentPlayer.Value = Winner.Value Then
            Set Loser = CurrentPlayerObject()
        Else
            Set Loser = OtherPlayerObject()
        End If
    End IF
End Sub

Private Function CurrentPlayerObject() As IPlayer
    Set CurrentPlayerObject = MeServer.Player(CurrentPlayer.Value)
End Function

Private Function OtherPlayerObject() As IPlayer
    If CurrentPlayer.Value = p1.Value Then
        OtherPlayerObject = MeServer.Player(p2.Value)
    Else
        OtherPlayerObject = MeServer.Player(p1.Value)
    End If
End Function

Private Sub ChangeFumon(NewMon As Fumon)
    If CurrentPlayer.Value = p1.Value Then
        mon1.Value = NewMon.Number
    Else
        mon2.Value = NewMon.Number
    End If
End Sub

Private Sub Attack(UsedAttack As Attack)
    Call UsedAttack.Use(CurrentPlayerObject, P2Fumon)
End Sub

Private Sub UseItem(UsedItem As Item)
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
    If Not p_CurrentPlayer Is Nothing Then p_CurrentPlayer.Clear
    If Not p_Winner        Is Nothing Then p_Winner.Clear
End Sub