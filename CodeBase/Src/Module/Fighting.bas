Attribute VB_Name = "Fighting"


Option Explicit

Public Sub StartFightFromDistance(ByVal Player1 As IPlayer, ByVal Offset As Long, ByVal MessageStart As String, ByVal MessageWin As String, ByVal MessageLose As String)
    Dim Player2 As IPlayer
    Set Player2 = HumanPlayerInFront(Player1, Offset)
    If IsSomething(Player2) Then
        Call DoFightAndWalkBack(Player1, Player2, MessageStart, MessageWin, MessageLose)
    End If
End Sub

Public Function HumanPlayerInFront(ByVal Player As ComPlayer, ByVal Offset As Long) As IPlayer
    Dim i As Long
    Dim Human As IPlayer
    For i = 1 To Offset
        Set Human = Player.InFront(i)
        If IsSomething(Human) Then
            If TypeName(Human) = "HumanPlayer" Then
                HumanPlayerInFront = Human
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub DoFightAndWalkBack(ByVal Player1 As IPlayer, ByVal Player2 As IPlayer, ByVal MessageStart As String, ByVal MessageWin As String, ByVal MessageLose As String)
    Dim x() As Long
    Dim y() As Long
    Dim PrevDirection As xlDirection
    PrevDirection = MeServer.Player(Player1.Number).LookDirection
    Call FindPath(Player1.Number, Player2.Number, x, y)
    Call MovePath(Player1.Number, x, y)
    Call DoFight(Player1, Player2, MessageStart, MessageWin, MessageLose)

    Dim ReversedX() As Long
    Dim ReversedY() As Long
    Call ReversePath(x, y, ReversedX, ReversedY)
    Call MovePath(Player1.Number, x, y)
    Call MeServer.Player(Player1.Number).Look(PrevDirection)
End Sub

Public Sub DoFight(ByVal Player1 As IPlayer, ByVal Player2 As IPlayer, ByVal MessageStart As String, ByVal MessageWin As String, ByVal MessageLose As String)
    If TypeName(Player1) = "ComPlayer" Then
        Call Say(Player1.Name, MessageStart)
    Else
        Call Say(Player2.Name, MessageStart)
    End If
    If TypeName(Player1) = "HumanPlayer" And TypeName(Player2) = "ComPlayer" Then
        Dim Temp As HumanPlayer
        Set Temp = Player1
        If VBGLFind(Temp.Beaten, Player2) <> -1 Then
            Exit Sub
        End If
    End If
    Dim Winner As IPlayer
    Set Winner = Fight(Player1, Player2)
    If TypeName(Player1) = "ComPlayer" Then
        If Player1 Is Winner Then
            Call Say(Player1.Name, MessageWin)
        Else
            Call Say(Player1.Name, MessageLose)
        End If
    Else
        If Player2 Is Winner Then
            Call Say(Player2.Name, MessageWin)
        Else
            Call Say(Player2.Name, MessageLose)
        End If
    End If
End Sub

Public Function Fight(ByVal Player1 As IPlayer, ByVal Player2 As IPlayer) As IPlayer
    Dim CurrentFight As Fight

    Dim Winner As IPlayer
    Dim Loser As IPlayer

    Set CurrentFight = FactoryFight.Create(Player1, Player2)
    Call CurrentFight.Fight(Player1, Winner, Loser)
    Set CurrentFight = Nothing

    If TypeName(Loser) = "ComPlayer" Then
        Dim TempLoser As ComPlayer
        Set TempLoser = Loser
        Call IncrementMoney(Winner, TempLoser.Money.Value)
    End If
    
    If TypeName(Winner) = "HumanPlayer" And TypeName(Loser) = "ComPlayer" Then
        Dim TempWinner As HumanPlayer
        Set TempWinner = Winner
        Call TempWinner.AddBeaten(TempLoser)
    End If
    Set Fight = Winner
End Function

Public Function AskHumanToFight() As IPlayer

End Function

Public Sub IncrementMoney(ByVal Player As HumanPlayer, Value As Long)
    Dim Money As Range: Set Money = Player.Money
    Money.Value = Money.Value + Value
End Sub