Attribute VB_Name = "Fighting"


Option Explicit

Public Sub StartFightFromDistance(ByVal Player1 As IFighter, ByVal Offset As Long, ByVal MessageStart As String, ByVal MessageWin As String, ByVal MessageLose As String)
    Dim Player2 As IFighter
    Set Player2 = HumanPlayerInFront(Player1, Offset)
    If IsSomething(Player2) Then
        Call DoFightAndWalkBack(Player1, Player2, MessageStart, MessageWin, MessageLose)
    End If
End Sub

Public Function HumanPlayerInFront(ByVal Player As ComPlayer, ByVal Offset As Long) As IFighter
    Dim i As Long
    Dim Human As IPlayer
    For i = 1 To Offset
        Set Human = Player.MoveBase.InFront(i).Player
        If IsSomething(Human) Then
            If TypeName(Human) = "HumanPlayer" Then
                HumanPlayerInFront = Human
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub DoFightAndWalkBack(ByVal Player1 As IFighter, ByVal Player2 As IFighter, ByVal MessageStart As String, ByVal MessageWin As String, ByVal MessageLose As String)
    Dim p1Number      As Long        : p1Number = Player1.PlayerBase.Number.Value
    Dim p2Number      As Long        : p2Number = Player2.PlayerBase.Number.Value
    Dim p1Player      As IPlayer     : Set p1Player = Player1
    Dim PrevDirection As xlDirection : PrevDirection = p1Player.LookDirection
    Dim x() As Long
    Dim y() As Long
    Call FindPath(p1Number, p2Number, x, y)
    Call MovePath(p1Number, x, y)
    Call DoFight(Player1, Player2, MessageStart, MessageWin, MessageLose)

    Dim ReversedX() As Long
    Dim ReversedY() As Long
    Call ReversePath(x, y, ReversedX, ReversedY)
    Call MovePath(p1Number, x, y)
    Call MeServer.Player(p1Number).Look(PrevDirection)
End Sub

Public Sub DoFight(ByVal Player1 As IFighter, ByVal Player2 As IFighter, ByVal MessageStart As String, ByVal MessageWin As String, ByVal MessageLose As String)
    Dim p1Name      As Long        : p1Name = Player1.PlayerBase.Name.Value
    Dim p2Name      As Long        : p2Name = Player2.PlayerBase.Name.Value
    If TypeName(Player1) = "ComPlayer" Then Call Say(p1Name, MessageStart) Else Call Say(p2Name, MessageStart)

    If TypeName(Player1) = "HumanPlayer" And TypeName(Player2) = "ComPlayer" Then
        Dim Temp As HumanPlayer
        Set Temp = Player1
        If VBGLFind(Temp.Beaten, Player2) <> -1 Then
            Exit Sub
        End If
    End If

    Dim Winner As IFighter
    Set Winner = Fight(Player1, Player2)
    If TypeName(Player1) = "ComPlayer" Then
        If Player1 Is Winner Then Call Say(p1Name, MessageWin) Else Call Say(p1Name, MessageLose)
    Else
        If Player2 Is Winner Then Call Say(p2Name, MessageWin) Else Call Say(p2Name, MessageLose)
    End If
End Sub

Public Function Fight(ByVal Player1 As IFighter, ByVal Player2 As IFighter) As IFighter
    Dim CurrentFight As Fight

    Dim Winner As IFighter
    Dim Loser As IFighter

    Dim Temp As Fight : Set Temp = New Fight
    Set CurrentFight = Temp.Create(MeServer.Workbook.Worksheets("Fights").Range("A1"), Player1, Player2)
    Call CurrentFight.Fight(Player1, Winner, Loser)
    Set CurrentFight = Nothing

    If TypeName(Loser) = "ComPlayer" Then
        Dim TempLoser As ComPlayer
        Set TempLoser = Loser
        Call IncrementMoney(Winner, TempLoser.MoveBase.Money.Value)
    End If
    
    If TypeName(Winner) = "HumanPlayer" And TypeName(Loser) = "ComPlayer" Then
        Dim TempWinner As HumanPlayer
        Set TempWinner = Winner
        Call TempWinner.AddBeaten(TempLoser)
    End If
    Set Fight = Winner
End Function

Public Function AskHumanToFight() As IFighter

End Function

Public Sub IncrementMoney(ByVal Player As HumanPlayer, Value As Long)
    Dim Money As Range: Set Money = Player.MoveBase.Money
    Money.Value = Money.Value + Value
End Sub