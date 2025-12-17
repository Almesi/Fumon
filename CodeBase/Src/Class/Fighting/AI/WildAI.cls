VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WildAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IAI

Private Function IAI_Decide(ByVal MyFight As Fight, ByVal Caller As FightBase, ByVal OtherPlayer As FightBase) As MoveDecision
    Dim FirstFumon As Fumon
    Set FirstFumon = Caller.Fumons.FirstAlive
    If IsSomething(FirstFumon) Then
        Set IAI_Decide = MoveDecision.Create(FightMove.FightMoveAttack, FirstFumon.Attacks.Attack(0), 0)
    Else
        Set IAI_Decide = MoveDecision.Create(FightMove.FightMoveNothing, Nothing, 0)
    End If
End Function

Private Function IAI_AwaitTurn(ByVal MyFight As Fight, ByVal Caller As FightBase, ByVal OtherPlayer As FightBase) As MoveDecision
End Function