VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HumanAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IAI

Private Const WaitTime As Long = 60

Private Function IAI_Decide(ByVal MyFight As Fight, ByVal Caller As FightBase, ByVal OtherPlayer As FightBase) As MoveDecision
    Dim StartTime As Single
    StartTime = Timer
    Do Until Caller.CurrentMove.Value <> FightMove.FightMoveNothing
        Dim Diff As Single
        Diff = Timer - StartTime
        If Diff > WaitTime Then
            Exit Do
        End If
        If MyFight.ImFinished(Caller) Then
            Exit Do
        End If
        Call glutMainLoopEvent()
        Call MyFight.UpdateTimer(Caller, Diff / WaitTime)
        Call CurrentRenderObject.Loopp
    Loop
    Set IAI_Decide = MoveDecision.CreateFromRange(MyFight, Caller, Caller.CurrentMove, Caller.CurrentValue)
    Caller.CurrentMove.Value  = FightMove.FightMoveNothing
    Caller.CurrentValue.Value = -1
End Function

Private Function IAI_AwaitTurn(ByVal MyFight As Fight, ByVal Caller As FightBase, ByVal OtherPlayer As FightBase) As MoveDecision
    Dim n_CurrentMove  As FightMove  : Let n_CurrentMove = -1
    Dim n_CurrentValue As Object     : Set n_CurrentValue = FightMove.FightMoveNothing
    Dim StartTime      As Single     : Let StartTime = Timer
    Dim Diff As Single
    Do Until Diff > WaitTime
        Diff = Timer - StartTime
        If OtherPlayer.FightBase.GetCurrentMove <> FightMove.FightMoveNothing Then
            Exit Do
        End If
        Call UpdateTimer(OtherPlayer, Diff / WaitTime)
        Call CurrentRenderObject.Loopp()
        n_CurrentMove = Caller.CurrentMove.Value
        n_CurrentValue = Caller.CurrentValue.Value
    Loop
    Set IAI_AwaitTurn = MoveDecision.Create(n_CurrentMove, n_CurrentValue)
End Function