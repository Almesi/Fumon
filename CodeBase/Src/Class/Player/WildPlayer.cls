VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WildPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IFighter


Private p_PlayerBase As PlayerBase
Private p_FightBase  As FightBase


Public Property Let PlayerBase(ByVal n_PlayerBase As PlayerBase) : Set p_PlayerBase = n_PlayerBase : End Property
Public Property Let FightBase(ByVal n_FightBase   As FightBase)  : Set p_FightBase  = n_FightBase  : End Property

Public Property Get PlayerBase()                  As PlayerBase  : Set PlayerBase   = p_PlayerBase : End Property
Public Property Get FightBase()                   As FightBase   : Set FightBase    = p_FightBase  : End Property

'==========IFighter==========
Private Property Get IFighter_Number()  As Range  : Set IPlayer_Number       = PlayerBase.Number : End Property
Private Property Get IFighter_Name()    As Range  : Set IPlayer_Name         = PlayerBase.Name   : End Property
Private Property Get IFighter_Fumons()  As Fumons : Set IPlayer_Fumons       = FightBase.Fumons  : End Property
Private Property Get IFighter_Items()   As Items  : Set IPlayer_Items        = FightBase.Items   : End Property

Private Property Get IFighter_PlayerBase() As PlayerBase
    Set IFighter_PlayerBase = p_PlayerBase
End Property

Private Property Get IFighter_FightBase() As FightBase
    Set IFighter_FightBase = p_FightBase
End Property


Private Sub IFighter_DoAI(ByVal MyFight As Fight, ByVal OtherPlayer As IFighter)
    Call FightBase.LetCurrentValue(-1)
    Call FightBase.LetCurrentMove(FightMove.FightMoveNothing)
    Call ComAI(OtherPlayer)
End Sub

Public Function Create(ByVal Pointer As Range) As WildPlayer
    Set Create = New WildPlayer
    With Create
        .PlayerBase        = PlayerBase.Create(Pointer.Offset(0, 0))
        .FightBase         = PlayerBase.Create(Pointer.Offset(0, PlayerBase.Offset + MoveBase.Offset))
    End With
End Function

Private Sub ComAI(ByVal OtherPlayer As IFighter)
    Call FightBase.LetCurrentMove(FightMove.FightMoveAttack)
    Call FightBase.LetCurrentValue(0)
End Sub