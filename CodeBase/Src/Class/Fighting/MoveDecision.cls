VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MoveDecision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Value As Object
Private p_Move  As FightMove

Public Property Let Move(ByVal n_Move      As FightMove) : Let p_Move  = n_Move  : End Property
Public Property Let Value(ByVal n_Value    As Object)    : Set p_Value = n_Value : End Property

Public Property Get Move()                 As FightMove  : Let Move    = p_Move  : End Property
Public Property Get Value()                As Object     : Set Value   = p_Value : End Property

Public Function Create(ByVal n_Move As FightMove, ByVal n_Value As Object) As MoveDecision
    Set Create = New MoveDecision
    With Create
        .Move  = n_Move
        .Value = n_Value
    End With
End Function

Public Function CreateFromRange(ByVal MyFight As Fight, ByVal MyPlayer As FightBase, ByVal Temp_Move As IRange, ByVal Temp_Value As IRange) As MoveDecision
    Dim n_Move  As FightMove : n_Move = Temp_Move.Value  
    Dim n_Value As Object
    Select Case n_Move
        Case FightMove.FightMoveAttack      : Set n_Value = MyFight.MyFumon(MyPlayer).Attacks.Attack(Temp_Value.Value)
        Case FightMove.FightMoveChangeFumon : Set n_Value = MyPlayer.Fumons.Fumon(Temp_Value.Value)
        Case FightMove.FightMoveItem        : Set n_Value = MyPlayer.Items.Item(Temp_Value.Value)
        Case FightMove.FightMoveFlee        : Set n_Value = Nothing
        Case FightMove.FightMoveNothing     : Set n_Value = Nothing
    End Select
    Set CreateFromRange = Create(n_Move, n_Value)
End Function