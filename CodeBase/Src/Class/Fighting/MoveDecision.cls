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

Private p_Value    As Object
Private p_Move     As FightMove
Private p_Priority As Long
Private p_Success  As Boolean

Public Property Let Move(ByVal n_Move          As FightMove) : Let p_Move     = n_Move       : End Property
Public Property Let Value(ByVal n_Value        As Object)    : Set p_Value    = n_Value      : End Property
Public Property Let Priority(ByVal n_Priority  As Long)      : Let p_Priority = n_Priority   : End Property
Public Property Let Success(ByVal n_Success    As Boolean)   : Let p_Success  = n_Success    : End Property

Public Property Get Move()                     As FightMove  : Let Move       = p_Move       : End Property
Public Property Get Value()                    As Object     : Set Value      = p_Value      : End Property
Public Property Get Priority()                 As Long       : Let Priority   = p_Priority   : End Property
Public Property Get Success()                  As Boolean    : Let Success    = p_Success    : End Property

Public Function Create(ByVal n_Move As FightMove, ByVal n_Value As Object, ByVal n_Priority As Long) As MoveDecision
    Set Create = New MoveDecision
    With Create
        .Move     = n_Move
        .Value    = n_Value
        .Priority = n_Priority
    End With
End Function

Public Function CreateFromRange(ByVal MyFight As Fight, ByVal MyPlayer As FightBase, ByVal Temp_Move As IRange, ByVal Temp_Value As IRange) As MoveDecision
    Dim n_Move     As FightMove : n_Move = Temp_Move.Value  
    Dim n_Value    As Object
    Dim n_Priority As Long
    Select Case n_Move
        Case FightMove.FightMoveAttack      : n_Priority = 1: Set n_Value = MyFight.MyFumon(MyPlayer).Attacks.Attack(Temp_Value.Value)
        Case FightMove.FightMoveChangeFumon : n_Priority = 2: Set n_Value = MyPlayer.Fumons.Fumon(Temp_Value.Value)
        Case FightMove.FightMoveItem        : n_Priority = 3: Set n_Value = MyPlayer.Items.Item(Temp_Value.Value)
        Case FightMove.FightMoveFlee        : n_Priority = 4: Set n_Value = Nothing
        Case FightMove.FightMoveNothing     : n_Priority = 0: Set n_Value = Nothing
    End Select
    Set CreateFromRange = Create(n_Move, n_Value, n_Priority)
End Function