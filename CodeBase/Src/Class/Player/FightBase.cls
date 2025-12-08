VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FightBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Enum AIType
    HumanAIType = 0
    WildAIType  = 1
    ChildAIType = 2
End Enum

Private p_AI           As IAI
Private p_CurrentValue As IRange
Private p_CurrentMove  As IRange
Private p_Fumons       As Fumons
Private p_Items        As Items

Public Property Let AI(ByVal n_AI                        As IAI)        : Set p_AI           = n_AI           : End Property
Public Property Let CurrentMove(ByVal n_CurrentMove      As IRange)      : Set p_CurrentMove  = n_CurrentMove  : End Property
Public Property Let CurrentValue(ByVal n_CurrentValue    As IRange)      : Set p_CurrentValue = n_CurrentValue : End Property
Public Property Let Fumons(ByVal n_Fumons                As Fumons)     : Set p_Fumons       = n_Fumons       : End Property
Public Property Let Items(ByVal n_Items                  As Items)      : Set p_Items        = n_Items        : End Property

Public Property Get AI()                                 As IAI         : Set AI             = p_AI           : End Property
Public Property Get CurrentMove()                        As IRange       : Set CurrentMove    = p_CurrentMove  : End Property
Public Property Get CurrentValue()                       As IRange       : Set CurrentValue   = p_CurrentValue : End Property
Public Property Get Fumons()                             As Fumons      : Set Fumons         = p_Fumons       : End Property
Public Property Get Items()                              As Items       : Set Items          = p_Items        : End Property

Public Function Create(ByVal Rng As IRange, ByVal FumonDefinitions As PropCollection, ByVal ItemDefinitions As PropCollection) As FightBase
    Set Create = New FightBase
    Dim TempFumons As Fumons: Set TempFumons = New Fumons
    Dim TempItems  As Items : Set TempItems  = New Items
    With Create
        .AI                   = GetAI(Rng.Offset(0, 0).Value)
        .CurrentMove          = Rng.Offset(0, 1)
        .CurrentValue         = Rng.Offset(0, 2)
        .Fumons               = TempFumons.Create(Rng.Offset(0, 3), FumonDefinitions)
        .Items                = TempItems.Create(Rng.Offset(0, 101), ItemDefinitions)
        .CurrentMove.Value    = FightMove.FightMoveNothing
        .CurrentValue.Value   = -1
    End With
End Function

Public Function DecideMove(ByVal MyFight As Fight, ByVal OtherPlayer As FightBase) As MoveDecision
    Set DecideMove = AI.Decide(MyFight, Me, OtherPlayer)
End Function

Public Function AwaitTurn(ByVal MyFight As Fight, ByVal OtherPlayer As FightBase) As MoveDecision
    Set AwaitTurn = AI.AwaitTurn(MyFight, Me, OtherPlayer)
End Function

Private Function GetAI(ByVal Val As Long) As IAI
    Select Case Val
        Case AIType.HumanAIType : Set GetAI = HumanAI
        Case AIType.WildAIType  : Set GetAI = WildAI
        Case AIType.ChildAIType : Set GetAI = ChildAI
    End Select
End Function

Private Sub Class_Initialize()
End Sub