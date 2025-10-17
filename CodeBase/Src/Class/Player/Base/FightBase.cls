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

Public p_CurrentValue As Range
Public p_CurrentMove  As Range
Public p_Fumons       As Fumons
Public p_Items        As Items

Public Property Let CurrentMove(ByVal n_CurrentMove      As Range)  : Set p_CurrentMove  = n_CurrentMove  : End Property
Public Property Let CurrentValue(ByVal n_CurrentValue    As Range)  : Set p_CurrentValue = n_CurrentValue : End Property
Public Property Let Fumons(ByVal n_Fumons                As Fumons) : Set p_Fumons       = n_Fumons       : End Property
Public Property Let Items(ByVal n_Items                  As Items)  : Set p_Items        = n_Items        : End Property

Public Property Get CurrentMove()                        As Range   : Set CurrentMove    = p_CurrentMove  : End Property
Public Property Get CurrentValue()                       As Range   : Set CurrentValue   = p_CurrentValue : End Property
Public Property Get Fumons()                             As Fumons  : Set Fumons         = p_Fumons       : End Property
Public Property Get Items()                              As Items   : Set Items          = p_Items        : End Property

Public Sub LetCurrentMove(ByVal n_CurrentMove As FightMove)
    CurrentMove.Value = n_CurrentMove
End Sub
Public Sub LetCurrentValue(ByVal n_CurrentValue As Variant)
    CurrentValue.Value = n_CurrentValue
End Sub

Public Function GetCurrentMove() As FightMove
    GetCurrentMove = CurrentMove.Value
End Function
Public Function GetCurrentValue(ByVal MyFight As Fight, ByVal MyPlayer As IFighter) As Variant
    Select Case GetCurrentMove
        Case FightMove.FightMoveAttack      : Set GetCurrentValue = MyFight.MyFumon(MyPlayer).Attacks.Attack(CurrentValue.Value)
        Case FightMove.FightMoveChangeFumon : Set GetCurrentValue = Fumons.Fumon(CurrentValue.Value)
        Case FightMove.FightMoveItem        : Set GetCurrentValue = Items.Item(CurrentValue.Value)
        Case FightMove.FightMoveFlee        : Set GetCurrentValue = Nothing
        Case FightMove.FightMoveNothing     : Set GetCurrentValue = Nothing
    End Select
End Function

Public Function Create(ByVal Rng As Range) As FightBase
    Set Create = New FightBase
    Dim TempFumons As Fumons: Set TempFumons = New Fumons
    Dim TempItems  As Items : Set TempItems  = New Items
    With Create
        .CurrentMove   = Rng.Offset(0, 0)
        .CurrentValue  = Rng.Offset(0, 1)
        .Fumons        = TempFumons.Create(Rng.Offset(0, 2))
        .Items         = TempItems.Create(Rng.Offset(0, 100))
    End With
End Function

Private Sub Class_Initialize()
    Set p_CurrentValue = Nothing
End Sub