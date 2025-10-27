VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Attacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private Const p_AttackCount As Long = 3
Private p_Attacks(p_AttackCount) As Range
Private p_Selected As Long

Public Property Get AttackCount() As Long
    AttackCount = p_AttackCount
End Property

Public Property Get Attacks() As Attack()
    Attacks = p_Attacks
End Property
Public Property Let Attacks(n_Attacks() As Attack)
    p_Attacks = n_Attacks
End Property

Public Property Get Selected() As Long
    Selected = p_Selected
End Property
Public Property Let Selected(n_Selected As Long)
    p_Selected = n_Selected
End Property

Public Property Get AttackRange(Index As Long) As Range
    Set AttackRange = p_Attacks(Index)
End Property
Public Property Let AttackRange(Index As Long, n_AttackRange As Range)
    Set p_Attacks(Index) = n_AttackRange
End Property

Public Property Get Attack(Index As Long) As Attack
    Set Attack = MeServer.Attack(AttackRange(Index).Value)
End Property
Public Property Let Attack(Index As Long, n_Attack As Attack)
    MeServer.Attack(AttackRange(Index)) = n_Attack
End Property

Public Property Get SelectedAttack() As Attack
    Set SelectedAttack = Attack(Selected)
End Property


Public Function Create(Rng As Range) As Attacks
    Dim i As Long
    Set Create = New Attacks
    Dim CurrentRng As Long
    With Create
        For i = 0 To AttackCount
            .AttackRange(i) = Rng.Offset(0, i)
        Next i
    End With
End Function

Public Sub Swap(Index1 As Long, Index2 As Long)
    Dim Temp As Attack
    Set Temp               = p_Attacks(Index1)
    Set p_Attacks(Index1)  = p_Attacks(Index2)
    Set p_Attacks(Index2)  = Temp
End Sub