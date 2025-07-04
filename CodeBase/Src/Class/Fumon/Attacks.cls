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
Private p_Attacks(p_AttackCount) As Attack

Public Property Get AttackCount() As Long
    AttackCount = p_AttackCount
End Property

Public Property Get Attacks() As Attack()
    Attacks = p_Attacks
End Property
Public Property Let Attacks(n_Attacks() As Attack)
    p_Attacks = n_Attacks
End Property

Public Property Get Attack(Index As Long) As Attack
    Set Attack = p_Attacks(Index)
End Property
Public Property Let Attack(Index As Long, n_Attack As Attack)
    Set p_Attacks(Index) = n_Attack
End Property

Public Function Create(Rng As Range) As Attacks
    Dim i As Long
    Set Create = New Attacks
    With Create
        For i = 0 To AttackCount
            .Attack(i) = FumonGame.FactoryServer.Attack(Rng.Offset(0, i).Value)
        Next i
    End With
End Function

Public Sub Swap(Index1 As Long, Index2 As Long)
    Dim Temp As Attack
    Set Temp               = p_Attacks(Index1)
    Set p_Attacks(Index1)  = p_Attacks(Index2)
    Set p_Attacks(Index2)  = Temp
End Sub