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

Private Const p_Count As Long = 3
Private p_Attacks(p_Count) As IRange
Private p_Selected As Long

Public Property Get Count() As Long
    Count = p_Count
End Property

Public Property Get Attacks() As Variant
    Attacks = p_Attacks
End Property
Public Property Let Attacks(ByVal n_Attacks As Variant)
    p_Attacks = n_Attacks
End Property

Public Property Get Selected() As Long
    Selected = p_Selected
End Property
Public Property Let Selected(ByVal n_Selected As Long)
    p_Selected = n_Selected
End Property

Public Property Get AttackRange(ByVal Index As Long) As IRange
    Set AttackRange = p_Attacks(Index)
End Property
Public Property Let AttackRange(ByVal Index As Long, ByVal n_AttackRange As IRange)
    Set p_Attacks(Index) = n_AttackRange
End Property

Public Property Get Attack(ByVal Index As Long) As Attack
    Set Attack = MeServer.Attack(AttackRange(Index).Value)
End Property
Public Property Let Attack(ByVal Index As Long, ByVal n_Attack As Attack)
    MeServer.Attack(AttackRange(Index)) = n_Attack
End Property

Public Property Get SelectedAttack() As Attack
    Set SelectedAttack = Attack(Selected)
End Property

Public Function CreateFromObject(ByVal Rng As IRange, ByVal Definition As FumonDefinition, ByVal Level As Long) As Attacks
    Dim MyAttacks() As Attack
    MyAttacks = GetAttacksFromDefinition(Definition, Level)
    Dim i As Long
    For i = 0 To Count
        If i > USize(MyAttacks) Then Exit For
        Rng.Offset(0, i).Value = MyAttacks(i).Number
    Next i
    Set CreateFromObject = Create(Rng)
End Function

Public Function Create(ByVal Rng As IRange) As Attacks
    Dim i As Long
    Set Create = New Attacks
    Dim CurrentRng As Long
    With Create
        For i = 0 To Count
            .AttackRange(i) = Rng.Offset(0, i)
        Next i
    End With
End Function


Public Sub Swap(ByVal Index1 As Long, ByVal Index2 As Long)
    Dim Temp As Attack
    Set Temp               = p_Attacks(Index1)
    Set p_Attacks(Index1)  = p_Attacks(Index2)
    Set p_Attacks(Index2)  = Temp
End Sub

Private Function GetAttacksFromDefinition(ByVal Definition As FumonDefinition, ByVal Level As Long) As Attack()
    Dim i As Long
    Dim Arr() As Attack
    Dim Offset As Long

    For i = 0 To USize(Definition.Attacks)
        If Definition.LearnAttack(i) < Level Then
            If USize(Arr) = Count Then
                Set Arr(Offset) = Definition.AttackDef(i)
                Offset = Offset + 1
                If Offset > Count Then Offset = 0
            Else
                Call VBGLAdd(Arr, Definition.AttackDef(i))
            End If
        End If
    Next i
    GetAttacksFromDefinition = Arr
End Function