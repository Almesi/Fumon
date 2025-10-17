VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fumons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private Const p_TotalFumons As Long = 7
Private p_Fumons(p_TotalFumons) As Fumon
Private p_Selected As Long

Public Property Get TotalFumons() As Long
    TotalFumons = p_TotalFumons
End Property

Public Property Get FumonCount() As Long
    Dim i As Long
    For i = 0 To TotalFumons
        If IsNothing(Fumon(i)) Then
            FumonCount = i - 1
            Exit Function
        End If
    Next i
End Property

Public Property Get Fumons() As Fumon()
    Fumons = p_Fumons
End Property
Public Property Let Fumons(n_Fumons() As Fumon)
    p_Fumons = n_Fumons
End Property

Public Property Get Selected() As Long
    Selected = p_Selected
End Property
Public Property Let Selected(ByVal n_Selected As Long)
    p_Selected = n_Selected
End Property

Public Property Get Fumon(ByVal Index As Long) As Fumon
    Set Fumon = p_Fumons(Index)
End Property
Public Property Let Fumon(ByVal Index As Long, ByVal n_Fumon As Fumon)
    Set p_Fumons(Index) = n_Fumon
End Property

Public Function GetFumonIndex(ByVal Search As Fumon) As Long
    Dim i As Long
    For i = 0 To TotalFumons
        If Search Is Fumon(i) Then
            GetFumonIndex = i
            Exit Function
        End If
    Next i
    GetFumonIndex = -1
End Function

Public Function Create(ByVal Rng As Range) As Fumons
    Dim i As Long
    Set Create = New Fumons
    Dim Temp As Fumon: Set Temp = New Fumon
    With Create
        For i = 0 To TotalFumons
            If Rng.Offset(0, i * Temp.PropertyCount).Formula = "" Then Exit For
            .Fumon(i) = Temp.Create(Rng.Offset(0, i * Temp.PropertyCount))
            Call .Fumon(i).HealAndReset()
        Next i
    End With
End Function

Public Sub Swap(ByVal Index1 As Long, ByVal Index2 As Long)
    Dim Fumon1 As Fumon : Set Fumon1 = Fumon(Index1)
    Dim Fumon2 As Fumon : Set Fumon2 = Fumon(Index2)
    Call SwapRange(Fumon1, Fumon2, "Expirience")
    Call SwapRange(Fumon1, Fumon2, "CurrentHealth")
    Call SwapRange(Fumon1, Fumon2, "CurrentAttack")
    Call SwapRange(Fumon1, Fumon2, "CurrentDefense")
    Call SwapRange(Fumon1, Fumon2, "CurrentSpAttack")
    Call SwapRange(Fumon1, Fumon2, "CurrentSpDefense")
    Call SwapRange(Fumon1, Fumon2, "CurrentInitiative")
    Call SwapRange(Fumon1, Fumon2, "Attacks")
End Sub

Private Sub SwapRange(ByVal Fumon1 As Fumon, ByVal Fumon2 As Fumon, ByVal Prop As String)
    Dim Rng1 As Range, Rng2 As Range
    Select Case Prop
        Case "Expirience"        : Set Rng1 = Fumon1.Expirience        : Set Rng2 = Fumon2.Expirience
        Case "CurrentHealth"     : Set Rng1 = Fumon1.CurrentHealth     : Set Rng2 = Fumon2.CurrentHealth
        Case "CurrentAttack"     : Set Rng1 = Fumon1.CurrentAttack     : Set Rng2 = Fumon2.CurrentAttack
        Case "CurrentDefense"    : Set Rng1 = Fumon1.CurrentDefense    : Set Rng2 = Fumon2.CurrentDefense
        Case "CurrentSpAttack"   : Set Rng1 = Fumon1.CurrentSpAttack   : Set Rng2 = Fumon2.CurrentSpAttack
        Case "CurrentSpDefense"  : Set Rng1 = Fumon1.CurrentSpDefense  : Set Rng2 = Fumon2.CurrentSpDefense
        Case "CurrentInitiative" : Set Rng1 = Fumon1.CurrentInitiative : Set Rng2 = Fumon2.CurrentInitiative
        Case "Attacks"
            Call SwapAttacks(Fumon1, Fumon2)
            Exit Sub
    End Select
    Dim TempRange As Range
    Set TempRange = Rng1
    Set Rng1 = Rng2
    Set Rng2 = TempRange

    Dim TempValue As Variant
    TempValue = Rng1.Value
    Rng1.Value = Rng2.Value
    Rng2.Value = TempValue

    Select Case Prop
        Case "Expirience"        : Fumon1.Expirience        = Rng1 : Fumon2.Expirience        = Rng2
        Case "CurrentHealth"     : Fumon1.CurrentHealth     = Rng1 : Fumon2.CurrentHealth     = Rng2
        Case "CurrentAttack"     : Fumon1.CurrentAttack     = Rng1 : Fumon2.CurrentAttack     = Rng2
        Case "CurrentDefense"    : Fumon1.CurrentDefense    = Rng1 : Fumon2.CurrentDefense    = Rng2
        Case "CurrentSpAttack"   : Fumon1.CurrentSpAttack   = Rng1 : Fumon2.CurrentSpAttack   = Rng2
        Case "CurrentSpDefense"  : Fumon1.CurrentSpDefense  = Rng1 : Fumon2.CurrentSpDefense  = Rng2
        Case "CurrentInitiative" : Fumon1.CurrentInitiative = Rng1 : Fumon2.CurrentInitiative = Rng2
    End Select
End Sub

Private Sub SwapAttacks(ByVal Fumon1 As Fumon, ByVal Fumon2 As Fumon)
    Dim i As Long
    Dim Rng1 As Range
    Dim Rng2 As Range
    Dim TempRange As Range
    Dim TempValue As Variant
    For i = 0 To Fumon1.Attacks.AttackCount
        Set Rng1 = Fumon1.Attacks.AttackRange(i)
        Set Rng2 = Fumon2.Attacks.AttackRange(i)

        Set TempRange = Rng1
        Set Rng1 = Rng2
        Set Rng2 = TempRange

        TempValue = Rng1.Value
        Rng1.Value = Rng2.Value
        Rng2.Value = TempValue
        Fumon1.Attacks.AttackRange(i) = Rng1
        Fumon2.Attacks.AttackRange(i) = Rng2
    Next i
End Sub

Public Function First() As Fumon
    Dim i As Long
    For i = 0 To TotalFumons
        If IsSomething(Fumon(i)) Then
            Set First = Fumon(i)
            Exit Function
        End If
    Next i
End Function

Public Function FirstAlive() As Fumon
    Dim i As Long
    For i = 0 To TotalFumons
        If IsSomething(Fumon(i)) Then
            If Fumon(i).Alive Then
                Set FirstAlive = Fumon(i)
                Exit Function
            End If
        End If
    Next i
End Function