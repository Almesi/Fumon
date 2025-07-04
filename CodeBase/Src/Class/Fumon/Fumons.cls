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

Private Const p_FumonCount As Long = 7
Private p_Fumons(p_FumonCount) As Fumon

Public Property Get FumonCount() As Long
    FumonCount = p_FumonCount
End Property

Public Property Get Fumons() As Fumon()
    Fumons = p_Fumons
End Property
Public Property Let Fumons(n_Fumons() As Fumon)
    p_Fumons = n_Fumons
End Property

Public Property Get Fumon(Index As Long) As Fumon
    Set Fumon = p_Fumons(Index)
End Property
Public Property Let Fumon(Index As Long, n_Fumon As Fumon)
    Set p_Fumons(Index) = n_Fumon
End Property

Public Function Create(Rng As Range) As Fumons
    Dim i As Long
    Set Create = New Fumons
    With Create
        For i = 0 To FumonCount
            If Rng.Offset(0, i * FumonGame.Fumon.PropertyCount).Formula = "" Then Exit For
            .Fumon(i) = FumonGame.Fumon.Create(Rng.Offset(0, i * FumonGame.Fumon.PropertyCount))
        Next i
    End With
End Function

Public Sub Swap(Index1 As Long, Index2 As Long)
    Dim Temp As Variant
    Temp                                     = p_Fumons(Index1).Expirience.Value
    p_Fumons(Index1).Expirience.Value        = p_Fumons(Index2).Expirience.Value
    p_Fumons(Index2).Expirience.Value        = Temp
    Temp                                     = p_Fumons(Index1).CurrentHealth.Value
    p_Fumons(Index1).CurrentHealth.Value     = p_Fumons(Index2).CurrentHealth.Value
    p_Fumons(Index2).CurrentHealth.Value     = Temp
    Temp                                     = p_Fumons(Index1).CurrentAttack.Value
    p_Fumons(Index1).CurrentAttack.Value     = p_Fumons(Index2).CurrentAttack.Value
    p_Fumons(Index2).CurrentAttack.Value     = Temp
    Temp                                     = p_Fumons(Index1).CurrentDefense.Value
    p_Fumons(Index1).CurrentDefense.Value    = p_Fumons(Index2).CurrentDefense.Value
    p_Fumons(Index2).CurrentDefense.Value    = Temp
    Temp                                     = p_Fumons(Index1).CurrentSpAttack.Value
    p_Fumons(Index1).CurrentSpAttack.Value   = p_Fumons(Index2).CurrentSpAttack.Value
    p_Fumons(Index2).CurrentSpAttack.Value   = Temp
    Temp                                     = p_Fumons(Index1).CurrentSpDefense.Value
    p_Fumons(Index1).CurrentSpDefense.Value  = p_Fumons(Index2).CurrentSpDefense.Value
    p_Fumons(Index2).CurrentSpDefense.Value  = Temp
    Temp                                     = p_Fumons(Index1).CurrentInitiative.Value
    p_Fumons(Index1).CurrentInitiative.Value = p_Fumons(Index2).CurrentInitiative.Value
    p_Fumons(Index2).CurrentInitiative.Value = Temp

    Dim TempRange As Range
    Set TempRange                          = p_Fumons(Index1).Expirience
    Set p_Fumons(Index1).Expirience        = p_Fumons(Index2).Expirience
    Set p_Fumons(Index2).Expirience        = TempRange
    Set TempRange                          = p_Fumons(Index1).CurrentHealth
    Set p_Fumons(Index1).CurrentHealth     = p_Fumons(Index2).CurrentHealth
    Set p_Fumons(Index2).CurrentHealth     = TempRange
    Set TempRange                          = p_Fumons(Index1).CurrentAttack
    Set p_Fumons(Index1).CurrentAttack     = p_Fumons(Index2).CurrentAttack
    Set p_Fumons(Index2).CurrentAttack     = TempRange
    Set TempRange                          = p_Fumons(Index1).CurrentDefense
    Set p_Fumons(Index1).CurrentDefense    = p_Fumons(Index2).CurrentDefense
    Set p_Fumons(Index2).CurrentDefense    = TempRange
    Set TempRange                          = p_Fumons(Index1).CurrentSpAttack
    Set p_Fumons(Index1).CurrentSpAttack   = p_Fumons(Index2).CurrentSpAttack
    Set p_Fumons(Index2).CurrentSpAttack   = TempRange
    Set TempRange                          = p_Fumons(Index1).CurrentSpDefense
    Set p_Fumons(Index1).CurrentSpDefense  = p_Fumons(Index2).CurrentSpDefense
    Set p_Fumons(Index2).CurrentSpDefense  = TempRange
    Set TempRange                          = p_Fumons(Index1).CurrentInitiative
    Set p_Fumons(Index1).CurrentInitiative = p_Fumons(Index2).CurrentInitiative
    Set p_Fumons(Index2).CurrentInitiative = TempRange
End Sub

Public Function FirstFumon() As Fumon
    Dim i As Long
    For i = 0 To FumonCount
        If Fumon(i).CurrentHealth > 0 Then
            Set FirstFumon = Fumon(i)
            Exit Function
        End If
    Next i
End Function

Private Function CheckFumonHealth() As Boolean
    Dim i As Long
    For i = 0 To Fumons.FumonCount
        If Fumon(i).CurrentHealth > 0 Then
            CheckFumonHealth = True
            Exit Function
        End If
    Next i
End Function