VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'True
END
Attribute VB_Name = "StatusEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private Enum Effect
    NoEffect  = 0
    Burned    = 1
    Paralyzed = 2
    Frozen    = 3
    Sleeping  = 4
    Poisoned  = 5
End Enum

Public Function ClearEffect(ByVal Fumon As Fumon, Optional WhatEffect As Long = 0) As Boolean
    Dim CurrentEffect As Long
    CurrentEffect = Fumon.CurrentEffect.Value

    If WhatEffect <> 0 Then
        If WhatEffect = CurrentEffect Then
            Fumon.CurrentEffect.Value = 0
        Else
            Exit Function
        End If
    Else
        Fumon.CurrentEffect.Value = 0
    End If
    ClearEffect = True
End Function

Public Sub DoEffect(ByVal Fumon As Fumon, ByVal MoveDec As MoveDecision)
    Select Case Fumon.CurrentEffect.Value
        Case Effect.Burned, Effect.Poisoned
            Dim Damage As Long
            Damage = Fumon.MaxHealth / 10
            Fumon.CurrentHealth.Value = Fumon.CurrentHealth.Value - Damage
        Case Effect.Paralyzed, Effect.Frozen, Effect.Sleeping
            Dim Rand As Long
            Rand = RandomNumber()
            If Rand < 10 Then
                MoveDec = Nothing ' Skip turn
            End If
    End Select
End Sub