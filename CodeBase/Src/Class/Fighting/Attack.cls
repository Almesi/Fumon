VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Attack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit



Private Const p_PropertyCount As Long = 5

Private p_Number             As Long
Private p_Name               As String
Private p_Func               As String
Private p_Typee              As AttackType
Private p_ElementTypee       As String

Public Enum AttackType
    Health         = 0
    Attack         = 1
    Defense        = 2
    SpAttack       = 3
    SpDefense      = 4
    Initiative     = 5
End Enum

Public Property Let Number(n_Number                       As Long)               : Let p_Number            = n_Number            : End Property
Public Property Let Name(n_Name                           As String)             : Let p_Name              = n_Name              : End Property
Public Property Let Func(n_Func                           As String)             : Let p_Func              = n_Func              : End Property
Public Property Let Typee(n_Typee                         As AttackType)         : Let p_Typee             = n_Typee             : End Property
Public Property Let ElementType(n_ElementTypee            As String)             : Let p_ElementTypee      = n_ElementTypee      : End Property

Public Property Get PropertyCount()                       As Long                : Let PropertyCount       = p_PropertyCount     : End Property

Public Property Get Number()                              As Long                : Let Number              = p_Number            : End Property
Public Property Get Name()                                As String              : Let Name                = p_Name              : End Property
Public Property Get Func()                                As String              : Let Func                = p_Func              : End Property
Public Property Get Typee()                               As AttackType          : Let Typee               = p_Typee             : End Property
Public Property Get ElementType()                         As String              : Let ElementType         = p_ElementTypee      : End Property




Public Function Create(ByVal Rng As IRange) As Attack
    Set Create = New Attack
    With Create
        .Number            = Rng.Offset(0, 0).Value
        .Name              = Rng.Offset(0, 1).Value
        .Func              = Rng.Offset(0, 2).Value
        .Typee             = GetType(Rng.Offset(0, 3).Value)
        .ElementType       = Rng.Offset(0, 4).Value
    End With
End Function

Private Function GetType(ByVal Text As String) As AttackType
    Select Case UCase(Text)
        Case "HEALTH"          : GetType = AttackType.Health
        Case "ATTACK"          : GetType = AttackType.Attack
        Case "DEFENSE"         : GetType = AttackType.Defense
        Case "SPATTACK"        : GetType = AttackType.SpAttack
        Case "SPDEFENSE"       : GetType = AttackType.SpDefense
        Case "INITIATIVE"      : GetType = AttackType.Initiative
    End Select
End Function

Public Function GetTypeName() As String
    Select Case Typee
        Case AttackType.Health          : GetTypeName = "Health"
        Case AttackType.Attack          : GetTypeName = "Attack"
        Case AttackType.Defense         : GetTypeName = "Defense"
        Case AttackType.SpAttack        : GetTypeName = "SpAttack"
        Case AttackType.SpDefense       : GetTypeName = "SpDefense"
        Case AttackType.Initiative      : GetTypeName = "Initiative"
    End Select
End Function

Public Function FullName() As String
    FullName = Name & " " & GetTypeName() & ": ( " & ElementType & " | " & Func & " )"
End Function

Public Sub Use(ByVal Fumon1 As Fumon, ByVal Fumon2 As Fumon)
    Dim Temp As stdLambda
    Dim Factor As Double
    Dim Destination As Fumon

    Set Temp = stdLambda.Create(Func)
    Factor = MeServer.ElementTypes.GetFactor(ElementType, Fumon2)
    If Func Like "$1*=*" Then
        Set Destination = Fumon1
    Else
        Set Destination = Fumon2
    End If

    Select Case Typee
        Case AttackType.Health     : Destination.CurrentHealth.Value      = Destination.CurrentHealth.Value - (Temp.Run(Fumon1) * Factor)
        Case AttackType.Attack     : Destination.CurrentAttack.Value      = Temp.Run(Fumon1) * Factor
        Case AttackType.Defense    : Destination.CurrentDefense.Value     = Temp.Run(Fumon1) * Factor
        Case AttackType.SpAttack   : Destination.CurrentSpAttack.Value    = Temp.Run(Fumon1) * Factor
        Case AttackType.SpDefense  : Destination.CurrentSpDefense.Value   = Temp.Run(Fumon1) * Factor
        Case AttackType.Initiative : Destination.CurrentInitiative.Value  = Temp.Run(Fumon1) * Factor
    End Select
    Call NegativeValue(Fumon1)
    Call NegativeValue(Fumon2)
End Sub

Public Sub NegativeValue(ByVal Mon As Fumon)
    If Mon.CurrentHealth.Value     < 0 Then Mon.CurrentHealth.Value     = 0
    If Mon.CurrentAttack.Value     < 0 Then Mon.CurrentAttack.Value     = 0
    If Mon.CurrentDefense.Value    < 0 Then Mon.CurrentDefense.Value    = 0
    If Mon.CurrentSpAttack.Value   < 0 Then Mon.CurrentSpAttack.Value   = 0
    If Mon.CurrentSpDefense.Value  < 0 Then Mon.CurrentSpDefense.Value  = 0
    If Mon.CurrentInitiative.Value < 0 Then Mon.CurrentInitiative.Value = 0
End Sub