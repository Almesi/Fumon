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
Private p_TextureDefinition  As TextureDefinition

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
Public Property Let TextureDefinition(n_TextureDefinition As TextureDefinition)  : Set p_TextureDefinition = n_TextureDefinition : End Property

Public Property Get PropertyCount()                       As Long                : Let PropertyCount       = p_PropertyCount     : End Property

Public Property Get Number()                              As Long                : Let Number              = p_Number            : End Property
Public Property Get Name()                                As String              : Let Name                = p_Name              : End Property
Public Property Get Func()                                As String              : Let Func                = p_Func              : End Property
Public Property Get Typee()                               As AttackType          : Let Typee               = p_Typee             : End Property
Public Property Get ElementType()                         As String              : Let ElementTypee        = p_ElementTypee      : End Property
Public Property Get TextureDefinition()                   As TextureDefinition   : Set TextureDefinition   = p_TextureDefinition : End Property


Public Function Create(Rng As Range) As Attack
    Set Create = New Attack
    With Create
        .Number            = Rng.Offset(0, 0).Value
        .Name              = Rng.Offset(0, 1).Value
        .Func              = Rng.Offset(0, 2).Value
        .Typee             = GetType(Rng.Offset(0, 3).Value)
        .ElementType       = Rng.Offset(0, 4).Value
        .TextureDefinition = FumonGame.TextureDefinition.Create(Rng.Offset(0, 5).Value, Rng.Offset(0, 6).Value, FumonGame.FactoryServer.GameMap.Folder.Value & "\Attack.png")
    End With
End Function

Private Function GetType(Text As String) As AttackType
    Select Case Text
        Case "Health"          : GetType = AttackType.Health
        Case "Attack"          : GetType = AttackType.Attack
        Case "Defense"         : GetType = AttackType.Defense
        Case "SpAttack"        : GetType = AttackType.SpAttack
        Case "SpDefense"       : GetType = AttackType.SpDefense
        Case "Initiative"      : GetType = AttackType.Initiative
    End Select
End Function

Public Sub Use(Fumon1 As Fumon, Fumon2 As Fumon)
    Dim Temp As stdLambda
    Dim Factor As Double

    Set Temp = stdLambda.Create(Func)
    Factor = GameServer.ElementType.GetFactor(ElementType, Fumon2)
    If Func Like "$1*=*" Then
        Select Case Typee
            Case AttackType.Health     : Fumon1.Health.Value      = Temp.Run(Fumon1, Fumon2) * Factor
            Case AttackType.Attack     : Fumon1.Attack.Value      = Temp.Run(Fumon1, Fumon2) * Factor
            Case AttackType.Defense    : Fumon1.Defense.Value     = Temp.Run(Fumon1, Fumon2) * Factor
            Case AttackType.SpAttack   : Fumon1.SpAttack.Value    = Temp.Run(Fumon1, Fumon2) * Factor
            Case AttackType.SpDefense  : Fumon1.SpDefense.Value   = Temp.Run(Fumon1, Fumon2) * Factor
            Case AttackType.Initiative : Fumon1.Initiative.Value  = Temp.Run(Fumon1, Fumon2) * Factor
        End Select
    Else
        Select Case Typee
            Case AttackType.Health     : Fumon2.Health.Value      = Temp.Run(Fumon1, Fumon2) * Factor
            Case AttackType.Attack     : Fumon2.Attack.Value      = Temp.Run(Fumon1, Fumon2) * Factor
            Case AttackType.Defense    : Fumon2.Defense.Value     = Temp.Run(Fumon1, Fumon2) * Factor
            Case AttackType.SpAttack   : Fumon2.SpAttack.Value    = Temp.Run(Fumon1, Fumon2) * Factor
            Case AttackType.SpDefense  : Fumon2.SpDefense.Value   = Temp.Run(Fumon1, Fumon2) * Factor
            Case AttackType.Initiative : Fumon2.Initiative.Value  = Temp.Run(Fumon1, Fumon2) * Factor
        End Select
    End If
    Call NegativeValue(Fumon1)
    Call NegativeValue(Fumon2)
End Sub

Public Sub NegativeValue(Mon As Fumon)
    If Mon.Health.Value     < 0 Then Mon.Health.Value     = 0
    If Mon.Attack.Value     < 0 Then Mon.Attack.Value     = 0
    If Mon.Defense.Value    < 0 Then Mon.Defense.Value    = 0
    If Mon.SpAttack.Value   < 0 Then Mon.SpAttack.Value   = 0
    If Mon.SpDefense.Value  < 0 Then Mon.SpDefense.Value  = 0
    If Mon.Initiative.Value < 0 Then Mon.Initiative.Value = 0
End Sub