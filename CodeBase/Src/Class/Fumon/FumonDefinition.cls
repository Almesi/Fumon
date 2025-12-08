VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FumonDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

' Define the Fumon by using the Stats on Level 100. The modifiers will divide the Max Stat x-times to reach a certain level

Private p_Number             As Long
Private p_Name               As String
Private p_EvolveTo           As String
Private p_EvolveCondition    As String
Private p_Type1              As String
Private p_Type2              As String
Private p_Health             As Long
Private p_Attack             As Long
Private p_Defense            As Long
Private p_SpAttack           As Long
Private p_SpDefense          As Long
Private p_Initiative         As Long
Private p_HealthModifier     As Double
Private p_AttackModifier     As Double
Private p_DefenseModifier    As Double
Private p_SpAttackModifier   As Double
Private p_SpDefenseModifier  As Double
Private p_InitiativeModifier As Double
Private p_LevelFormula       As String
Private p_Attacks()          As Attack
Private p_LearnAttacks()     As Long

Public Property Let Number(n_Number                            As Long)              : Let p_Number                = n_Number             : End Property
Public Property Let Name(n_Name                                As String)            : Let p_Name                  = n_Name               : End Property
Public Property Let EvolveTo(n_EvolveTo                        As String)            : Let p_EvolveTo              = n_EvolveTo           : End Property
Public Property Let EvolveCondition(n_EvolveCondition          As String)            : Let p_EvolveCondition       = n_EvolveCondition    : End Property
Public Property Let Type1(n_Type1                              As String)            : Let p_Type1                 = n_Type1              : End Property
Public Property Let Type2(n_Type2                              As String)            : Let p_Type2                 = n_Type2              : End Property
Public Property Let Health(n_Health                            As Long)              : Let p_Health                = n_Health             : End Property
Public Property Let Attack(n_Attack                            As Long)              : Let p_Attack                = n_Attack             : End Property
Public Property Let Defense(n_Defense                          As Long)              : Let p_Defense               = n_Defense            : End Property
Public Property Let SpAttack(n_SpAttack                        As Long)              : Let p_SpAttack              = n_SpAttack           : End Property
Public Property Let SpDefense(n_SpDefense                      As Long)              : Let p_SpDefense             = n_SpDefense          : End Property
Public Property Let Initiative(n_Initiative                    As Long)              : Let p_Initiative            = n_Initiative         : End Property
Public Property Let HealthModifier(n_HealthModifier            As Double)            : Let p_HealthModifier        = n_HealthModifier     : End Property
Public Property Let AttackModifier(n_AttackModifier            As Double)            : Let p_AttackModifier        = n_AttackModifier     : End Property
Public Property Let DefenseModifier(n_DefenseModifier          As Double)            : Let p_DefenseModifier       = n_DefenseModifier    : End Property
Public Property Let SpAttackModifier(n_SpAttackModifier        As Double)            : Let p_SpAttackModifier      = n_SpAttackModifier   : End Property
Public Property Let SpDefenseModifier(n_SpDefenseModifier      As Double)            : Let p_SpDefenseModifier     = n_SpDefenseModifier  : End Property
Public Property Let InitiativeModifier(n_InitiativeModifier    As Double)            : Let p_InitiativeModifier    = n_InitiativeModifier : End Property
Public Property Let LevelFormula(n_LevelFormula                As String)            : Let p_LevelFormula          = n_LevelFormula       : End Property
Public Property Let Attacks(n_Attacks                          As Variant)           : Let p_Attacks               = n_Attacks            : End Property
Public Property Let LearnAttacks(n_LearnAttacks                As Variant)           : Let p_LearnAttacks          = n_LearnAttacks       : End Property

Public Property Let AttackDef(Index As Long, n_Attack          As Attack)            : Let p_Attacks(Index)        = n_Attack             : End Property
Public Property Let LearnAttack(Index As Long, n_LearnAttack   As Long)              : Let p_LearnAttacks(Index)   = n_LearnAttack        : End Property

Public Property Get Number()                                   As Long               : Let Number                  = p_Number             : End Property
Public Property Get Name()                                     As String             : Let Name                    = p_Name               : End Property
Public Property Get EvolveTo()                                 As String             : Let EvolveTo                = p_EvolveTo           : End Property
Public Property Get EvolveCondition()                          As String             : Let EvolveCondition         = p_EvolveCondition    : End Property
Public Property Get Type1()                                    As String             : Let Type1                   = p_Type1              : End Property
Public Property Get Type2()                                    As String             : Let Type2                   = p_Type2              : End Property
Public Property Get Health()                                   As Long               : Let Health                  = p_Health             : End Property
Public Property Get Attack()                                   As Long               : Let Attack                  = p_Attack             : End Property
Public Property Get Defense()                                  As Long               : Let Defense                 = p_Defense            : End Property
Public Property Get SpAttack()                                 As Long               : Let SpAttack                = p_SpAttack           : End Property
Public Property Get SpDefense()                                As Long               : Let SpDefense               = p_SpDefense          : End Property
Public Property Get Initiative()                               As Long               : Let Initiative              = p_Initiative         : End Property
Public Property Get HealthModifier()                           As Double             : Let HealthModifier          = p_HealthModifier     : End Property
Public Property Get AttackModifier()                           As Double             : Let AttackModifier          = p_AttackModifier     : End Property
Public Property Get DefenseModifier()                          As Double             : Let DefenseModifier         = p_DefenseModifier    : End Property
Public Property Get SpAttackModifier()                         As Double             : Let SpAttackModifier        = p_SpAttackModifier   : End Property
Public Property Get SpDefenseModifier()                        As Double             : Let SpDefenseModifier       = p_SpDefenseModifier  : End Property
Public Property Get InitiativeModifier()                       As Double             : Let InitiativeModifier      = p_InitiativeModifier : End Property
Public Property Get LevelFormula()                             As String             : Let LevelFormula            = p_LevelFormula       : End Property
Public Property Get Attacks()                                  As Variant            : Let Attacks                 = p_Attacks            : End Property
Public Property Get LearnAttacks()                             As Variant            : Let LearnAttacks            = p_LearnAttacks       : End Property

Public Property Get AttackDef(Index As Long)                   As Attack             : Set AttackDef               = p_Attacks(Index)     : End Property
Public Property Get LearnAttack(Index As Long)                 As Long               : Let LearnAttack             = p_LearnAttacks(Index): End Property


Public Function Create(ByVal Rng As IRange, ByVal AttackDefs As PropCollection) As FumonDefinition
    Set Create = New FumonDefinition
    With Create
        .Number             = Rng.Offset(0, 00).Value
        .Name               = Rng.Offset(0, 01).Value
        .EvolveTo           = Rng.Offset(0, 02).Value
        .EvolveCondition    = Rng.Offset(0, 03).Value
        .Type1              = Rng.Offset(0, 04).Value
        .Type2              = Rng.Offset(0, 05).Value
        .Health             = Rng.Offset(0, 06).Value
        .Attack             = Rng.Offset(0, 07).Value
        .Defense            = Rng.Offset(0, 08).Value
        .SpAttack           = Rng.Offset(0, 09).Value
        .SpDefense          = Rng.Offset(0, 10).Value
        .Initiative         = Rng.Offset(0, 11).Value
        .HealthModifier     = Rng.Offset(0, 12).Value
        .AttackModifier     = Rng.Offset(0, 13).Value
        .DefenseModifier    = Rng.Offset(0, 14).Value
        .SpAttackModifier   = Rng.Offset(0, 15).Value
        .SpDefenseModifier  = Rng.Offset(0, 16).Value
        .InitiativeModifier = Rng.Offset(0, 17).Value
        .LevelFormula       = Rng.Offset(0, 18).Value
        Call .GetAttacks(Rng.Offset(0, 19).Value, AttackDefs)
    End With
End Function

Public Sub GetAttacks(ByVal Definition As String, ByVal AttackDefs As PropCollection)
    Dim i As Long
    Dim AttackDefinitions() As String
    AttackDefinitions = Split(Definition, ", ")
    For i = 0 To USize(AttackDefinitions)
        Dim Text           As String : Text           = AttackDefinitions(i)
        Dim Point          As Long   : Point          = InStr(1, Text, "(")
        Dim AttackIndex    As Long   : AttackIndex    = Clng(VBGLMidP(Text, 1, Point - 1))
        Dim LearnAttackDef As Long   : LearnAttackDef = CLng(VBGLMidP(Text, Point + 1, Len(Text) - 1))

        Call VBGLAdd(p_Attacks, AttackDefs.Object(AttackIndex))
        Call VBGLAdd(p_LearnAttacks, LearnAttackDef)
    Next i
End Sub