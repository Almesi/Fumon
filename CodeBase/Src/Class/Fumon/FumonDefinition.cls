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
Private p_TextureDefinition  As TextureDefinition

Public Property Let Number(n_Number                          As Long)              : p_Number                = n_Number             : End Property
Public Property Let Name(n_Name                              As String)            : p_Name                  = n_Name               : End Property
Public Property Let EvolveTo(n_EvolveTo                      As String)            : p_EvolveTo              = n_EvolveTo           : End Property
Public Property Let EvolveCondition(n_EvolveCondition        As String)            : p_EvolveCondition       = n_EvolveCondition    : End Property
Public Property Let Type1(n_Type1                            As String)            : p_Type1                 = n_Type1              : End Property
Public Property Let Type2(n_Type2                            As String)            : p_Type2                 = n_Type2              : End Property
Public Property Let Health(n_Health                          As Long)              : p_Health                = n_Health             : End Property
Public Property Let Attack(n_Attack                          As Long)              : p_Attack                = n_Attack             : End Property
Public Property Let Defense(n_Defense                        As Long)              : p_Defense               = n_Defense            : End Property
Public Property Let SpAttack(n_SpAttack                      As Long)              : p_SpAttack              = n_SpAttack           : End Property
Public Property Let SpDefense(n_SpDefense                    As Long)              : p_SpDefense             = n_SpDefense          : End Property
Public Property Let Initiative(n_Initiative                  As Long)              : p_Initiative            = n_Initiative         : End Property
Public Property Let HealthModifier(n_HealthModifier          As Double)            : p_HealthModifier        = n_HealthModifier     : End Property
Public Property Let AttackModifier(n_AttackModifier          As Double)            : p_AttackModifier        = n_AttackModifier     : End Property
Public Property Let DefenseModifier(n_DefenseModifier        As Double)            : p_DefenseModifier       = n_DefenseModifier    : End Property
Public Property Let SpAttackModifier(n_SpAttackModifier      As Double)            : p_SpAttackModifier      = n_SpAttackModifier   : End Property
Public Property Let SpDefenseModifier(n_SpDefenseModifier    As Double)            : p_SpDefenseModifier     = n_SpDefenseModifier  : End Property
Public Property Let InitiativeModifier(n_InitiativeModifier  As Double)            : p_InitiativeModifier    = n_InitiativeModifier : End Property
Public Property Let LevelFormula(n_LevelFormula              As String)            : p_LevelFormula          = n_LevelFormula       : End Property
Public Property Let TextureDefinition(n_TextureDefinition    As TextureDefinition) : Set p_TextureDefinition = n_TextureDefinition  : End Property

Public Property Get Number()                                 As Long               : Number                  = p_Number             : End Property
Public Property Get Name()                                   As String             : Name                    = p_Name               : End Property
Public Property Get EvolveTo()                               As String             : EvolveTo                = p_EvolveTo           : End Property
Public Property Get EvolveCondition()                        As String             : EvolveCondition         = p_EvolveCondition    : End Property
Public Property Get Type1()                                  As String             : Type1                   = p_Type1              : End Property
Public Property Get Type2()                                  As String             : Type2                   = p_Type2              : End Property
Public Property Get Health()                                 As Long               : Health                  = p_Health             : End Property
Public Property Get Attack()                                 As Long               : Attack                  = p_Attack             : End Property
Public Property Get Defense()                                As Long               : Defense                 = p_Defense            : End Property
Public Property Get SpAttack()                               As Long               : SpAttack                = p_SpAttack           : End Property
Public Property Get SpDefense()                              As Long               : SpDefense               = p_SpDefense          : End Property
Public Property Get Initiative()                             As Long               : Initiative              = p_Initiative         : End Property
Public Property Get HealthModifier()                         As Double             : HealthModifier          = p_HealthModifier     : End Property
Public Property Get AttackModifier()                         As Double             : AttackModifier          = p_AttackModifier     : End Property
Public Property Get DefenseModifier()                        As Double             : DefenseModifier         = p_DefenseModifier    : End Property
Public Property Get SpAttackModifier()                       As Double             : SpAttackModifier        = p_SpAttackModifier   : End Property
Public Property Get SpDefenseModifier()                      As Double             : SpDefenseModifier       = p_SpDefenseModifier  : End Property
Public Property Get InitiativeModifier()                     As Double             : InitiativeModifier      = p_InitiativeModifier : End Property
Public Property Get LevelFormula()                           As String             : LevelFormula            = p_LevelFormula       : End Property
Public Property Get TextureDefinition()                      As TextureDefinition  : Set TextureDefinition   = p_TextureDefinition  : End Property


Public Function Create(Rng As Range) As FumonDefinition
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
        .TextureDefinition  = FumonGame.TextureDefinition.Create(Rng.Offset(0, 19).Value, Rng.Offset(0, 20).Value, FumonGame.FactoryServer.GameMap.Folder.Value & "\Fumon.png")
    End With
End Function