VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fumon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private Const p_PropertyCount As Long = 12
Private Const MaximumLevel    As Long = 100

Private p_Definition         As FumonDefinition
Private p_Expirience         As Range
Private p_CurrentHealth      As Range
Private p_CurrentAttack      As Range
Private p_CurrentDefense     As Range
Private p_CurrentSpAttack    As Range
Private p_CurrentSpDefense   As Range
Private p_CurrentInitiative  As Range
Private p_Attacks            As Attacks



Public Property Let Definition(n_Definition        As FumonDefinition) : Set p_Definition        = n_Definition         : End Property
Public Property Let Expirience(n_Expirience        As Range)           : Set p_Expirience        = n_Expirience         : End Property
Public Property Let CurrentHealth(n_Health         As Range)           : Set p_CurrentHealth     = n_Health             : End Property
Public Property Let CurrentAttack(n_Attack         As Range)           : Set p_CurrentAttack     = n_Attack             : End Property
Public Property Let CurrentDefense(n_Defense       As Range)           : Set p_CurrentDefense    = n_Defense            : End Property
Public Property Let CurrentSpAttack(n_SpAttack     As Range)           : Set p_CurrentSpAttack   = n_SpAttack           : End Property
Public Property Let CurrentSpDefense(n_SpDefense   As Range)           : Set p_CurrentSpDefense  = n_SpDefense          : End Property
Public Property Let CurrentInitiative(n_Initiative As Range)           : Set p_CurrentInitiative = n_Initiative         : End Property
Public Property Let Attacks(n_Attacks              As Attacks)         : Set p_Attacks           = n_Attacks            : End Property

Public Property Get PropertyCount()                As Long             : Let PropertyCount       = p_PropertyCount      : End Property

Public Property Get Definition()                   As FumonDefinition  : Set Definition          = p_Definition         : End Property
Public Property Get Expirience()                   As Range            : Set Expirience          = p_Expirience         : End Property
Public Property Get CurrentHealth()                As Range            : Set CurrentHealth       = p_CurrentHealth      : End Property
Public Property Get CurrentAttack()                As Range            : Set CurrentAttack       = p_CurrentAttack      : End Property
Public Property Get CurrentDefense()               As Range            : Set CurrentDefense      = p_CurrentDefense     : End Property
Public Property Get CurrentSpAttack()              As Range            : Set CurrentSpAttack     = p_CurrentSpAttack    : End Property
Public Property Get CurrentSpDefense()             As Range            : Set CurrentSpDefense    = p_CurrentSpDefense   : End Property
Public Property Get CurrentInitiative()            As Range            : Set CurrentInitiative   = p_CurrentInitiative  : End Property
Public Property Get Attacks()                      As Attacks          : Set Attacks             = p_Attacks            : End Property


Public Function Create(ByVal Rng As Range, ByVal Definitions As PropCollection) As Fumon
    Set Create = New Fumon
    With Create
        .Definition        = Definitions.Object(Rng.Offset(0, 0).Value)
        .Expirience        = Rng.Offset(0, 1)
        .CurrentHealth     = Rng.Offset(0, 2)
        .CurrentAttack     = Rng.Offset(0, 3)
        .CurrentDefense    = Rng.Offset(0, 4)
        .CurrentSpAttack   = Rng.Offset(0, 5)
        .CurrentSpDefense  = Rng.Offset(0, 6)
        .CurrentInitiative = Rng.Offset(0, 7)
        Dim Temp As Attacks
        Set Temp = New Attacks
        .Attacks           = Temp.Create(Rng.Offset(0, 8))
        .HealAndReset
    End With
End Function

Public Sub Printt(Rng As Range)
    Rng.Offset(0, 0).Value = Definition.Number
    Rng.Offset(0, 1).Value = Expirience.Value
    Rng.Offset(0, 2).Value = CurrentHealth.Value
    Rng.Offset(0, 3).Value = CurrentAttack.Value
    Rng.Offset(0, 4).Value = CurrentDefense.Value
    Rng.Offset(0, 5).Value = CurrentSpAttack.Value
    Rng.Offset(0, 6).Value = CurrentSpDefense.Value
    Rng.Offset(0, 7).Value = CurrentInitiative.Value
End Sub

Public Sub FinishFight(GainEXP As Long)
    Expirience.Value = Expirience.Value + GainEXP
    Call Reset()
End Sub

Public Sub HealAndReset()
    Call Heal()
    Call Reset()
End Sub
Public Sub Heal()
    CurrentHealth.Value = MaxHealth()
End Sub
Public Sub Reset()
    CurrentAttack.Value     = MaxAttack()
    CurrentDefense.Value    = MaxDefense()
    CurrentSpAttack.Value   = MaxSpAttack()
    CurrentSpDefense.Value  = MaxSpDefense()
    CurrentInitiative.Value = MaxInitiative()
End Sub

Public Function GetLevel() As Long
    Dim Lambda As stdLambda
    Set Lambda = stdLambda.Create(Definition.LevelFormula)
    GetLevel = Lambda.Run(Expirience.Value)
End Function

Public Function MaxHealth()     As Long: MaxHealth     = CLng(Definition.Health     / (Definition.HealthModifier     ^ (MaximumLevel - GetLevel))) : End Function
Public Function MaxAttack()     As Long: MaxAttack     = CLng(Definition.Attack     / (Definition.AttackModifier     ^ (MaximumLevel - GetLevel))) : End Function
Public Function MaxDefense()    As Long: MaxDefense    = CLng(Definition.Defense    / (Definition.DefenseModifier    ^ (MaximumLevel - GetLevel))) : End Function
Public Function MaxSpAttack()   As Long: MaxSpAttack   = CLng(Definition.SpAttack   / (Definition.SpAttackModifier   ^ (MaximumLevel - GetLevel))) : End Function
Public Function MaxSpDefense()  As Long: MaxSpDefense  = CLng(Definition.SpDefense  / (Definition.SpDefenseModifier  ^ (MaximumLevel - GetLevel))) : End Function
Public Function MaxInitiative() As Long: MaxInitiative = CLng(Definition.Initiative / (Definition.InitiativeModifier ^ (MaximumLevel - GetLevel))) : End Function

Public Sub CheckHealth()
    Dim Level As Long
    Level = GetLevel(Expirience.Value)
    If CurrentHealth.Value > CLng(Definition.Health / (Definition.HealthModifier ^ (MaximumLevel - Level))) Then
        CurrentHealth.Value = CLng(Definition.Health / (Definition.HealthModifier ^ (MaximumLevel - Level)))
    End If
End Sub

Public Function Alive() As Boolean
    Alive = CurrentHealth > 0
End Function