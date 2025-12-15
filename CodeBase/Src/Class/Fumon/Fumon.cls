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



Private Const p_PropertyCount As Long = 13
Private Const MaximumLevel    As Long = 100
Private Const ExpFactor       As Single = 0.01

Private p_Definition         As FumonDefinition
Private p_Number             As IRange
Private p_Level              As IRange
Private p_Expirience         As IRange
Private p_CurrentHealth      As IRange
Private p_CurrentAttack      As IRange
Private p_CurrentDefense     As IRange
Private p_CurrentSpAttack    As IRange
Private p_CurrentSpDefense   As IRange
Private p_CurrentInitiative  As IRange
Private p_Attacks            As Attacks



Public Property Let Definition(n_Definition        As FumonDefinition)  : Set p_Definition        = n_Definition         : End Property
Public Property Let Number(n_Number                As IRange)           : Set p_Number            = n_Number             : End Property
Public Property Let Level(n_Level                  As IRange)           : Set p_Level             = n_Level              : End Property
Public Property Let Expirience(n_Expirience        As IRange)           : Set p_Expirience        = n_Expirience         : End Property
Public Property Let CurrentHealth(n_Health         As IRange)           : Set p_CurrentHealth     = n_Health             : End Property
Public Property Let CurrentAttack(n_Attack         As IRange)           : Set p_CurrentAttack     = n_Attack             : End Property
Public Property Let CurrentDefense(n_Defense       As IRange)           : Set p_CurrentDefense    = n_Defense            : End Property
Public Property Let CurrentSpAttack(n_SpAttack     As IRange)           : Set p_CurrentSpAttack   = n_SpAttack           : End Property
Public Property Let CurrentSpDefense(n_SpDefense   As IRange)           : Set p_CurrentSpDefense  = n_SpDefense          : End Property
Public Property Let CurrentInitiative(n_Initiative As IRange)           : Set p_CurrentInitiative = n_Initiative         : End Property
Public Property Let Attacks(n_Attacks              As Attacks)          : Set p_Attacks           = n_Attacks            : End Property

Public Property Get PropertyCount()                As Long              : Let PropertyCount       = p_PropertyCount      : End Property

Public Property Get Definition()                   As FumonDefinition   : Set Definition          = p_Definition         : End Property
Public Property Get Number()                       As IRange            : Set Number              = p_Number             : End Property
Public Property Get Level()                        As IRange            : Set Level               = p_Level              : End Property
Public Property Get Expirience()                   As IRange            : Set Expirience          = p_Expirience         : End Property
Public Property Get CurrentHealth()                As IRange            : Set CurrentHealth       = p_CurrentHealth      : End Property
Public Property Get CurrentAttack()                As IRange            : Set CurrentAttack       = p_CurrentAttack      : End Property
Public Property Get CurrentDefense()               As IRange            : Set CurrentDefense      = p_CurrentDefense     : End Property
Public Property Get CurrentSpAttack()              As IRange            : Set CurrentSpAttack     = p_CurrentSpAttack    : End Property
Public Property Get CurrentSpDefense()             As IRange            : Set CurrentSpDefense    = p_CurrentSpDefense   : End Property
Public Property Get CurrentInitiative()            As IRange            : Set CurrentInitiative   = p_CurrentInitiative  : End Property
Public Property Get Attacks()                      As Attacks           : Set Attacks             = p_Attacks            : End Property


Public Function CreateFromObject(ByVal Rng As IRange, ByVal Definition As FumonDefinition, ByVal Level As Long) As Fumon
    Set CreateFromObject = New Fumon
    With CreateFromObject
        .Definition        = Definition
        .Number            = Rng.Offset(0, 0)
        .Level             = Rng.Offset(0, 1)
        .Level.Value       = Level
        .Expirience        = Rng.Offset(0, 2)
        .CurrentHealth     = Rng.Offset(0, 3)
        .CurrentAttack     = Rng.Offset(0, 4)
        .CurrentDefense    = Rng.Offset(0, 5)
        .CurrentSpAttack   = Rng.Offset(0, 6)
        .CurrentSpDefense  = Rng.Offset(0, 7)
        .CurrentInitiative = Rng.Offset(0, 8)
        Dim Temp As Attacks
        Set Temp = New Attacks
        .Attacks           = Temp.CreateFromObject(Rng.Offset(0, 8), Definition, Level)
        .HealAndReset
    End With
End Function

Public Function Create(ByVal Rng As IRange, ByVal Definitions As PropCollection) As Fumon
    Dim Definition As FumonDefinition : Set Definition = Definitions.Object(Rng.Offset(0, 0).Value)
    Dim Level      As Long            : Level      = Rng.Offset(0, 1).Value
    Set Create = CreateFromObject(Rng, Definition, Level)
End Function

Public Sub FinishFight(ByVal BeatenFumon As Fumon)
    Expirience.Value = Expirience.Value + (BeatenFumon.Expirience.Value * ExpFactor)
    Call Reset()
End Sub

Public Function FullName() As String
    FullName = Definition.Name & _
               " Lv: "  & Level.Value & _
               " Hea: " & CurrentHealth.Value & _
               " Att: " & CurrentAttack.Value & _
               " Def: " & CurrentDefense.Value & _
               " SpA: " & CurrentSpAttack.Value & _
               " SpD: " & CurrentSpDefense.Value & _
               " Ini: " & CurrentInitiative.Value
End Function

Public Function FrontName() As String
    FrontName = Definition.Name & "Front"
End Function

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

Public Sub UpdateLevel()
    Dim Lambda As stdLambda
    Set Lambda = stdLambda.Create(Definition.LevelFormula)
    Dim i As Long

    Do While Expirience.Value >= NeededExp
        Dim NeededExp As Long
        Let NeededExp = Lambda.Run(Level.Value)
        Level.Value = Level.Value + 1
        Expirience.Value = Expirience.Value - NeededExp
    Loop
End Sub

Public Function MaxHealth()     As Long: MaxHealth     = CLng(Definition.Health     / (Definition.HealthModifier     ^ (MaximumLevel - Level.Value))) : End Function
Public Function MaxAttack()     As Long: MaxAttack     = CLng(Definition.Attack     / (Definition.AttackModifier     ^ (MaximumLevel - Level.Value))) : End Function
Public Function MaxDefense()    As Long: MaxDefense    = CLng(Definition.Defense    / (Definition.DefenseModifier    ^ (MaximumLevel - Level.Value))) : End Function
Public Function MaxSpAttack()   As Long: MaxSpAttack   = CLng(Definition.SpAttack   / (Definition.SpAttackModifier   ^ (MaximumLevel - Level.Value))) : End Function
Public Function MaxSpDefense()  As Long: MaxSpDefense  = CLng(Definition.SpDefense  / (Definition.SpDefenseModifier  ^ (MaximumLevel - Level.Value))) : End Function
Public Function MaxInitiative() As Long: MaxInitiative = CLng(Definition.Initiative / (Definition.InitiativeModifier ^ (MaximumLevel - Level.Value))) : End Function

Public Sub CheckHealth()
    Dim Level As Long
    Level = GetLevel(Expirience.Value)
    If CurrentHealth.Value > CLng(Definition.Health / (Definition.HealthModifier ^ (MaximumLevel - Level))) Then
        CurrentHealth.Value = CLng(Definition.Health / (Definition.HealthModifier ^ (MaximumLevel - Level)))
    End If
End Sub

Public Function Alive() As Boolean
    Alive = CurrentHealth.Value > 0
End Function