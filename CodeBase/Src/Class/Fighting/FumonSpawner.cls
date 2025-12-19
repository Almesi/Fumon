VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FumonSpawner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private Const ItemOffset As Long = 4

'Sum of all Chances should add up to 100
Private Type FumonSpawn
    Definition  As FumonDefinition
    LowerLevel  As Long
    UpperLevel  As Long
    Chance      As Long
End Type

Private p_Number        As Long
Private p_Name          As String
Private p_TriggerChance As Long

Private Spawns() As FumonSpawn
Private Size As Long


Public Property Let Number(n_Value        As Long)   : Let p_Number        = n_Value         : End Property
Public Property Let Name(n_Value          As String) : Let p_Name          = n_Value         : End Property
Public Property Let TriggerChance(n_Value As Long)   : Let p_TriggerChance = n_Value         : End Property

Public Property Get Number()              As Long    : Let Number          = p_Number        : End Property
Public Property Get Name()                As String  : Let Name            = p_Name          : End Property
Public Property Get TriggerChance()       As Long    : Let TriggerChance   = p_TriggerChance : End Property


Public Function Create(ByVal Rng As IRange, ByVal FumonDefinitions As PropCollection) As FumonSpawner
    Set Create = New FumonSpawner
    Dim i As Long
    With Create
        .Number         = Rng.Offset(0, 0).Value
        .Name           = Rng.Offset(0, 1).Value
        .TriggerChance  = Rng.Offset(0, 2).Value
        
        Dim ItemRng As IRange
        Set ItemRng = Rng.Offset(0, 3 + i)
        Do Until ItemRng.Formula = Empty
            Set ItemRng = Rng.Offset(0, 3 + i)
            Call Create.AddSpawn(FumonDefinitions.Object(ItemRng.Value), _
                                             ItemRng.Offset(0, 1).Value, _
                                             ItemRng.Offset(0, 2).Value, _
                                             ItemRng.Offset(0, 3).Value)
            i = i + ItemOffset
        Loop
    End With
End Function

Public Sub AddSpawn(ByVal Definition As FumonDefinition, ByVal Lowerlevel As Long, ByVal Upperlevel As Long, ByVal Chance As Long)
    Size = Size + 1
    ReDim Preserve Spawns(Size)
    With Spawns(Size)
        Set .Definition = Definition
        Let .LowerLevel = LowerLevel
        Let .UpperLevel = UpperLevel
        Let .Chance     = Chance
    End With
End Sub

Public Function Spawn(ByVal Rng As IRange) As Fumon
    If RandomNumber(0, 100) > TriggerChance Then Exit Function

    Dim SpawnChance As Long
    SpawnChance = RandomNumber(0, 100)

    Dim Index As Long
    Dim Lower As Long
    Dim Upper As Long

    Dim i As Long
    For i = 0 To Size
        Upper = Spawns(i).Chance
        If Lower > SpawnChance And Upper =< SpawnChance Then
            Index = i
            Exit For
        End If
        Lower = Spawns(i).Chance
    Next i

    Dim Level      As Long            : Let Level      = RandomNumber(Spawns(Index).LowerLevel, Spawns(Index).UpperLevel)
    Dim Definition As FumonDefinition : Set Definition = Spawns(Index).Definition

    Set Spawn = CreateFumon(Rng, Definition, Level)
End Function

Private Function CreateFumon(ByVal Rng As IRange, ByVal Definition As FumonDefinition, ByVal Level As Long) As Fumon
    Set CreateFumon = New Fumon
    With CreateFumon
        .Definition        = Definition
        .Number            = Rng.Offset(0, 0)
        .Number.Value      = Definition.Number
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
        .Attacks           = Temp.CreateFromObject(Rng.Offset(0, 9), Definition, Level)
        .HealAndReset
    End With
End Function

Private Sub Class_Initialize()
    Size = -1
End Sub