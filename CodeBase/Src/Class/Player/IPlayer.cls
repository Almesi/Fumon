VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_PlayerBase   As PlayerBase
Private p_SpecificBase As SpecificBase
Private p_MoveBase     As MoveBase
Private p_Box          As FumonBox
Private p_FightBase    As FightBase


Public Property Let PlayerBase(ByVal n_PlayerBase     As PlayerBase)     : Set p_PlayerBase   = n_PlayerBase   : End Property
Public Property Let SpecificBase(ByVal n_SpecificBase As SpecificBase)   : Set p_SpecificBase = n_SpecificBase : End Property
Public Property Let MoveBase(ByVal n_MoveBase         As MoveBase)       : Set p_MoveBase     = n_MoveBase     : End Property
Public Property Let Box(ByVal n_Box                   As FumonBox)       : Set p_Box          = n_Box          : End Property
Public Property Let FightBase(ByVal n_FightBase       As FightBase)      : Set p_FightBase    = n_FightBase    : End Property

Public Property Get PlayerBase()                      As PlayerBase      : Set PlayerBase     = p_PlayerBase   : End Property
Public Property Get SpecificBase()                    As SpecificBase    : Set SpecificBase   = p_SpecificBase : End Property
Public Property Get MoveBase()                        As MoveBase        : Set MoveBase       = p_MoveBase     : End Property
Public Property Get Box()                             As FumonBox        : Set Box            = p_Box          : End Property
Public Property Get FightBase()                       As FightBase       : Set FightBase      = p_FightBase    : End Property

Public Function Create(ByVal Rng As IRange, ByVal FumonDefinitions As PropCollection, ByVal ItemDefinitions As PropCollection, ByVal GameMaps As PropCollection, ByVal Scripts As PropCollection) As IPlayer
    Set Create = New IPlayer
    Dim TempPlayerBase   As PlayerBase   : Set TempPlayerBase   = New PlayerBase
    Dim TempSpecificBase As SpecificBase : Set TempSpecificBase = New SpecificBase
    Dim TempMoveBase     As MoveBase     : Set TempMoveBase     = New MoveBase
    Dim TempFightBase    As FightBase    : Set TempFightBase    = New FightBase

    Dim SpecificOffset As Long : SpecificOffset = TempPlayerBase.Offset()
    Dim MoveOffset     As Long : MoveOffset     = SpecificOffset + TempSpecificBase.Offset()
    Dim FightOffset    As Long : FightOffset    = MoveOffset + TempMoveBase.Offset()

    With Create
        .PlayerBase        = TempPlayerBase.Create(Rng)
        .SpecificBase      = TempSpecificBase.Create(Rng.Offset(0, SpecificOffset), Scripts)
        .MoveBase          = TempMoveBase.Create(Rng.Offset(0, MoveOffset), .PlayerBase, GameMaps)
        .Box               = FumonBox.Create(Rng.Offset(0, FightOffset))
        .FightBase         = TempFightBase.Create(Rng.Offset(0, FightOffset + 1), FumonDefinitions, ItemDefinitions)
    End With
End Function