VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IPlayer
Implements IFighter

Private p_PlayerBase        As PlayerBase
Private p_MoveBase          As MoveBase
Private p_FightBase         As FightBase
Private p_Script            As Script
Private p_ScriptValues      As String


Public Property Let PlayerBase(ByVal n_PlayerBase         As PlayerBase)         : Set p_PlayerBase      = n_PlayerBase        : End Property
Public Property Let MoveBase(ByVal n_MoveBase             As MoveBase)           : Set p_MoveBase        = n_MoveBase          : End Property
Public Property Let FightBase(ByVal n_FightBase           As FightBase)          : Set p_FightBase       = n_FightBase         : End Property
Public Property Let Script(ByVal n_Script                 As Script)             : Set p_Script          = n_Script            : End Property
Public Property Let ScriptValues(ByVal n_ScriptValues     As String)             : Let p_ScriptValues    = n_ScriptValues      : End Property

Public Property Get PlayerBase()                          As PlayerBase          : Set PlayerBase        = p_PlayerBase        : End Property
Public Property Get MoveBase()                            As MoveBase            : Set MoveBase          = p_MoveBase          : End Property
Public Property Get FightBase()                           As FightBase           : Set FightBase         = p_FightBase         : End Property
Public Property Get Script()                              As Script              : Set Script            = p_Script            : End Property
Public Property Get ScriptValues()                        As String              : Let ScriptValues      = p_ScriptValues      : End Property


'==========IFighter==========
Private Property Get IFighter_Number()  As Range  : Set IFighter_Number       = PlayerBase.Number : End Property
Private Property Get IFighter_Name()    As Range  : Set IFighter_Name         = PlayerBase.Name   : End Property
Private Property Get IFighter_Fumons()  As Fumons : Set IFighter_Fumons       = FightBase.Fumons  : End Property
Private Property Get IFighter_Items()   As Items  : Set IFighter_Items        = FightBase.Items   : End Property

Private Property Get IFighter_PlayerBase() As PlayerBase
    Set IFighter_PlayerBase = p_PlayerBase
End Property

Private Property Get IFighter_FightBase() As FightBase
    Set IFighter_FightBase = p_FightBase
End Property


Private Sub IFighter_DoAI(ByVal MyFight As Fight, ByVal OtherPlayer As IFighter)
    Call FightBase.LetCurrentValue(-1)
    Call FightBase.LetCurrentMove(FightMove.FightMoveNothing)
    Call ComAI(OtherPlayer)
End Sub


'==========IPlayer==========
Private Property Get IPlayer_Number()  As Range  : Set IPlayer_Number       = PlayerBase.Number : End Property
Private Property Get IPlayer_Name()    As Range  : Set IPlayer_Name         = PlayerBase.Name   : End Property

Private Property Get IPlayer_PlayerBase() As PlayerBase
    Set IPlayer_PlayerBase = p_PlayerBase
End Property

Private Property Get IPlayer_MoveBase() As MoveBase
    Set IPlayer_MoveBase = p_MoveBase
End Property

Private Sub IPlayer_Teleport(ByVal x As Long, ByVal y As Long)
    Call MoveBase.Teleport(x, y)
End Sub

Private Sub IPlayer_Move(ByVal x As Long, ByVal y As Long)
    Call MoveBase.Move(x, y)
End Sub

Private Sub IPlayer_MovePath(ByRef x() As Long, ByRef y() As Long)
    Call MoveBase.MovePath(x, y)
End Sub

Private Sub IPlayer_Look(ByVal Direction As XlDirection)
    Call MoveBase.Look(Direction)
End Sub

Private Sub IPlayer_Interact(ByVal Offset As Long)
    Call MoveBase.Interact(Me, Offset)
End Sub

Private Function IPlayer_SubTextureName() As String
    IPlayer_SubTextureName = MoveBase.SubTextureName(PlayerBase.Name.value)
End Function


'==========Public==========
Public Function Create(ByVal Pointer As Range, ByVal FumonDefinitions As PropCollection, ByVal ItemDefinitions As PropCollection, ByVal GameMaps As PropCollection, ByVal Scripts As PropCollection) As ComPlayer
    Set Create = New ComPlayer
    Dim TempFumons      As Fumons     : Set TempFumons     = New Fumons
    Dim TempItems       As Items      : Set TempItems      = New Items
    Dim TempPlayerBase  As PlayerBase : Set TempPlayerBase = New PlayerBase
    Dim TempMoveBase    As MoveBase   : Set TempMoveBase   = New MoveBase
    Dim TempFightBase   As FightBase  : Set TempFightBase  = New FightBase
    With Create
        .PlayerBase        = TempPlayerBase.Create(Pointer.Offset(0, 0))
        Call SpecificBase.Assign(Pointer.Offset(0, .PlayerBase.Offset), Create, Scripts)
        .MoveBase          = TempMoveBase.Create(.PlayerBase.Number.Value, Pointer.Offset(0, SpecificBase.Offset + .PlayerBase.Offset), GameMaps)
        .FightBase         = TempFightBase.Create(Pointer.Offset(0, SpecificBase.Offset + .PlayerBase.Offset + .MoveBase.Offset), FumonDefinitions, ItemDefinitions)
    End With
End Function

Private Sub ComAI(ByVal OtherPlayer As IPlayer)
    Call FightBase.LetCurrentMove(FightMove.FightMoveAttack)
    Call FightBase.LetCurrentValue(0)
End Sub