VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HumanPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IPlayer
Implements IFighter

Private p_PlayerBase As PlayerBase
Private p_MoveBase   As MoveBase
Private p_FightBase  As FightBase
Private p_Beaten     As PropCollection
Private p_Quests     As PropCollection


Public Property Let PlayerBase(ByVal n_PlayerBase As PlayerBase)     : Set p_PlayerBase = n_PlayerBase : End Property
Public Property Let MoveBase(ByVal n_MoveBase     As MoveBase)       : Set p_MoveBase   = n_MoveBase   : End Property
Public Property Let FightBase(ByVal n_FightBase   As FightBase)      : Set p_FightBase  = n_FightBase  : End Property
Public Property Let Beaten(ByVal n_Beaten         As PropCollection) : Set p_Beaten     = n_Beaten     : End Property
Public Property Let Quests(ByVal n_Quests         As PropCollection) : Set p_Quests     = n_Quests     : End Property

Public Property Get PlayerBase()                  As PlayerBase      : Set PlayerBase   = p_PlayerBase : End Property
Public Property Get MoveBase()                    As MoveBase        : Set MoveBase     = p_MoveBase   : End Property
Public Property Get FightBase()                   As FightBase       : Set FightBase    = p_FightBase  : End Property
Public Property Get Beaten()                      As PropCollection  : Set Beaten       = p_Beaten     : End Property
Public Property Get Quests()                      As PropCollection  : Set Quests       = p_Quests     : End Property

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
    Dim StartTime As Single
    StartTime = Timer
    Do Until FightBase.GetCurrentMove <> FightMove.FightMoveNothing
        Dim Diff As Single
        Diff = Timer - StartTime
        If Diff > 60 Then
            Exit Do
        End If
        If MyFight.ImFinished(Me) Then
            Exit Do
        End If
        Call glutMainLoopEvent()
        Call MyFight.UpdateTimer(Me, Diff / 60)
        Call CurrentRenderObject.Loopp
    Loop
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
Public Function Create(ByVal Pointer As Range, ByVal FumonDefinitions As PropCollection, ByVal ItemDefinitions As PropCollection, ByVal GameMaps As PropCollection, ByVal Scripts As PropCollection) As HumanPlayer
    Set Create = New HumanPlayer
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


Public Sub AddBeaten(ByVal Player As ComPlayer)
    Call Beaten.AddUnique(Player)
End Sub

Public Sub AddQuest(ByVal nQuest As Quest)
    Call Quests.AddUnique(nQuest)
End Sub