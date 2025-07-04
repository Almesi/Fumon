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

Public Property Let Number(n_Number                       As Range)              : End Property
Public Property Let Name(n_Name                           As Range)              : End Property
Public Property Let Row(n_Row                             As Range)              : End Property
Public Property Let Column(n_Column                       As Range)              : End Property
Public Property Let LookDirection(n_LookDirection         As Range)              : End Property
Public Property Let Money(n_Money                         As Range)              : End Property
Public Property Let Fumons(n_Fumons                       As Fumons)             : End Property
Public Property Let Items(n_Items                         As Items)              : End Property
Public Property Let CurrentFight(n_CurrentFight           As CurrentFight)       : End Property
Public Property Let TextureDefinition(n_TextureDefinition As TextureDefinition)  : End Property

Public Property Get Number()                              As Range               : End Property
Public Property Get Name()                                As Range               : End Property
Public Property Get Row()                                 As Range               : End Property
Public Property Get Column()                              As Range               : End Property
Public Property Get LookDirection()                       As Range               : End Property
Public Property Get Money()                               As Range               : End Property
Public Property Get Fumons()                              As Fumons              : End Property
Public Property Get Items()                               As Items               : End Property
Public Property Get CurrentFight()                        As CurrentFight        : End Property
Public Property Get TextureDefinition()                   As TextureDefinition   : End Property

Public Function Create(Pointer As Range) As IPlayer
End Function

Public Sub Move(x As Long, y As Long)
End Sub

Public Sub MovePath(x() As Long, y() As Long, LookDirection() As Long)
End Sub

Public Sub Look(Direction As Long)
End Sub

Private Sub Interact(TileRange As Range)
End Sub

Public Sub DoAI(OtherPlayer As IPlayer)
End Sub

Public Sub ChangeFumon(Index As Long)
End Sub

Public Sub Attack(UsedAttack As Attack, P2Fumon As Fumon)
End Sub

Public Sub UseItem(UsedItem As Item)
End Sub

Public Function Flee(OtherPlayer As IPlayer) As IPlayer
End Function

Public Sub AddBeaten(Rng As Range)
End Sub

Public Sub AddQuest(NewQuest As Quest)
End Sub

Public Function PlayerDefeated(Index As Long) As Boolean
End Function