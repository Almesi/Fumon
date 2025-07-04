VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WildPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IPlayer

Private p_Number            As Range
Private p_Name              As Range
Private p_Row               As Range
Private p_Column            As Range
Private p_LookDirection     As Range
Private p_Money             As Range
Private p_Fumons            As Fumons
Private p_Items             As Items
Private p_TextureDefinition As TextureDefinition

Private Property Let IPlayer_Number(n_Number                       As Range)              : Set p_Number                  = n_Number            : End Property
Private Property Let IPlayer_Name(n_Name                           As Range)              : Set p_Name                    = n_Name              : End Property
Private Property Let IPlayer_Row(n_Row                             As Range)              : Set p_Row                     = n_Row               : End Property
Private Property Let IPlayer_Column(n_Column                       As Range)              : Set p_Column                  = n_Column            : End Property
Private Property Let IPlayer_LookDirection(n_LookDirection         As Range)              : Set p_LookDirection           = n_LookDirection     : End Property
Private Property Let IPlayer_Money(n_Money                         As Range)              : Set p_Money                   = n_Money             : End Property
Private Property Let IPlayer_Fumons(n_Fumons                       As Fumons)             : Set p_Fumons                  = n_Fumons            : End Property
Private Property Let IPlayer_Items(n_Items                         As Items)              : Set p_Items                   = n_Items             : End Property
Private Property Let IPlayer_TextureDefinition(n_TextureDefinition As TextureDefinition)  : Set p_TextureDefinition       = n_TextureDefinition : End Property

Private Property Get IPlayer_Number()                              As Range               : Set IPlayer_Number            = p_Number            : End Property
Private Property Get IPlayer_Name()                                As Range               : Set IPlayer_Name              = p_Name              : End Property
Private Property Get IPlayer_Row()                                 As Range               : Set IPlayer_Row               = p_Row               : End Property
Private Property Get IPlayer_Column()                              As Range               : Set IPlayer_Column            = p_Column            : End Property
Private Property Get IPlayer_LookDirection()                       As Range               : Set IPlayer_LookDirection     = p_LookDirection     : End Property
Private Property Get IPlayer_Money()                               As Range               : Set IPlayer_Money             = p_Money             : End Property
Private Property Get IPlayer_Fumons()                              As Fumons              : Set IPlayer_Fumons            = p_Fumons            : End Property
Private Property Get IPlayer_Items()                               As Items               : Set IPlayer_Items             = p_Items             : End Property
Private Property Get IPlayer_TextureDefinition()                   As TextureDefinition   : Set IPlayer_TextureDefinition = p_TextureDefinition : End Property

Private Function IPlayer_Create(Pointer As Range) As IPlayer
End Function

Private Sub IPlayer_Move(x As Long, y As Long)
End Sub

Private Sub IPlayer_MovePath(x() As Long, y() As Long, LookDirection() As Long)
End Sub

Private Sub IPlayer_Look(Direction As Long)
End Sub

Private Sub IPlayer_Interact(TileRange As Range)
End Sub

Private Sub IPlayer_DoAI(OtherPlayer As IPlayer)
End Sub

Private Sub IPlayer_ChangeFumon(Index As Long)
End Sub

Private Sub IPlayer_Attack(UsedAttack As Attack, P2Fumon As Fumon)
End Sub

Private Sub IPlayer_UseItem(UsedItem As Item)
End Sub

Private Function IPlayer_Flee(OtherPlayer As IPlayer) As IPlayer
End Function

Private Sub IPlayer_AddBeaten(Rng As Range)
End Sub

Private Sub IPlayer_AddQuest(NewQuest As Quest)
End Sub


Private Function IPlayer_PlayerDefeated(Index As Long) As Boolean
End Function