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
Private p_CurrentMove       As Range
Private p_CurrentValue      As Range
Private p_Fumons            As Fumons
Private p_Items             As Items
Private p_SubTexture        As VBGLSubTexture


Public Property Let Number(ByVal n_Number                 As Range)              : Set p_Number          = n_Number            : End Property
Public Property Let Name(ByVal n_Name                     As Range)              : Set p_Name            = n_Name              : End Property
Public Property Let CurrentMove(ByVal n_CurrentMove       As Range)              : Set p_CurrentMove     = n_CurrentMove       : End Property
Public Property Let CurrentValue(ByVal n_CurrentValue     As Range)              : Set p_CurrentValue    = n_CurrentValue      : End Property
Public Property Let Fumons(ByVal n_Fumons                 As Fumons)             : Set p_Fumons          = n_Fumons            : End Property
Public Property Let Items(ByVal n_Items                   As Items)              : Set p_Items           = n_Items             : End Property
Public Property Let SubTexture(ByVal n_SubTexture         As VBGLSubTexture)     : Set p_SubTexture      = n_SubTexture        : End Property

Public Property Get Number()                              As Range               : Set Number            = p_Number            : End Property
Public Property Get Name()                                As Range               : Set Name              = p_Name              : End Property
Public Property Get CurrentMove()                         As Range               : Set CurrentMove       = p_CurrentMove       : End Property
Public Property Get CurrentValue()                        As Range               : Set CurrentValue      = p_CurrentValue      : End Property
Public Property Get Fumons()                              As Fumons              : Set Fumons            = p_Fumons            : End Property
Public Property Get Items()                               As Items               : Set Items             = p_Items             : End Property
Public Property Get SubTexture()                          As VBGLSubTexture      : Set SubTexture        = p_SubTexture        : End Property

Private Property Let IPlayer_Number(ByVal n_Number        As Range)              : Set p_Number          = n_Number            : End Property
Private Property Let IPlayer_Name(ByVal n_Name            As Range)              : Set p_Name            = n_Name              : End Property
Private Property Let IPlayer_Fumons(ByVal n_Fumons        As Fumons)             : Set p_Fumons          = n_Fumons            : End Property
Private Property Let IPlayer_Items(ByVal n_Items          As Items)              : Set p_Items           = n_Items             : End Property

Private Property Get IPlayer_Number()                     As Range               : Set IPlayer_Number    = p_Number            : End Property
Private Property Get IPlayer_Name()                       As Range               : Set IPlayer_Name      = p_Name              : End Property
Private Property Get IPlayer_Fumons()                     As Fumons              : Set IPlayer_Fumons    = p_Fumons            : End Property
Private Property Get IPlayer_Items()                      As Items               : Set IPlayer_Items     = p_Items             : End Property

Private Sub IPlayer_LetCurrentMove(ByVal n_CurrentMove As FightMove)
    CurrentMove.Value = n_CurrentMove
End Sub
Private Sub IPlayer_LetCurrentValue(ByVal n_CurrentValue As Variant)
    CurrentValue.Value = n_CurrentValue
End Sub

Private Function IPlayer_GetCurrentMove() As FightMove
    IPlayer_GetCurrentMove = CurrentMove.Value
End Function
Private Function IPlayer_GetCurrentValue(ByVal MyFight As Fight) As Variant
    Select Case IPlayer_GetCurrentMove
        Case FightMove.FightMoveAttack      : Set IPlayer_GetCurrentValue = MyFight.MyFumon(Me).Attacks.Attack(CurrentValue.Value)
        Case FightMove.FightMoveChangeFumon : Set IPlayer_GetCurrentValue = IPlayer_Fumons.Fumon(CurrentValue.Value)
        Case FightMove.FightMoveItem        : Set IPlayer_GetCurrentValue = IPlayer_Items.Item(CurrentValue.Value)
        Case FightMove.FightMoveFlee        : Set IPlayer_GetCurrentValue = Nothing
        Case FightMove.FightMoveNothing     : Set IPlayer_GetCurrentValue = Nothing
    End Select
End Function

Private Sub IPlayer_DoAI(ByVal MyFight As Fight, ByVal OtherPlayer As IPlayer)
    Call IPlayer_LetCurrentValue(-1)
    Call IPlayer_LetCurrentMove(FightMove.FightMoveNothing)
    Call ComAI(OtherPlayer)
End Sub

Public Function Create(ByVal Pointer As Range) As WildPlayer
    Set Create = New WildPlayer
    With Create
        .Number            = Pointer.Offset(0, 0)
        .Name              = Pointer.Offset(0, 1)
        .Money             = Pointer.Offset(0, 2)
        .CurrentMove       = Pointer.Offset(0, 10)
        .CurrentValue      = Pointer.Offset(0, 11)
        .Fumons            = FactoryFumons.Create(Pointer.Offset(0, 12))
        .Items             = FactoryItems.Create(Pointer.Offset(0, 110))
        .SubTexture        = FactoryServer.GetTexture("Players").SubTextureID(.SubTextureName())
    End With
End Function

Private Sub ComAI(ByVal OtherPlayer As IPlayer)
    Call IPlayer_LetCurrentMove(FightMove.FightMoveAttack)
    Call IPlayer_LetCurrentValue(0)
End Sub

Private Sub Class_Initialize()
    Set p_CurrentValue = Nothing
End Sub