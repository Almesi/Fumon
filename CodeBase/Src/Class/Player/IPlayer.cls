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

Public Property Let Number(ByVal n_Number                 As Range)      : End Property
Public Property Let Name(ByVal n_Name                     As Range)      : End Property
Public Property Let Fumons(ByVal n_Fumons                 As Fumons)     : End Property
Public Property Let Items(ByVal n_Items                   As Items)      : End Property

Public Property Get Number()                              As Range       : End Property
Public Property Get Name()                                As Range       : End Property
Public Property Get Fumons()                              As Fumons      : End Property
Public Property Get Items()                               As Items       : End Property

Public Sub LetCurrentMove(ByVal n_CurrentMove As FightMove)
End Sub
Public Sub LetCurrentValue(ByVal n_CurrentValue As Variant)
End Sub

Public Function GetCurrentMove() As FightMove 
End Function
Public Function GetCurrentValue(ByVal MyFight As Fight) As Variant
End Function



Public Sub DoAI(ByVal MyFight As Fight, ByVal OtherPlayer As IPlayer)
End Sub