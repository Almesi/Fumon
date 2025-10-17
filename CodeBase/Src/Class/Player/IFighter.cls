VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Property Get IFighter_Number()  As Range  : End Property
Private Property Get IFighter_Name()    As Range  : End Property
Private Property Get IFighter_Fumons()  As Fumons : End Property
Private Property Get IFighter_Items()   As Items  : End Property

Public Property Get PlayerBase() As PlayerBase
End Property

Public Property Get FightBase() As FightBase
End Property

Public Sub DoAI(ByVal MyFight As Fight, ByVal OtherPlayer As IFighter)
End Sub