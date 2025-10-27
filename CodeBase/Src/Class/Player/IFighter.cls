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

Public Property Get Number()  As Range  : End Property
Public Property Get Name()    As Range  : End Property
Public Property Get Fumons()  As Fumons : End Property
Public Property Get Items()   As Items  : End Property

Public Property Get PlayerBase() As PlayerBase
End Property

Public Property Get FightBase() As FightBase
End Property

Public Sub DoAI(ByVal MyFight As Fight, ByVal OtherPlayer As IFighter)
End Sub