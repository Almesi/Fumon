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

Public Property Get Number()  As Range
End Property
Public Property Get Name()    As Range
End Property

Public Property Get PlayerBase() As PlayerBase
End Property

Public Property Get MoveBase() As MoveBase
End Property

Public Sub Teleport(ByVal x As Long, ByVal y As Long)
End Sub

Public Sub Move(ByVal x As Long, ByVal y As Long)
End Sub

Public Sub MovePath(ByRef x() As Long, ByRef y() As Long)
End Sub

Public Sub Look(ByVal Direction As XlDirection)
End Sub

Public Sub Interact(ByVal Offset As Long)
End Sub

Public Function SubTextureName() As String
End Function