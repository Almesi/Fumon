VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecificBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit


Public Function Offset() As Long
    Offset = 5
End Function

Public Sub Assign(ByVal Rng As Range, ByVal Origin As Object, ByVal Scripts As PropCollection)
    If TypeName(Origin) = "HumanPlayer" Then
        Origin.Beaten = New PropCollection
        Origin.Quests = New PropCollection
    ElseIf TypeName(Origin) = "ComPlayer" Then
        Origin.Script       = Scripts.ObjectByName(Rng.Offset(0, 0).Value)
        Origin.ScriptValues = Rng.Offset(0, 1).Value
    End If
    ' Offset 5 is Textureindex needed for initializing same texture for players
End Sub