VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit


Public Property Get Value() As Variant
End Property
Public Property Let Value(ByVal n_Value As Variant)
End Property

Public Property Get Formula() As Variant
End Property
Public Property Let Formula(ByVal n_Formula As Variant)
End Property

Public Function Offset(ByVal RowOffset As Long, ByVal ColOffset As Long) As IRange
End Function