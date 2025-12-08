VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NormalRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit
Implements IRange

Private p_Rng As IRange

Private Property Let IRange_Value(ByVal n_Value As Variant)
    Rng.Value = n_Value
End Property
Private Property Get IRange_Value() As Variant
    IRange_Value = Rng.Value
End Property

Private Property Let IRange_Formula(ByVal n_Formula As Variant)
    Rng.Formula = n_Formula
End Property
Private Property Get IRange_Formula() As Variant
    IRange_Formula = Rng.Formula
End Property

Private Function IRange_Offset(ByVal RowOffset As Long, ByVal ColOffset As Long) As IRange
    Set IRange_Offset = Rng.Offset(RowOffset, ColOffset)
End Function

Public Property Let Rng(ByVal n_Rng As IRange)
    Set p_Rng = n_Rng
End Property
Public Property Get Rng() As IRange
    Set Rng = p_Rng
End Property

Public Function Create(ByVal n_Rng As IRange) As NormalRange
    Set Create = New NormalRange
    Create.Rng = n_Rng
End Function