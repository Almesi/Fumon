VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ElementTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Elements() As Range

Public Property Get Elements() As Range()
    Elements = p_Elements
End Property
Public Property Let Elements(n_Elements() As Range)
    p_Elements = n_Elements
End Property

Public Property Get Element(Index As Long) As Range
    Set Element = p_Elements(Index)
End Property
Public Property Let Element(Index As Long, n_Element As Range)
    Set p_Elements(Index) = n_Element
End Property

Public Function Create(ByVal Rng As Range) As ElementTypes
    Dim i As Long
    Dim Count As Long
    Do While Rng.Offset(0, Count).Formula <> Empty
        Count = Count + 1
    Loop

    Set Create = New ElementTypes
    Dim Temp() As Range
    ReDim Temp(Count - 1)

    With Create
        .Elements = Temp
        For i = 0 To Count - 1
            .Element(i) = Rng.Offset(0, i)
        Next i
    End With
End Function

Public Function GetFactor(AttElement As String, Fumon As Fumon) As Single
    Dim AttElementIndex As Long
    Dim DefElement1Index As Long
    Dim DefElement2Index As Long
    AttElementIndex  = GetIndex(AttElement)
    DefElement1Index = GetIndex(Fumon.Definition.Type1) + 1
    DefElement2Index = GetIndex(Fumon.Definition.Type2) + 1

    If DefElement2Index = 0 Then
        GetFactor = Element(AttElementIndex).Offset(DefElement1Index, 0).Value
    Else
        GetFactor = Element(AttElementIndex).Offset(DefElement1Index, 0).Value * Element(AttElementIndex).Offset(DefElement2Index, 0).Value
    End If
End Function

Private Function GetIndex(NameOrIndex As String) As Long
    Dim i As Long
    GetIndex = -1
    If IsNumeric(NameOrIndex) Then
        GetIndex = NameOrIndex
    Else
        If NameOrIndex = Empty Then Exit Function
        For i = 0 To Ubound(Elements)
            If Element(i).Value = NameOrIndex Then
                GetIndex = i
                Exit Function
            End If
        Next i
    End If
End Function