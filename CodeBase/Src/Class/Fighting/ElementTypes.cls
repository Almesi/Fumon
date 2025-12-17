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

Private p_Elements() As IRange

Public Property Get Elements() As IRange()
    Elements = p_Elements
End Property
Public Property Let Elements(n_Elements() As IRange)
    p_Elements = n_Elements
End Property

Public Property Get Element(ByVal Index As Long) As IRange
    Set Element = p_Elements(Index)
End Property
Public Property Let Element(ByVal Index As Long, ByVal n_Element As IRange)
    Set p_Elements(Index) = n_Element
End Property

Public Function Create(ByVal Rng As IRange) As ElementTypes
    Dim i As Long
    Dim Count As Long
    Count = Rng.GetSelf.Parent.RowCount - 1

    Set Create = New ElementTypes
    Dim Temp() As IRange
    ReDim Temp(Count - 1)

    With Create
        .Elements = Temp
        For i = 0 To Count - 1
            .Element(i) = Rng.Offset(0, i)
        Next i
    End With
End Function

Public Function GetFactor(ByVal AttElement As String, ByVal Fumon As Fumon) As Single
    Dim AttElementIndex  As Long : AttElementIndex  = GetIndex(AttElement)
    Dim DefElement1Index As Long : DefElement1Index = GetIndex(Fumon.Definition.Type1)
    Dim DefElement2Index As Long : DefElement2Index = GetIndex(Fumon.Definition.Type2)

    If DefElement2Index = -1 Then
        GetFactor = Element(AttElementIndex).Offset(DefElement1Index, 0).Value
    Else
        GetFactor = Element(AttElementIndex).Offset(DefElement1Index, 0).Value * Element(AttElementIndex).Offset(DefElement2Index, 0).Value
    End If
End Function

Public Function Color(ByVal NameOrIndex As String) As Single()
    Dim Found As Long     : Found = GetIndex(NameOrIndex)
    Dim Arr() As Single   : ReDim Arr(3)
    If Found <> -1 Then
        Dim ColorValue As Long
        ColorValue = Element(Found).GetSelf.Parent.Rng.Interior.Color
        Arr(0) = ColorValue Mod 256
        Arr(1) = (ColorValue \ 256) Mod 256
        Arr(2) = (ColorValue \ 65536) Mod 256
        Arr(3) = (ColorValue \ 16777216) Mod 256
    End If
    Color = Arr
End Function

Private Function GetIndex(ByVal NameOrIndex As String) As Long
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