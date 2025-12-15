VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Objects() As Object
Private p_Count As Long

Public Property Get Count() As Long
    Count = p_Count
End Property

Public Sub LetObjects(ByRef n_Objects As Object)
    p_Objects = n_Objects
    p_Count = USize(p_Objects)
End Sub

Public Sub GetObjects(ByRef Arr As Variant)
    Dim i As Long
    ReDim Arr(Count)
    For i = 0 To Count
        Call VBGLAssign(Arr(i), Object(i))
    Next i
End Sub

Public Property Get Objects() As Object()
    Objects = p_Objects
End Property

Public Property Get Object(ByVal Index As Long) As Object
    Set Object = p_Objects(Index)
End Property

Public Function Create(ByRef n_Objects As Variant) As PropCollection
    Set Create = New PropCollection
    Dim i As Long
    With Create
        For i = 0 To USize(n_Objects)
            Call .Add(n_Objects(i))
        Next i
    End With
End Function

Public Function Add(ByVal n_Object As Object) As Long
    Call VBGLAdd(p_Objects, n_Object)
    p_Count = USize(p_Objects)
    Add = Count
End Function

Public Sub Remove(ByVal Index As Long)
    Dim NewSize As Long
    NewSize = Count - 1
    Dim i As Long
    Dim Arr() as Object
    ReDim Arr(NewSize)
    For i = 0 To Index - 1
        Set Arr(i) = Object(i)
    Next i
    For i = Index + 1 To Count
        Set Arr(i - 1) = Object(i)
    Next i
    p_Objects = Arr
    p_Count = NewSize
End Sub

Public Sub AddUnique(ByVal n_Object As Object)
    If Find(n_Object) = -1 Then Call VBGLAdd(p_Objects, n_Object)
End Sub

Public Function Find(ByVal Search As Object) As Long
    Dim i As Long
    For i = 0 To Count
        If VBGLCompare(Object(i), Search) Then
            Find = i
            Exit Function
        End If
    Next i
    Find = -1
End Function

Public Function ObjectByName(ByVal Name As String) As Object
    On Error GoTo Error
    Dim i As Long
    For i = 0 To Count
        If Object(i).Name = Name Then
            Set ObjectByName = Object(i)
            Exit Function
        End If
    Next i
    Error:
End Function

Public Function Properties(ByVal Property As String) As Variant
    Dim i As Long
    Dim ReturnArr As Variant
    ReDim ReturnArr(0)
    Erase ReturnArr
    For i = 0 To Count
        Call VBGLAdd(ReturnArr, CallByName(Object(i), Property, vbGet))
    Next i
    Properties = ReturnArr
End Function

Private Sub Class_Initialize()
    p_Count = -1
End Sub