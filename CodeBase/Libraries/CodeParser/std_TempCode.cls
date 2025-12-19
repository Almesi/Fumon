VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_TempCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private Proj As VBIDE.VBProject
Private Component As VBIDE.VBComponent
Private Functions() As String
Private Size As Long

Public Function Create(ByVal n_Proj As VBIDE.VBProject) As std_TempCode
    Set Create = New std_TempCode
    With Create
        Call .Init(n_Proj)
    End With
End Function

Public Sub Init(ByVal n_Proj As VBIDE.VBProject)
    Set Proj = n_Proj
    Set Component = Proj.VBComponents.Add(vbext_ct_StdModule)
    Component.Name = "TEMP" & ObjPtr(Component)
End Sub

Public Sub AddFunction(ByVal Name As String, ByVal Code As String)
    Dim Found As Long
    Found = Find(Name)
    If Found = -1 Then
        Call Add(Name)
        Call Component.CodeModule.AddFromString(Code)
    End If
End Sub

Public Function CallFunction(ByVal Name As String) As Variant
    Dim Found As Long
    Found = Find(Name)
    If Found <> -1 Then Call Assign(CallFunction, Application.Run(Name))
End Function

Private Function Add(ByVal Name As String) As Long
    Size = Size + 1
    ReDim Preserve Functions(Size)
    Functions(Size) = Name
    Add = Size
End Function

Private Function Find(ByVal Name As String) As Long
    Dim i As Long
    For i = 0 To Size
        If Functions(i) = Name Then
            Find = i
            Exit Function
        End If
    Next i
    Find = -1
End Function

Private Sub Assign(ByRef Goal As Variant, ByVal Value As Variant)
    If IsObject(Value) Then
        Set Goal = Value
    Else
        Let Goal = Value
    End If
End Sub

Private Sub Class_Initialize()
    Size = -1
End Sub

Private Sub Class_Terminate()
    If Not Component Is Nothing Then
        Call Proj.VBComponents.Remove(Component)
    End If
End Sub