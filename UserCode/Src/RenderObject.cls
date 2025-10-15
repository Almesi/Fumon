VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RenderObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private p_Drawables() As Object
Private p_Input       As VBGLIInput
Private p_Clear       As Boolean

Public Property Let Inputt(ByVal n_Value As VBGLIInput)
    Set p_Input = n_Value
End Property

Public Property Get Drawable(Index As Long) As Object
    Set Drawable = p_Drawables(Index)
End Property
Public Property Get Inputt() As VBGLIInput
    Set Inputt = p_Input
End Property

Public Property Let Clear(ByVal n_Clear As Boolean)
    p_Clear = n_Clear
End Property

Public Property Get Clear() As Boolean
    Clear = p_Clear
End Property

Public Function Create(ByVal n_Input As VBGLIInput) As RenderObject
    Set Create = New RenderObject
    With Create
        .Inputt = n_Input
    End With
End Function

Public Sub AddDrawable(n_Drawable As Object)
    Call VBGLAdd(p_Drawables, n_Drawable)
End Sub

Public Sub RemoveDrawable()
    Call VBGLArrayPop(p_Drawables)
End Sub


Public Sub Loopp()
    Dim i As Long
    If Clear Then
        Call CurrentContext.Clear()
        Call CurrentContext.ClearColor(0, 0, 1, 1)
    End If
    For i = 0 To USize(p_Drawables)
        Call Drawable(i).Draw()
        'Debug.Print "FPS:", CurrentContext.CurrentWindow.LimitFPS
    Next i
    Call glutSwapBuffers
End Sub

Public Sub KeyBoard(ByVal Char As Byte, ByVal x As Long, ByVal y As Long)
    Call Inputt.KeyDown(Char, x, y)
End Sub

Public Sub KeyBoardUp(ByVal Char As Byte, ByVal x As Long, ByVal y As Long)
    Call Inputt.KeyUp(Char, x, y)
End Sub

Public Sub MouseMove(ByVal x As Long, ByVal y As Long)
    Call Inputt.KeyDown(0, x, y)
End Sub

Public Sub MouseWheel(ByVal wheel As Long, ByVal direction As Long, ByVal x As Long, ByVal y As Long)
    Call Inputt.MouseWheel(wheel, direction, x, y)
End Sub

Private Sub Class_Initialize()
    Clear = True
End Sub