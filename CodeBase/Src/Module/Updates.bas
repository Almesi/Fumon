Attribute VB_Name = "Updates"


Option Explicit

Public Function CreateMovePathScriptsParam(ByVal Player As IPlayer, ParamArray Points() As Variant) As VBGLCallable()
    Dim Size As Long
    Size = ((Ubound(Points) + 1) / 2) - 1

    Dim x() As Long
    Dim y() As Long
    ReDim x(Size)
    ReDim y(Size)

    Dim i As Long
    For i = 0 To Size
        x(i) = Points(i * 2 + 0)
        y(i) = Points(i * 2 + 1)
    Next i

    CreateMovePathScriptsParam = CreateMovePathScripts(Player, x, y)
End Function

Public Function CreateMovePathScripts(ByVal Player As IPlayer, ByRef x() As Long, ByRef y() As Long) As VBGLCallable()
    Dim Size As Long
    Size = USize(x)
    Dim Arr() As VBGLCallable
    Dim i As Long
    For i = 0 To Size - 1
        Call VBGLMerge(Arr, CreateMovePath(Player, x(i), x(i + 1), y(i), y(i + 1)))
    Next i
    Call VBGLMerge(Arr, CreateMovePath(Player, x(i), x(0), y(i), y(0)))
    CreateMovePathScripts = Arr
End Function

Private Function CreateMovePath(ByVal Player As IPlayer, ByVal StartX As Long, ByVal EndX As Long, ByVal StartY As Long, ByVal EndY As Long) As VBGLCallable()
    Dim x() As Long
    Dim y() As Long
    Call FindPathAlgorithm(StartX, EndX, StartY, EndY, x, y)
    Dim i As Long
    Dim Size As Long
    Size = Usize(x)
    Dim Arr() As VBGLCallable
    ReDim Arr(Size)
    For i = 0 To Size
        Set Arr(i)     = VBGLCallable.Create(Player , "Move", vbMethod, 1, x(i), y(i))
    Next i
    CreateMovePath = Arr
End Function

Public Function CreateUpdateOverWorld() As VBGLCallable()
    Dim Arr() As VBGLCallable
    ReDim Arr(0)
    Set Arr(0) = VBGLCallable.Create(Nothing, "UpdateOverWorld", vbMethod, 0, True)
    CreateUpdateOverWorld = Arr
End Function