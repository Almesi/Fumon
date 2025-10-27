VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UpdateQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit


Private Type Thread
    Always()        As std_Callable
    Once()          As std_Callable
    AfterMinutes    As Long
    LastTime        As Long
    Index           As Long
End Type

Private Type Frame
    Threads() As Thread
    Size As Long
End Type

Private Frames As Frame

Public Function Create() As UpdateQueue
    Set Create = New UpdateQueue
End Function

Public Sub AddThread(ByRef Always() As std_Callable, ByRef Once() As std_Callable, ByVal AfterMinutes As Long)
    Dim NewSize As Long
    NewSize = Frames.Size + 1
    ReDim Preserve Frames.Threads(NewSize)

    Frames.Threads(NewSize).Always       = Always
    Frames.Threads(NewSize).Once         = Once
    Frames.Threads(NewSize).AfterMinutes = AfterMinutes
    Frames.Size = NewSize
End Sub

Public Sub AddAlways(ByVal n_Callable As std_Callable, Optional ByVal Thread As Long = 0)
    Call VBGLAdd(Frames.Threads(Thread).Always, n_Callable)
End Sub

Public Sub AddOnce(ByVal n_Callable As std_Callable, Optional ByVal Thread As Long = 0)
    Call VBGLAdd(Frames.Threads(Thread).Once, n_Callable)
End Sub

Public Sub UpdateForEach(ByRef std_Callables() As std_Callable)
    Dim i As Long
    For i = 0 To USize(std_Callables)
        Call std_Callables(i).Run
        Call UpdateScreen()
    Next i
End Sub

Public Sub Update(ByVal CurrentMinute As Long)
    Dim i As Long
    
    If Not CurrentRenderObject Is OverWorldRenderObject Then Exit Sub ' For Debugging
    For i = 0 To Frames.Size
        Dim Diff As Long
        Diff = (CurrentMinute - Frames.Threads(i).LastTime)
        If Diff >= Frames.Threads(i).AfterMinutes Then
            Call RunAlways(Frames.Threads(i))
            Call RunOnce(Frames.Threads(i))
        End If
        Frames.Threads(i).LastTime = CurrentMinute
    Next i

    Call UpdateScreen()
End Sub

Public Sub UpdateScreenEvent()
    Call glutMainLoopEvent()
    Call UpdateScreen()
End Sub

Public Sub UpdateScreen()
    Call CurrentRenderObject.Loopp
End Sub

Private Function RunAlways(ByRef Thread As Thread) As Variant
    Call VBGLAssign(RunAlways, Thread.Always(Thread.Index).Run())
    Thread.Index = Thread.Index + 1
    If Thread.Index > USize(Thread.Always) Then Thread.Index = 0
End Function

Private Function RunOnce(ByRef Thread As Thread) As Variant
    If USize(Thread.Once) = -1 Then Exit Function
    Call VBGLAssign(Thread.Once, Thread.Once(0).Run())
    Call VBGLArrayRemove(Thread.Once, 0)
End Function

Private Sub Class_Initialize()
    Frames.Size = -1
End Sub