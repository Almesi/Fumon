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


Private p_Script       As Script
Private p_ScriptValues As String
Private p_Beaten       As PropCollection
Private p_Quests       As PropCollection


Public Function Offset() As Long
    Offset = 5
End Function

Public Property Let Script(ByVal       n_Script       As Script)         : Set p_Script       = n_Script       : End Property
Public Property Let ScriptValues(ByVal n_ScriptValues As String)         : Let p_ScriptValues = n_ScriptValues : End Property
Public Property Let Beaten(ByVal       n_Beaten       As PropCollection) : Set p_Beaten       = n_Beaten       : End Property
Public Property Let Quests(ByVal       n_Quests       As PropCollection) : Set p_Quests       = n_Quests       : End Property

Public Property Get Script()                          As Script          : Set Script         = p_Script       : End Property
Public Property Get ScriptValues()                    As String          : Let ScriptValues   = p_ScriptValues : End Property
Public Property Get Beaten()                          As PropCollection  : Set Beaten         = p_Beaten       : End Property
Public Property Get Quests()                          As PropCollection  : Set Quests         = p_Quests       : End Property

Public Function Create(ByVal Rng As IRange, ByVal Scripts As PropCollection) As SpecificBase
    Set Create = New SpecificBase
    With Create
        .Script       = Scripts.ObjectByName(Rng.Offset(0, 0).Value)
        .ScriptValues = Rng.Offset(0, 1).Value
        .Beaten       = New PropCollection
        .Quests       = New PropCollection
    End With
End Function

Public Function RunScript(ByVal ScriptHolder As IPlayer, ByVal Caller As IPlayer) As Variant
    Dim Args() As Variant
    Args = InterpretArguments(ScriptHolder, Caller)
    Call VBGLAssign(RunScript, Script.RunArr(Args))
End Function

Private Function InterpretArguments(ByVal ScriptHolder As IPlayer, ByVal Caller As IPlayer) As Variant()
    Dim Types()   As String    : Types = Split(ScriptValues, ", ")
    Dim ReturnArr() As Variant
    Dim i As Long
    For i = 0 To USize(Types)
        Call VBGLAdd(ReturnArr, InterpretArgument(ScriptHolder, Caller, Types(i)))
    Next i
    InterpretArguments = ReturnArr
End Function

Private Function InterpretArgument(ByVal ScriptHolder As IPlayer, ByVal Caller As IPlayer, ByVal Value As String) As Variant
    Select Case Value
        Case "ME"  : Set InterpretArgument = ScriptHolder
        Case "YOU" : Set InterpretArgument = Caller
        Case Else  : Let InterpretArgument = CVar(Value)
    End Select
End Function