VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Number         As Long
Private p_Name           As String
Private p_Command        As VBGLCallable
Private p_ArgDef()       As String
Private p_Active         As Long
Private p_Minute         As Long

Public Property Let Number(ByVal n_Number       As Long)         : Let p_Number    = n_Number   : End Property
Public Property Let Name(ByVal n_Name           As String)       : Let p_Name      = n_Name     : End Property
Public Property Let Command(ByVal n_Command     As VBGLCallable) : Set p_Command   = n_Command  : End Property
Public Property Let ArgDef(ByVal n_ArgDef       As Variant)      : Let p_ArgDef    = n_ArgDef   : End Property
Public Property Let Active(ByVal n_Active       As Long)         : Let p_Active    = n_Active   : End Property
Public Property Let Minute(ByVal n_Minute       As Long)         : Let p_Minute    = n_Minute   : End Property


Public Property Get Number()              As Long                : Let Number      = p_Number   : End Property
Public Property Get Name()                As String              : Let Name        = p_Name     : End Property
Public Property Get Command()             As VBGLCallable        : Set Command     = p_Command  : End Property
Public Property Get ArgDef()              As Variant             : Let ArgDef      = p_ArgDef   : End Property
Public Property Get Active()              As Long                : Let Active      = p_Active   : End Property
Public Property Get Minute()              As Long                : Let Minute      = p_Minute   : End Property


Public Function Create(Rng As Range) As Script
    Dim Temp() As String
    Temp = Split(Rng.Offset(0, 1).Value, ".")
    Set Create = New Script
    With Create
        .Number   = Rng.Offset(0, 0).Value
        .Name     = Rng.Offset(0, 1).Value
        .Command = CreateCommand(Rng.Offset(0, 2).Value)
        .ArgDef   = CreateArgDef(Rng.Offset(0, 3).Value)
        .Active   = Rng.Offset(0, 4).Value
        .Minute   = Rng.Offset(0, 5).Value
    End With
End Function

Public Function Run(ParamArray Arguments() As Variant) As Variant
    Dim Temp() As Variant
    Temp = Arguments
    Call VBGLAssign(Run, RunArr(Temp))
End Function

Public Function RunArr(Arguments() As Variant) As Variant
    Call VBGLAssign(RunArr, Command.HandleArr(Arguments))
End Function

Private Function CreateCommand(ByVal Text As String) As VBGLCallable
    Dim Obj As Object
    Dim i As Long
    Dim Args() As String
    Args = Split(Text, ", ")
    If Args(0) <> "Nothing" Then Set Obj = Nothing ' Currently only RunApplication
    Set CreateCommand = VBGLCallable.Create(Obj, Args(1), CallTypee(Args(2)), CLng(Args(3)))
End Function

Private Function CreateArgDef(ByVal Text As String) As String()
    Dim Texts()     As String : Texts = Split(Text, vbCrLf)
    Dim ReturnArr() As String
    Dim i As Long, j As Long
    For i = 0 To USize(Texts)
        Dim Types() As String
        Types = Split(Texts(i), ", ")
        For j = 0 To USize(Types)
            Call VBGLAdd(ReturnArr, Types(j))
        Next j
    Next i
    CreateArgDef = ReturnArr
End Function

Private Function CallTypee(ByVal Text As String) As vbCallType
    Select Case Text
        Case vbMethod : CallTypee = vbCallType.vbMethod
        Case vbLet    : CallTypee = vbCallType.vbLet
        Case vbGet    : CallTypee = vbCallType.vbGet
        Case vbSet    : CallTypee = vbCallType.vbSet
    End Select
End Function