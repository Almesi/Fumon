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
Private p_Command        As std_Callable
Private p_Active         As Long
Private p_Minute         As Long

Public Property Let Number(ByVal n_Number       As Long)         : Let p_Number    = n_Number   : End Property
Public Property Let Name(ByVal n_Name           As String)       : Let p_Name      = n_Name     : End Property
Public Property Let Command(ByVal n_Command     As std_Callable) : Set p_Command   = n_Command  : End Property
Public Property Let Active(ByVal n_Active       As Long)         : Let p_Active    = n_Active   : End Property
Public Property Let Minute(ByVal n_Minute       As Long)         : Let p_Minute    = n_Minute   : End Property


Public Property Get Number()                    As Long          : Let Number      = p_Number   : End Property
Public Property Get Name()                      As String        : Let Name        = p_Name     : End Property
Public Property Get Command()                   As std_Callable  : Set Command     = p_Command  : End Property
Public Property Get Active()                    As Long          : Let Active      = p_Active   : End Property
Public Property Get Minute()                    As Long          : Let Minute      = p_Minute   : End Property


Public Function Create(ByVal Rng As IRange) As Script
    Dim Temp() As String
    Temp = Split(Rng.Offset(0, 1).Value, ".")
    Set Create = New Script
    With Create
        .Number   = Rng.Offset(0, 0).Value
        .Name     = Rng.Offset(0, 1).Value
        .Command  = AllCallables.CreateCallable(Rng.Offset(0, 2).Value)
        .Active   = Rng.Offset(0, 3).Value
        .Minute   = Rng.Offset(0, 4).Value
    End With
End Function

Public Function Run(ParamArray Arguments() As Variant) As Variant
    Dim Temp() As Variant
    Temp = Arguments
    Call VBGLAssign(Run, RunArr(Temp))
End Function

Public Function RunArr(Arguments() As Variant) As Variant
    Call VBGLAssign(RunArr, Command.RunArr(Arguments))
End Function