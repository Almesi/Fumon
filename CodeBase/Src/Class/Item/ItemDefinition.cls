VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Number             As Long
Private p_Name               As String
Private p_Command            As std_Callable

Public Property Let Number(n_Number                       As Long)               : p_Number                = n_Number     : End Property
Public Property Let Name(n_Name                           As String)             : p_Name                  = n_Name       : End Property
Public Property Let Command(n_Value                       As std_Callable)       : Set p_Command           = n_Value      : End Property


Public Property Get Number()                              As Long                : Number                  = p_Number     : End Property
Public Property Get Name()                                As String              : Name                    = p_Name       : End Property
Public Property Get Command()                             As std_Callable        : Set Command             = p_Command    : End Property

Public Function Create(ByVal Rng As IRange) As ItemDefinition
    Set Create = New ItemDefinition
    With Create
        .Number            = Rng.Offset(0, 0).Value
        .Name              = Rng.Offset(0, 1).Value
        .Command           = AllCallables.CreateCallable(Rng.Offset(0, 2).Value)
    End With
End Function