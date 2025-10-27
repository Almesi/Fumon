VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit


Private p_Number         As Long
Private p_Name           As String
Private p_Description    As String
Private p_Condition      As String

Public Property Let Number(n_Number           As Long)   : p_Number      = n_Number      : End Property
Public Property Let Name(n_Name               As String) : p_Name        = n_Name        : End Property
Public Property Let Description(n_Description As String) : p_Description = n_Description : End Property
Public Property Let Condition(n_Condition     As String) : p_Condition   = n_Condition   : End Property

Public Property Get Number()                  As Long    : Number        = p_Number      : End Property
Public Property Get Name()                    As String  : Name          = p_Name        : End Property
Public Property Get Description()             As String  : Description   = p_Description : End Property
Public Property Get Condition()               As String  : Condition     = p_Condition   : End Property

Public Function Create(ByVal Rng As Range) As Quest
    Set Create = New Quest
    With Create
        .Number       = Rng.Offset(0, 0).Value
        .Name         = Rng.Offset(0, 1).Value
        .Description  = Rng.Offset(0, 2).Value
        .Condition    = Rng.Offset(0, 3).Value
    End With
End Function