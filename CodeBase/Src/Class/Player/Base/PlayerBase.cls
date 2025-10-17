VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlayerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit


Private p_Number As Range
Private p_Name   As Range

Public Function Offset() As Long
    Offset = 2
End Function

Public Property Let Number(ByVal n_Number  As Range) : Set p_Number = n_Number  : End Property
Public Property Let Name(ByVal n_Name      As Range) : Set p_Name   = n_Name    : End Property

Public Property Get Number()               As Range  : Set Number   = p_Number  : End Property
Public Property Get Name()                 As Range  : Set Name     = p_Name    : End Property

Public Function Create(ByVal Rng As Range) As PlayerBase
    Set Create = New PlayerBase
    With Create
        .Number        = Rng.Offset(0, 0)
        .Name          = Rng.Offset(0, 1)
    End With
End Function