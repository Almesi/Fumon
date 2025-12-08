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


Private p_Number As IRange
Private p_Name   As IRange

Public Function Offset() As Long
    Offset = 2
End Function

Public Property Let Number(ByVal n_Number  As IRange) : Set p_Number = n_Number  : End Property
Public Property Let Name(ByVal n_Name      As IRange) : Set p_Name   = n_Name    : End Property

Public Property Get Number()               As IRange  : Set Number   = p_Number  : End Property
Public Property Get Name()                 As IRange  : Set Name     = p_Name    : End Property

Public Function Create(ByVal Rng As IRange) As PlayerBase
    Set Create = New PlayerBase
    With Create
        .Number        = Rng.Offset(0, 0)
        .Name          = Rng.Offset(0, 1)
    End With
End Function