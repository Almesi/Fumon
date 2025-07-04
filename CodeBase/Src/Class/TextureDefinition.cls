VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextureDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_X        As Double
Private p_Y        As Double
Private p_FilePath As String

Public Property Let X(n_X               As Double) : p_X        = n_X        : End Property
Public Property Let Y(n_Y               As Double) : p_Y        = n_Y        : End Property
Public Property Let FilePath(n_FilePath As String) : p_FilePath = n_FilePath : End Property

Public Property Get X()                 As Double  : X          = p_X        : End Property
Public Property Get Y()                 As Double  : Y          = p_Y        : End Property
Public Property Get FilePath()          As String  : FilePath   = p_FilePath : End Property

Public Function Create(n_X As Double, n_Y As Double, n_FilePath As String) As TextureDefinition
    Set Create = New TextureDefinition
    With Create
        .X        = n_X
        .Y        = n_Y
        .FilePath = n_FilePath
    End With
End Function