VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TileDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit


Private p_Number            As Long
Private p_Name              As String
Private p_TileType          As String
Private p_Speed             As Single
Private p_Depth             As Single

Public Property Let Number(n_Number                       As Long)              : Let p_Number            = n_Number            : End Property
Public Property Let Name(n_Name                           As String)            : Let p_Name              = n_Name              : End Property
Public Property Let TileType(n_TileType                   As String)            : Let p_TileType          = n_TileType          : End Property
Public Property Let Speed(n_Speed                         As Single)            : Let p_Speed             = n_Speed             : End Property
Public Property Let Depth(n_Depth                         As Single)            : Let p_Depth             = n_Depth             : End Property

Public Property Get Number()                              As Long               : Let Number              = p_Number            : End Property
Public Property Get Name()                                As String             : Let Name                = p_Name              : End Property
Public Property Get TileType()                            As String             : Let TileType            = p_TileType          : End Property
Public Property Get Speed()                               As Single             : Let Speed               = p_Speed             : End Property
Public Property Get Depth()                               As Single             : Let Depth               = p_Depth             : End Property

Public Function Create(Rng As Range) As TileDefinition
    Set Create = New TileDefinition
    With Create
        .Number     = Rng.Offset(0, 0).Value
        .Name       = Rng.Offset(0, 1).Value
        .TileType   = Rng.Offset(0, 2).Value
        .Speed      = Rng.Offset(0, 3).Value
        .Depth      = Rng.Offset(0, 4).Value
    End With
End Function