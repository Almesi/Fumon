VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit


Private p_Number            As Long
Private p_Name              As String
Private p_TileType          As String
Private p_Speed             As Double
Private p_TextureDefinition As TextureDefinition

Public Property Let Number(n_Number                       As Long)              : Let p_Number            = n_Number            : End Property
Public Property Let Name(n_Name                           As String)            : Let p_Name              = n_Name              : End Property
Public Property Let TileType(n_TileType                   As String)            : Let p_TileType          = n_TileType          : End Property
Public Property Let Speed(n_Speed                         As Double)            : Let p_Speed             = n_Speed             : End Property
Public Property Let TextureDefinition(n_TextureDefinition As TextureDefinition) : Set p_TextureDefinition = n_TextureDefinition : End Property

Public Property Get Number()                              As Long               : Let Number              = p_Number            : End Property
Public Property Get Name()                                As String             : Let Name                = p_Name              : End Property
Public Property Get TileType()                            As String             : Let TileType            = p_TileType          : End Property
Public Property Get Speed()                               As Double             : Let Speed               = p_Speed             : End Property
Public Property Get TextureDefinition()                   As TextureDefinition  : Set TextureDefinition   = p_TextureDefinition : End Property

Public Function Create(Rng As Range, FolderPath As String) As Tile
    Set Create = New Tile
    With Create
        .Number            = Rng.Offset(0, 0).Value
        .Name              = Rng.Offset(0, 1).Value
        .TileType          = Rng.Offset(0, 2).Value
        .Speed             = Rng.Offset(0, 3).Value
        .TextureDefinition = FumonGame.TextureDefinition.Create(Rng.Offset(0, 4).Value, Rng.Offset(0, 5).Value, FolderPath & "\Tile.png")
    End With
End Function