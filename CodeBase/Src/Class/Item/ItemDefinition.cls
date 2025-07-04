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
Private p_Script             As Script
Private p_TextureDefinition  As TextureDefinition

Public Property Let Number(n_Number                       As Long)               : p_Number                = n_Number            : End Property
Public Property Let Name(n_Name                           As String)             : p_Name                  = n_Name              : End Property
Public Property Let Script(n_Value                        As Script)             : Set p_Script                = n_Value             : End Property
Public Property Let TextureDefinition(n_TextureDefinition As TextureDefinition)  : Set p_TextureDefinition = n_TextureDefinition : End Property


Public Property Get Number()                              As Long                : Number                  = p_Number            : End Property
Public Property Get Name()                                As String              : Name                    = p_Name              : End Property
Public Property Get Script()                              As Script              : Set Script                  = p_Script            : End Property
Public Property Get TextureDefinition()                   As TextureDefinition   : Set TextureDefinition   = p_TextureDefinition : End Property

Public Function Create(Rng As Range) As ItemDefinition
    Set Create = New ItemDefinition
    With Create
        .Number            = Rng.Offset(0, 0).Value
        .Name              = Rng.Offset(0, 1).Value
        .Script            = FumonGame.FactoryServer.Script(Rng.Offset(0, 2).Value)
        .TextureDefinition = FumonGame.TextureDefinition.Create(Rng.Offset(0, 3).Value, Rng.Offset(0, 4).Value, FumonGame.FactoryServer.GameMap.Folder.Value & "\Item.png")
    End With
End Function