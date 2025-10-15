VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Tiles() As Tile

Public Property Get Tiles() As Variant
    Tiles = p_Tiles
End Property
Public Property Let Tiles(n_Tiles As Variant)
    p_Tiles = n_Tiles
End Property

Public Property Get Tile(ByVal y As Long, ByVal x As Long) As Tile
    Set Tile = p_Tiles(y, x)
End Property
Public Property Let Tile(ByVal y As Long, ByVal x As Long, ByVal n_Tile As Tile)
    Set p_Tiles(y, x) = n_Tile
End Property

Public Function Create(ByVal Pointer As Range, ByVal Rows As Long, ByVal Columns As Long) As Tiles
    Dim y As Long, x As Long
    Set Create = New Tiles
    Dim Temp() As Tile
    ReDim Temp(Rows, Columns)
    With Create
        For y = 0 To Rows
            For x = 0 To Columns
                Set Temp(y, x) = FactoryTile.Create(Pointer.Offset(y, x))
            Next x
        Next y
        .Tiles = Temp
    End With
End Function

Public Function TileData(ByVal GetWhat As String) As Variant
    Dim y As Long, x As Long
    Dim MaxY        As Long  : MaxY = USize(Tiles, 1)
    Dim MaxX        As Long  : MaxX = USize(Tiles, 2)
    Dim ReturnArr   As Variant: ReturnArr = GetArr(MaxY, MaxX, Tile(0, 0).ExtractPoint(GetWhat))
    For y = 0 To MaxY
        For x = 0 To MaxX
            Call VBGLAssign(ReturnArr(y, x), Tile(y, x).ExtractPoint(GetWhat))
        Next x
    Next y
    TileData = ReturnArr
End Function

Private Function GetArr(ByVal Rows As Long, ByVal Columns As Long, ByVal Var As Variant) As Variant
    Select Case TypeName(Var)
        Case "Long"
            Dim TempLong()           As Long
            ReDim TempLong(Rows, Columns)
            GetArr = TempLong
        Case "Single"
            Dim TempSingle()         As Single
            ReDim TempSingle(Rows, Columns)
            GetArr = TempSingle
        Case "String"
            Dim TempString()         As String
            ReDim TempString(Rows, Columns)
            GetArr = TempString
        Case "VBGLSubTexture"
            Dim TempVBGLSubTexture() As VBGLSubTexture
            ReDim TempVBGLSubTexture(Rows, Columns)
            GetArr = TempVBGLSubTexture
        Case "IPlayer"
            Dim TempPlayer()        As IPlayer
            ReDim TempPlayer(Rows, Columns)
            GetArr = TempPlayer
        Case "Tile"
            Dim TempTile()           As Tile
            ReDim TempTile(Rows, Columns)
            GetArr = TempTile
        Case "Script"
            Dim TempScript()         As Script
            ReDim TempScript(Rows, Columns)
            GetArr = TempScript
    End Select
End Function