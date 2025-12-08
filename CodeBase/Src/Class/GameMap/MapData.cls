VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Year          As IRange
Private p_Month         As IRange
Private p_Day           As IRange
Private p_Hour          As IRange
Private p_Minute        As IRange
Private p_Folder        As IRange
Private p_ServerStarter As IRange
Private p_Maps          As PropCollection


Public Property Let Year(n_Year                   As IRange)            : Set p_Year          = n_Year           : End Property
Public Property Let Month(n_Month                 As IRange)            : Set p_Month         = n_Month          : End Property
Public Property Let Day(n_Day                     As IRange)            : Set p_Day           = n_Day            : End Property
Public Property Let Hour(n_Hour                   As IRange)            : Set p_Hour          = n_Hour           : End Property
Public Property Let Minute(n_Minute               As IRange)            : Set p_Minute        = n_Minute         : End Property
Public Property Let Folder(n_Folder               As IRange)            : Set p_Folder        = n_Folder         : End Property
Public Property Let ServerStarter(n_ServerStarter As IRange)            : Set p_ServerStarter = n_ServerStarter  : End Property
Public Property Let Maps(n_Maps                   As PropCollection)   : Set p_Maps          = n_Maps           : End Property

Public Property Get Year()                        As IRange             : Set Year            = p_Year           : End Property
Public Property Get Month()                       As IRange             : Set Month           = p_Month          : End Property
Public Property Get Day()                         As IRange             : Set Day             = p_Day            : End Property
Public Property Get Hour()                        As IRange             : Set Hour            = p_Hour           : End Property
Public Property Get Minute()                      As IRange             : Set Minute          = p_Minute         : End Property
Public Property Get Folder()                      As IRange             : Set Folder          = p_Folder         : End Property
Public Property Get ServerStarter()               As IRange             : Set ServerStarter   = p_ServerStarter  : End Property
Public Property Get Maps()                        As PropCollection    : Set Maps            = p_Maps           : End Property

Public Property Get Map(ByVal Index As Long) As GameMap
    Set Map = Maps.Object(Index)
End Property

Public Function Create(ByVal WB As Workbook, ByVal Rng As IRange, ByVal TileDefinitions As PropCollection) As MapData
    Set Create = New MapData
    Dim Temp As Tiles: Set Temp = New Tiles
    With Create
        .Year          = Rng.Offset(00, 0)
        .Month         = Rng.Offset(01, 0)
        .Day           = Rng.Offset(02, 0)
        .Hour          = Rng.Offset(03, 0)
        .Minute        = Rng.Offset(04, 0)
        .Folder        = Rng.Offset(05, 0)
        .ServerStarter = Rng.Offset(06, 0)
        .Maps          = CreateMaps(WB, TileDefinitions)
    End With
End Function

Public Sub ClearPlayers()
    Dim i As Long
    For i = 0 To Maps.Count
        Call Map(i).ClearPlayers
    Next i
End Sub

Public Sub SetSpawnPoints(ByRef Players() As IPlayer)
    Dim i As Long
    For i = 0 To Maps.Count
        Call Map(i).SetSpawnPoints(Players)
    Next i
End Sub

Public Sub UpdateTime()
    Static LastTime As Single
    If LastTime = 0 Then LastTime = Timer

    Dim Diff As Long
    Diff = CLng(Timer - LastTime)

    Year.Value   = DecrementTime(Diff, (31536000 / 60))
    Month.Value  = DecrementTime(Diff, (2628000 / 60))
    Day.Value    = DecrementTime(Diff, (86400 / 60))
    Hour.Value   = DecrementTime(Diff, (3600 / 60))
    Minute.Value = DecrementTime(Diff, (60 / 60))
End Sub


Private Function DecrementTime(ByRef CurrentTime As Long, ByVal Amount As Long) As Long
    Dim Times As Long
    Times = Int(CurrentTime / Amount)
    DecrementTime = Times
    CurrentTime = CurrentTime - (Amount * Times)
End Function

Private Function CreateMaps(ByVal WB As Workbook, ByVal TileDefinitions As PropCollection) As PropCollection
    Dim i As Long

    Dim Rng       As IRange : Set Rng = GameMapsStart
    '2 because -1 for vba offset and the other -1 because Range includes titles, which are not part of the item
    Dim Count     As Long   : Let Count = Rng.GetSelf.Parent.RowCount - 2
    Dim TempArr() As Object : ReDim TempArr(Count)
    With Rng
        For i = 0 To Count
            Set TempArr(i) = GameMap.Create(WB, .Offset(i, 0), TileDefinitions)
        Next i
    End With
    Set CreateMaps = PropCollection.Create(TempArr)
End Function