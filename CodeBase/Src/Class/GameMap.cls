VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Name         As Range
Private p_Rows         As Range
Private p_Columns      As Range
Private p_Year         As Range
Private p_Month        As Range
Private p_Day          As Range
Private p_Hour         As Range
Private p_Minute       As Range
Private p_Folder       As Range
Private p_SpawnRow     As Range
Private p_SpawnColumn  As Range
Private p_ServerStarter As Range
Private p_MapPointer   As Range
Private p_Tiles        As Tiles



Public Property Let Name(n_Name                   As Range)   : Set p_Name          = n_Name           : End Property
Public Property Let Rows(n_Rows                   As Range)   : Set p_Rows          = n_Rows           : End Property
Public Property Let Columns(n_Columns             As Range)   : Set p_Columns       = n_Columns        : End Property
Public Property Let Year(n_Year                   As Range)   : Set p_Year          = n_Year           : End Property
Public Property Let Month(n_Month                 As Range)   : Set p_Month         = n_Month          : End Property
Public Property Let Day(n_Day                     As Range)   : Set p_Day           = n_Day            : End Property
Public Property Let Hour(n_Hour                   As Range)   : Set p_Hour          = n_Hour           : End Property
Public Property Let Minute(n_Minute               As Range)   : Set p_Minute        = n_Minute         : End Property
Public Property Let Folder(n_Folder               As Range)   : Set p_Folder        = n_Folder         : End Property
Public Property Let SpawnRow(n_SpawnRow           As Range)   : Set p_SpawnRow      = n_SpawnRow       : End Property
Public Property Let SpawnColumn(n_SpawnColumn     As Range)   : Set p_SpawnColumn   = n_SpawnColumn    : End Property
Public Property Let ServerStarter(n_ServerStarter As Range)   : Set p_ServerStarter = n_ServerStarter  : End Property
Public Property Let MapPointer(n_MapPointer       As Range)   : Set p_MapPointer    = n_MapPointer     : End Property
Public Property Let Tiles(n_Tiles                 As Tiles)   : Set p_Tiles         = n_Tiles          : End Property

Public Property Get Name()                        As Range    : Set Name            = p_Name           : End Property
Public Property Get Rows()                        As Range    : Set Rows            = p_Rows           : End Property
Public Property Get Columns()                     As Range    : Set Columns         = p_Columns        : End Property
Public Property Get Year()                        As Range    : Set Year            = p_Year           : End Property
Public Property Get Month()                       As Range    : Set Month           = p_Month          : End Property
Public Property Get Day()                         As Range    : Set Day             = p_Day            : End Property
Public Property Get Hour()                        As Range    : Set Hour            = p_Hour           : End Property
Public Property Get Minute()                      As Range    : Set Minute          = p_Minute         : End Property
Public Property Get Folder()                      As Range    : Set Folder          = p_Folder         : End Property
Public Property Get SpawnRow()                    As Range    : Set SpawnRow        = p_SpawnRow       : End Property
Public Property Get SpawnColumn()                 As Range    : Set SpawnColumn     = p_SpawnColumn    : End Property
Public Property Get ServerStarter()               As Range    : Set ServerStarter   = p_ServerStarter  : End Property
Public Property Get MapPointer()                  As Range    : Set MapPointer      = p_MapPointer     : End Property
Public Property Get Tiles()                       As Tiles    : Set Tiles           = p_Tiles          : End Property

Public Function Create(Map As Range, Rng As Range) As GameMap
    Set Create = New GameMap
    With Create
        .Name          = Rng.Offset(00, 0)
        .Rows          = Rng.Offset(01, 0)
        .Columns       = Rng.Offset(02, 0)
        .Year          = Rng.Offset(03, 0)
        .Month         = Rng.Offset(04, 0)
        .Day           = Rng.Offset(05, 0)
        .Hour          = Rng.Offset(06, 0)
        .Minute        = Rng.Offset(07, 0)
        .Folder        = Rng.Offset(08, 0)
        .SpawnRow      = Rng.Offset(09, 0)
        .SpawnColumn   = Rng.Offset(10, 0)
        .ServerStarter = Rng.Offset(11, 0)
        .MapPointer    = Map
        .Tiles         = FactoryTiles.Create(Map, CLng(.Rows.Value), CLng(.Columns.Value))
    End With
End Function

Public Function InBounds(ByVal Y As Long, ByVal X As Long) As Boolean
    InBounds = Y >= 0              And _
               Y =< Rows.Value     And _
               X >= 0              And _
               X =< Columns.Value
End Function

Public Function Traverseable(ByVal Y As Long, ByVal X As Long) As Boolean
    If InBounds(Y, X) Then
        Dim nTile As Tile
        Set nTile = Tiles.Tile(Y, X)
        If nTile.Tile.Speed = 0        Then Exit Function
        If IsSomething(nTile.Player)   Then Exit Function
        Traverseable = True
    End If
End Function

Public Sub ClearPlayers()
    Dim y As Long, x As Long
    For y = 0 To Rows.Value
        For x = 0 To Columns.Value
            Tiles.Tile(y, x).SetMapPointPlayer(-1)
        Next x
    Next y
End Sub

Public Sub SetSpawnPoints(ByRef Players() As IPlayer)
    Dim y As Long, x As Long
    Dim i As Long
    For i = 0 To USize(Players)
        If TypeName(Players(i)) = "ComPlayer" Then
            Dim Temp As ComPlayer
            Set Temp = Players(i)
            y = Temp.SpawnRow.Value
            x = Temp.SpawnColumn.Value
        Else
            Dim Temp2 As HumanPlayer
            Set Temp2 = Players(i)
            y = Temp.Row.Value
            x = Temp.Column.Value
        End If
        Tiles.Tile(y, x).SetMapPointPlayer(Players(i).Number.Value)
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