VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MoveBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit


Private p_Money             As Range
Private p_Row               As Range
Private p_Column            As Range
Private p_SpawnRow          As Range
Private p_SpawnColumn       As Range
Private p_LookDirection     As Range

Public Function Offset() As Long
    Offset = 6
End Function


Public Property Let Money(ByVal n_Money                   As Range)              : Set p_Money           = n_Money             : End Property
Public Property Let Row(ByVal n_Row                       As Range)              : Set p_Row             = n_Row               : End Property
Public Property Let Column(ByVal n_Column                 As Range)              : Set p_Column          = n_Column            : End Property
Public Property Let SpawnRow(ByVal n_SpawnRow             As Range)              : Set p_SpawnRow        = n_SpawnRow          : End Property
Public Property Let SpawnColumn(ByVal n_SpawnColumn       As Range)              : Set p_SpawnColumn     = n_SpawnColumn       : End Property
Public Property Let LookDirection(ByVal n_LookDirection   As Range)              : Set p_LookDirection   = n_LookDirection     : End Property

Public Property Get Money()                               As Range               : Set Money             = p_Money             : End Property
Public Property Get Row()                                 As Range               : Set Row               = p_Row               : End Property
Public Property Get Column()                              As Range               : Set Column            = p_Column            : End Property
Public Property Get SpawnRow()                            As Range               : Set SpawnRow          = p_SpawnRow          : End Property
Public Property Get SpawnColumn()                         As Range               : Set SpawnColumn       = p_SpawnColumn       : End Property
Public Property Get LookDirection()                       As Range               : Set LookDirection     = p_LookDirection     : End Property

Public Function Create(ByVal Rng As Range) As MoveBase
    Set Create = New MoveBase
    With Create
        .Money         = Rng.Offset(0, 0)
        .Row           = Rng.Offset(0, 1)
        .Column        = Rng.Offset(0, 2)
        .SpawnRow      = Rng.Offset(0, 3)
        .SpawnColumn   = Rng.Offset(0, 4)
        .LookDirection = Rng.Offset(0, 5)
        .Row.Value     = .SpawnRow.Value
        .Column.Value  = .SpawnColumn.Value
    End With
End Function

Public Sub Teleport(ByVal Number As Long, ByVal x As Long, ByVal y As Long)
    Dim Seed As Date
    Dim RandomNumber As Double
    Seed = Date()
    RandomNumber = Rnd(Seed)
    With MeServer.GameMap.Tiles
        If Traverseable(y, x) Then
            If .Tile(y, x).Tile.Speed >= RandomNumber Then ' If random value is smaller than chance to go to new tile [1>=0,1 = guaranteed]
                Call .Tile(Row.Value, Column.Value).SetMapPointPlayer(-1)
                Call .Tile(y, x).SetMapPointPlayer(Number)
                Call Look(DirectionByMovement(x - Column.Value, y - Row.Value))
                Row.Value    = y
                Column.Value = x
                Call TriggerEvent()
            End If
        Else
            ' Soundtrack for "cannot enter tile"
        End If
    End With
End Sub

Public Sub Move(ByVal Number As Long, ByVal x As Long, ByVal y As Long)
    Call Teleport(Number, Column.Value + x, Row.Value + y)
End Sub

Public Sub MovePath(ByVal Number As Long, ByRef x() As Long, ByRef y() As Long)
    Dim i As Long
    For i = 0 To ArraySize(x)
        Call Move(x(i), y(i))
    Next i
End Sub

Public Sub Look(ByVal Direction As XlDirection)
    LookDirection.Value = Direction
End Sub

Public Sub Interact(ByVal MyPlayer As IPlayer, ByVal Offset As Long)
    Dim TileRange As Tile
    Set TileRange = GetTileRange(Offset)
    Dim Player As IPlayer

    Set Player = TileRange.Player
    If IsSomething(Player) Then
        If TypeName(Player) = "ComPlayer" Then
            Dim Temp As ComPlayer
            Set Temp = Player
            Dim Args() As Variant
            Args = InterpretArguments(MyPlayer, TileRange, Temp.Script)
            Call Temp.Script.RunArr(Args)
        End If
    End If
    Dim TempFighter As IFighter
    Set TempFighter = MyPlayer
    Call TempFighter.FightBase.LetCurrentMove(FightMove.FightMoveNothing)
    Call TempFighter.FightBase.LetCurrentValue(Nothing)
End Sub

Public Function InFront(ByVal Offset As Long) As Tile
    Select Case LookDirection
        Case xlUp    : InFront = MeServer.GameMap.Tiles.Tile(Row.Value - Offset, Column.Value)
        Case xlDown  : InFront = MeServer.GameMap.Tiles.Tile(Row.Value + Offset, Column.Value)
        Case xlLeft  : InFront = MeServer.GameMap.Tiles.Tile(Row.Value         , Column.Value - Offset)
        Case xlRight : InFront = MeServer.GameMap.Tiles.Tile(Row.Value         , Column.Value + Offset)
End Function

Public Function SubTextureName(ByVal Name As String) As String
    Select Case LookDirection.Value
        Case xlUp    : SubTextureName = Name & "Up"
        Case xlLeft  : SubTextureName = Name & "Left"
        Case xlDown  : SubTextureName = Name & "Down"
        Case xlRight : SubTextureName = Name & "Right"
    End Select
End Function



Private Function InBounds(ByVal Y As Long, ByVal X As Long) As Boolean
    InBounds = Y >= 0                                And _
               Y =< MeServer.GameMap.Rows.Value      And _
               X >= 0                                And _
               X =< MeServer.GameMap.Columns.Value
End Function

Private Function Traverseable(ByVal Y As Long, ByVal X As Long) As Boolean
    If InBounds(Y, X) Then
        Dim nTile As Tile
        Set nTile = MeServer.GameMap.Tiles.Tile(Y, X)
        If nTile.Tile.Speed = 0        Then Exit Function
        If IsSomething(nTile.Player)   Then Exit Function
        Traverseable = True
    End If
End Function

Private Function DirectionByMovement(ByVal x As Long, ByVal y As Long) As XlDirection
    Select Case True
        Case x >= +1 And y =  0  : DirectionByMovement = xlRight
        Case x =< -1 And y =  0  : DirectionByMovement = xlLeft
        Case x =  +0 And y >= +1 : DirectionByMovement = xlDown
        Case x =  +0 And y =< -1 : DirectionByMovement = xlUp
    End Select
End Function

Private Sub TriggerEvent()
    Dim TileScript As Script
    Set TileScript = GetTileRange(0).Script
    If IsSomething(TileScript) Then
        If TileScript.Active = 0 Then
            Dim Temp() As Variant
            Call TileScript.Run(Temp)
        End If
    End If
End Sub

Private Function InterpretArguments(ByVal MyPlayer As IPlayer, ByVal TileRange As Tile, ByVal nScript As Script) As Variant()
    Dim Com       As ComPlayer : Set Com = TileRange.Player
    Dim Def()     As String    : Def = nScript.ArgDef
    Dim Types()   As String    : Types = Split(Com.ScriptValues, ", ")
    Dim ReturnArr() As Variant
    Dim i As Long
    For i = 0 To USize(Def)
        Call VBGLAdd(ReturnArr, InterpretArgument(MyPlayer, TileRange, Types(i), Def(i)))
    Next i
    InterpretArguments = ReturnArr
End Function

Private Function InterpretArgument(ByVal MyPlayer As IPlayer, ByVal TileRange As Tile, ByVal Value As String, ByVal Goal As String) As Variant
    Select Case Value
        Case "ME"
            Set InterpretArgument = MyPlayer
        Case "YOU"
            Select Case Goal
                Case "IPlayer"
                    Dim TempPlayer As IPlayer
                    Set TempPlayer = TileRange.Player
                    Set InterpretArgument = TempPlayer
                Case "IFighter"
                    Dim TempFighter As IFighter
                    Set TempFighter = TileRange.Player
                    Set InterpretArgument = TempFighter
            End Select
        Case Else
            Select Case Goal
                Case "Long"   : InterpretArgument = Clng(Value)
                Case "String" : InterpretArgument = CStr(Value)
            End Select
    End Select
End Function

Private Function GetTileRange(ByVal Offset As Long) As Tile
    Dim x As Long, y As Long
    x = Column.Value
    y = Row.Value
    Select Case LookDirection.Value
        Case xlUp    : Set GetTileRange = MeServer.GameMap.Tiles.Tile(y - Offset, x)
        Case xlDown  : Set GetTileRange = MeServer.GameMap.Tiles.Tile(y + Offset, x)
        Case xlLeft  : Set GetTileRange = MeServer.GameMap.Tiles.Tile(y, x - Offset)
        Case xlRight : Set GetTileRange = MeServer.GameMap.Tiles.Tile(y, x + Offset)
    End Select
End Function