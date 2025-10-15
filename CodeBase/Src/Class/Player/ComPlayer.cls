VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IPlayer


Private p_Number            As Range
Private p_Name              As Range
Private p_Money             As Range
Private p_Script            As Script
Private p_ScriptValues      As String
Private p_Row               As Range
Private p_Column            As Range
Private p_SpawnRow          As Range
Private p_SpawnColumn       As Range
Private p_LookDirection     As Range
Private p_CurrentMove       As Range
Private p_CurrentValue      As Range
Private p_Fumons            As Fumons
Private p_Items             As Items
Private p_SubTexture        As VBGLSubTexture


Public Property Let Number(ByVal n_Number                 As Range)              : Set p_Number          = n_Number            : End Property
Public Property Let Name(ByVal n_Name                     As Range)              : Set p_Name            = n_Name              : End Property
Public Property Let Money(ByVal n_Money                   As Range)              : Set p_Money           = n_Money             : End Property
Public Property Let Script(ByVal n_Script                 As Script)             : Set p_Script          = n_Script            : End Property
Public Property Let ScriptValues(ByVal n_ScriptValues     As String)             : Let p_ScriptValues    = n_ScriptValues      : End Property
Public Property Let Row(ByVal n_Row                       As Range)              : Set p_Row             = n_Row               : End Property
Public Property Let Column(ByVal n_Column                 As Range)              : Set p_Column          = n_Column            : End Property
Public Property Let SpawnRow(ByVal n_SpawnRow             As Range)              : Set p_SpawnRow        = n_SpawnRow          : End Property
Public Property Let SpawnColumn(ByVal n_SpawnColumn       As Range)              : Set p_SpawnColumn     = n_SpawnColumn       : End Property
Public Property Let LookDirection(ByVal n_LookDirection   As Range)              : Set p_LookDirection   = n_LookDirection     : End Property
Public Property Let CurrentMove(ByVal n_CurrentMove       As Range)              : Set p_CurrentMove     = n_CurrentMove       : End Property
Public Property Let CurrentValue(ByVal n_CurrentValue     As Range)              : Set p_CurrentValue    = n_CurrentValue      : End Property
Public Property Let Fumons(ByVal n_Fumons                 As Fumons)             : Set p_Fumons          = n_Fumons            : End Property
Public Property Let Items(ByVal n_Items                   As Items)              : Set p_Items           = n_Items             : End Property
Public Property Let SubTexture(ByVal n_SubTexture         As VBGLSubTexture)     : Set p_SubTexture      = n_SubTexture        : End Property

Public Property Get Number()                              As Range               : Set Number            = p_Number            : End Property
Public Property Get Name()                                As Range               : Set Name              = p_Name              : End Property
Public Property Get Money()                               As Range               : Set Money             = p_Money             : End Property
Public Property Get Script()                              As Script              : Set Script            = p_Script            : End Property
Public Property Get ScriptValues()                        As String              : Let ScriptValues      = p_ScriptValues      : End Property
Public Property Get Row()                                 As Range               : Set Row               = p_Row               : End Property
Public Property Get Column()                              As Range               : Set Column            = p_Column            : End Property
Public Property Get SpawnRow()                            As Range               : Set SpawnRow          = p_SpawnRow          : End Property
Public Property Get SpawnColumn()                         As Range               : Set SpawnColumn       = p_SpawnColumn       : End Property
Public Property Get LookDirection()                       As Range               : Set LookDirection     = p_LookDirection     : End Property
Public Property Get CurrentMove()                         As Range               : Set CurrentMove       = p_CurrentMove       : End Property
Public Property Get CurrentValue()                        As Range               : Set CurrentValue      = p_CurrentValue      : End Property
Public Property Get Fumons()                              As Fumons              : Set Fumons            = p_Fumons            : End Property
Public Property Get Items()                               As Items               : Set Items             = p_Items             : End Property
Public Property Get SubTexture()                          As VBGLSubTexture      : Set SubTexture        = p_SubTexture        : End Property


Private Property Let IPlayer_Number(ByVal n_Number                As Range)              : Set p_Number             = n_Number            : End Property
Private Property Let IPlayer_Name(ByVal n_Name                    As Range)              : Set p_Name               = n_Name              : End Property
Private Property Let IPlayer_Fumons(ByVal n_Fumons                As Fumons)             : Set p_Fumons             = n_Fumons            : End Property
Private Property Let IPlayer_Items(ByVal n_Items                  As Items)              : Set p_Items              = n_Items             : End Property
Private Property Let IPlayer_CurrentMove(ByVal n_CurrentMove      As FightMove)          : Let p_CurrentMove        = n_CurrentMove       : End Property
Private Property Let IPlayer_CurrentValue(ByVal n_CurrentValue    As Variant)            : Set p_CurrentValue       = n_CurrentValue      : End Property

Private Property Get IPlayer_Number()                             As Range               : Set IPlayer_Number       = p_Number            : End Property
Private Property Get IPlayer_Name()                               As Range               : Set IPlayer_Name         = p_Name              : End Property
Private Property Get IPlayer_Fumons()                             As Fumons              : Set IPlayer_Fumons       = p_Fumons            : End Property
Private Property Get IPlayer_Items()                              As Items               : Set IPlayer_Items        = p_Items             : End Property
Private Property Get IPlayer_CurrentMove()                        As FightMove           : Let IPlayer_CurrentMove  = p_CurrentMove       : End Property
Private Property Get IPlayer_CurrentValue()                       As Variant             : Set IPlayer_CurrentValue = p_CurrentValue      : End Property


Private Sub IPlayer_LetCurrentMove(ByVal n_CurrentMove As FightMove)
    CurrentMove.Value = n_CurrentMove
End Sub
Private Sub IPlayer_LetCurrentValue(ByVal n_CurrentValue As Variant)
    CurrentValue.Value = n_CurrentValue
End Sub

Private Function IPlayer_GetCurrentMove() As FightMove
    IPlayer_GetCurrentMove = CurrentMove.Value
End Function
Private Function IPlayer_GetCurrentValue(ByVal MyFight As Fight) As Variant
    Select Case IPlayer_GetCurrentMove
        Case FightMove.FightMoveAttack      : Set IPlayer_GetCurrentValue = MyFight.MyFumon(Me).Attacks.Attack(CurrentValue.Value)
        Case FightMove.FightMoveChangeFumon : Set IPlayer_GetCurrentValue = IPlayer_Fumons.Fumon(CurrentValue.Value)
        Case FightMove.FightMoveItem        : Set IPlayer_GetCurrentValue = IPlayer_Items.Item(CurrentValue.Value)
        Case FightMove.FightMoveFlee        : Set IPlayer_GetCurrentValue = Nothing
        Case FightMove.FightMoveNothing     : Set IPlayer_GetCurrentValue = Nothing
    End Select
End Function



Private Sub IPlayer_DoAI(ByVal MyFight As Fight, ByVal OtherPlayer As IPlayer)
    Call IPlayer_LetCurrentValue(-1)
    Call IPlayer_LetCurrentMove(FightMove.FightMoveNothing)
    Call ComAI(OtherPlayer)
End Sub


Public Function Create(ByVal Pointer As Range) As ComPlayer
    Set Create = New ComPlayer
    With Create
        .Number            = Pointer.Offset(0, 0)
        .Name              = Pointer.Offset(0, 1)
        .Money             = Pointer.Offset(0, 2)
        .Script            = FactoryServer.GetScript(Pointer.Offset(0, 3).Value)
        .ScriptValues      = Pointer.Offset(0, 4).Value
        .Row               = Pointer.Offset(0, 7)
        .Column            = Pointer.Offset(0, 8)
        .SpawnRow          = Pointer.Offset(0, 9)
        .SpawnColumn       = Pointer.Offset(0, 10)
        .LookDirection     = Pointer.Offset(0, 11)
        .CurrentMove       = Pointer.Offset(0, 12)
        .CurrentValue      = Pointer.Offset(0, 13)
        .Fumons            = FactoryFumons.Create(Pointer.Offset(0, 14))
        .Items             = FactoryItems.Create(Pointer.Offset(0, 112))
        .SubTexture        = FactoryServer.GetTexture("Players").SubTextureID(.SubTextureName())
        .Row.Value         = .SpawnRow.Value
        .Column.Value      = .SpawnColumn.Value
    End With
End Function

Public Sub Move(ByVal x As Long, ByVal y As Long)
    Dim Seed As Date
    Dim RandomNumber As Double
    Dim OffsetY As Long: OffsetY = Row.Value    + y
    Dim OffsetX As Long: OffsetX = Column.Value + x
    Seed = Date()
    RandomNumber = Rnd(Seed)
    With MeServer.GameMap.Tiles
        If Traverseable(OffsetY, OffsetX) Then
            If .Tile(OffsetY, OffsetX).Tile.Speed >= RandomNumber Then ' If random value is smaller than chance to go to new tile [1>=0,1 = guaranteed]
                Call .Tile(Row.Value, Column.Value).SetMapPointPlayer(-1)
                Call .Tile(OffsetY, OffsetX).SetMapPointPlayer(Number.Value)
                Row.Value    = OffsetY
                Column.Value = OffsetX
                Select Case True
                    Case x = -1 And y = 0  : LookDirection.Value = xlLeft
                    Case x = +1 And y = 0  : LookDirection.Value = xlRight
                    Case x = 0  And y = -1 : LookDirection.Value = xlUp
                    Case x = 0  And y = +1 : LookDirection.Value = xlDown
                End Select
                Call TriggerEvent()
            End If
        Else
            ' Soundtrack for "cannot enter tile"
        End If
    End With
End Sub

Public Sub MovePath(ByRef x() As Long, ByRef y() As Long, ByRef LookDirection() As XlDirection)
    Dim i As Long
    For i = 0 To ArraySize(x)
        Call Look(LookDirection(i))
        Call Move(x(i), y(i))
    Next i
End Sub

Public Sub Look(ByVal Direction As XlDirection)
    LookDirection.Value = Direction
End Sub

Public Sub Interact(ByVal Offset As Long)
    Dim TileRange As Tile
    Set TileRange = GetTileRange(Offset)

    If IsSomething(TileRange.Player) Then
        If IsSomething(TileRange.Player.Script) Then
            Dim Args() As Variant
            Args = InterpretArguments(TileRange, TileRange.Player.Script)
            Call TileRange.Player.Script.Run(Args)
        End If
    End If
End Sub

Public Function InFront(ByVal Offset As Long) As Tile
    Select Case LookDirection
        Case xlUp    : InFront = MeServer.GameMap.Tiles.Tile(Row.Value - Offset, Column.Value)
        Case xlDown  : InFront = MeServer.GameMap.Tiles.Tile(Row.Value + Offset, Column.Value)
        Case xlLeft  : InFront = MeServer.GameMap.Tiles.Tile(Row.Value         , Column.Value - Offset)
        Case xlRight : InFront = MeServer.GameMap.Tiles.Tile(Row.Value         , Column.Value + Offset)
End Function

Public Function SubTextureName() As String
    Select Case LookDirection.Value
        Case xlUp    : SubTextureName = Name.Value & "Up"
        Case xlLeft  : SubTextureName = Name.Value & "Left"
        Case xlDown  : SubTextureName = Name.Value & "Down"
        Case xlRight : SubTextureName = Name.Value & "Right"
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

Private Function InterpretArguments(ByVal TileRange As Tile, ByVal nScript As Script) As Variant()
    Dim Def()     As String : Def = nScript.ArgDef
    Dim Types()   As String : Types = Split(TileRange.Player.ScriptValues, ", ")
    Dim ReturnArr() As Variant
    Dim i As Long
    For i = 0 To USize(Def)
        Call VBGLAdd(ReturnArr, InterpretArgument(TileRange, Types(i), Def(i)))
    Next i
    InterpretArguments = ReturnArr
End Function

Private Function InterpretArgument(ByVal TileRange As Tile, ByVal Value As String, ByVal Goal As String) As Variant
    Select Case Value
        Case "ME"
            Set InterpretArgument = Me
        Case "YOU"
            Select Case Goal
                Case "IPlayer"
                    Set InterpretArgument = TileRange.Player
            End Select
        Case Else
            Select Case Goal
                Case "Long"   : InterpretArgument = Clng(Value)
                Case "String" : InterpretArgument = CStr(Value)
            End Select
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
    End If
    Traverseable = True
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

Private Sub ComAI(ByVal OtherPlayer As IPlayer)
    Call IPlayer_LetCurrentMove(FightMove.FightMoveAttack)
    Call IPlayer_LetCurrentValue(0)
End Sub

Private Sub Class_Initialize()
    Set p_CurrentValue = Nothing
End Sub