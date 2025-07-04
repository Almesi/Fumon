VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HumanPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IPlayer


Private p_Number            As Range
Private p_Name              As Range
Private p_Row               As Range
Private p_Column            As Range
Private p_LookDirection     As Range
Private p_Money             As Range
Private p_Fumons            As Fumons
Private p_Items             As Items
Private p_CurrentValue      As Variant
Private p_CurrentMove       As FightMove
Private p_TextureDefinition As TextureDefinition
Private p_Beaten()          As IPlayer
Private p_Quests()          As Quest





Private Property Let IPlayer_Number(n_Number                       As Range)              : Set p_Number                  = n_Number            : End Property
Private Property Let IPlayer_Name(n_Name                           As Range)              : Set p_Name                    = n_Name              : End Property
Private Property Let IPlayer_Row(n_Row                             As Range)              : Set p_Row                     = n_Row               : End Property
Private Property Let IPlayer_Column(n_Column                       As Range)              : Set p_Column                  = n_Column            : End Property
Private Property Let IPlayer_LookDirection(n_LookDirection         As Range)              : Set p_LookDirection           = n_LookDirection     : End Property
Private Property Let IPlayer_Money(n_Money                         As Range)              : Set p_Money                   = n_Money             : End Property
Private Property Let IPlayer_Fumons(n_Fumons                       As Fumons)             : Set p_Fumons                  = n_Fumons            : End Property
Private Property Let IPlayer_Items(n_Items                         As Items)              : Set p_Items                   = n_Items             : End Property
Private Property Let IPlayer_CurrentValue(n_CurrentValue           As Variant)            : Set p_CurrentValue            = n_CurrentValue      : End Property
Private Property Let IPlayer_CurrentMove(n_CurrentMove             As FightMove)          : Set p_CurrentMove             = n_CurrentMove       : End Property
Private Property Let IPlayer_TextureDefinition(n_TextureDefinition As TextureDefinition)  : Set p_TextureDefinition       = n_TextureDefinition : End Property

Private Property Get IPlayer_Number()                              As Range               : Set IPlayer_Number            = p_Number            : End Property
Private Property Get IPlayer_Name()                                As Range               : Set IPlayer_Name              = p_Name              : End Property
Private Property Get IPlayer_Row()                                 As Range               : Set IPlayer_Row               = p_Row               : End Property
Private Property Get IPlayer_Column()                              As Range               : Set IPlayer_Column            = p_Column            : End Property
Private Property Get IPlayer_LookDirection()                       As Range               : Set IPlayer_LookDirection     = p_LookDirection     : End Property
Private Property Get IPlayer_Money()                               As Range               : Set IPlayer_Money             = p_Money             : End Property
Private Property Get IPlayer_Fumons()                              As Fumons              : Set IPlayer_Fumons            = p_Fumons            : End Property
Private Property Get IPlayer_Items()                               As Items               : Set IPlayer_Items             = p_Items             : End Property
Private Property Get IPlayer_CurrentValue()                        As Variant             : Set IPlayer_CurrentValue      = p_CurrentValue      : End Property
Private Property Get IPlayer_CurrentMove()                         As FightMove           : Set IPlayer_CurrentMove       = p_CurrentMove       : End Property
Private Property Get IPlayer_TextureDefinition()                   As TextureDefinition   : Set IPlayer_TextureDefinition = p_TextureDefinition : End Property

Private Function IPlayer_Create(Pointer As Range) As IPlayer
    Dim i As Long
    Set IPlayer_Create = New HumanPlayer
    With IPlayer_Create
        .Name          = Pointer.Offset(1, 1)
        .Row           = Pointer.Offset(2, 1)
        .Column        = Pointer.Offset(3, 1)
        .LookDirection = Pointer.Offset(4, 1)
        .Money         = Pointer.Offset(5, 1)
        .Fumons        = FumonGame.Fumons.Create(Pointer.Offset(6, 1))
        .Items         = FumonGame.Items.Create(Pointer.Offset(7, 1))
        i = 0
        Do Until Pointer.Offset(8, 1 + i).Formula = Empty
            .AddBeaten(Pointer.Offset(8, 1 + i).Value)
            i = i + 1
        Loop
        i = 0
        Do Until Pointer.Offset(9, 1 + i).Formula = Empty
            .AddQuest(Pointer.Offset(9, 1 + i).Value)
            i = i + 1
        Loop
        .TextureDefinition = FumonGame.TextureDefinition.Create(Pointer.Offset(10, 1).Value, Pointer.Offset(11, 1), FumonGame.FactoryServer.GameMap.Folder.Value & "\Player.png")
    End With
End Function

Private Sub IPlayer_Move(x As Long, y As Long)
    Dim Seed As Date
    Dim RandomNumber As Double
    Dim OffsetY As Long: OffsetY = IPlayer_Row.Value    + y
    Dim OffsetX As Long: OffsetX = IPlayer_Column.Value + x

    Seed = Date()
    RandomNumber = Rnd(Seed)
    If Traverseable(OffsetY, OffsetX) Then
        If MeServer.GameMap.GetTile(OffsetY, OffsetX).Speed >= RandomNumber Then ' If random value is smaller than chance to go to new tile [1>=0,1 = guaranteed]
            IPlayer_Row.Value    = OffsetY
            IPlayer_Column.Value = OffsetX
        End If
    Else
        ' Soundtrack for "cannot enter tile"
    End If
End Sub

Private Sub IPlayer_MovePath(x() As Long, y() As Long, LookDirection() As Long)
    Dim i As Long
    For i = 0 To ArraySize(x)
        Call IPlayer_Look(LookDirection(i))
        Call IPlayer_Move(x(i), y(i))
    Next i
End Sub

Private Sub IPlayer_Look(Direction As Long)
    IPlayer_LookDirection.Value = Direction
End Sub

Private Sub IPlayer_Interact(TileRange As Range)
    If TileRange Is Nothing Then ' Cannot interact with range, as it doesnt exist(out of bounds)
    Else

    End If
End Sub

Private Sub IPlayer_DoAI(OtherPlayer As IPlayer)
    Dim MyFight As Fight
    Set MyFight = GetCurrentFight(OtherPlayer)
    If MyFight.WaitTime(Me, 60) = False Then ' If my time runs out
        Call MyFight.DoAction(CurrentValue, CurrentMove)
    End If
    Call MyFight.ChangePlayer()
End Sub

Private Sub IPlayer_AddBeaten(Rng As Range)
    Call ArrayPush(p_Beaten, GameServer.GetPlayer(Rng.Value))
    Rng.Value = GameServer.GetPlayer(Rng.Value).Number
End Sub

Private Sub IPlayer_AddQuest(NewQuest As Quest)
    Call ArrayPush(p_Quests, NewQuest)
End Sub

Private Function IPlayer_PlayerDefeated(Index As Long) As Boolean
    Dim i As Long
    For i = 0 To ArraySize(p_Beaten)
        If p_Beaten.Number = Index Then
            PlayerDefeated = True
            Exit Function
        End If
    Next i
End Function

Public Function TileInFront(Offset As Long) As Range
    On Error GoTo Error
    Select Case LookDirection.Value
        Case = 0: Set TileInFront = MeServer.GameMap.GetTile(Row.Value - Offset, Column.Value + 0)      ' Up
        Case = 1: Set TileInFront = MeServer.GameMap.GetTile(Row.Value + 0     , Column.Value - Offset) ' Left
        Case = 2: Set TileInFront = MeServer.GameMap.GetTile(Row.Value + Offset, Column.Value + 0)      ' Down
        Case = 3: Set TileInFront = MeServer.GameMap.GetTile(Row.Value + 0     , Column.Value + Offset) ' Right
    End Select
End Function

Public Function GetCurrentFight(OtherPlayer As IPlayer) As Fight
    Dim i As Long
    Dim p1 As Long
    Dim p2 As Long
    Dim MyNumber As Long
    Dim EnemyNumber As Long
    
    For i = 0 To Ubound(MeServer.Fights)
        p1 = MeServer.Fight(i).p1.Value
        p2 = MeServer.Fight(i).p2.Value
        MyNumber = Me.Number.Value
        EnemyNumber = OtherPlayer.Number.Value
        If (MyNumber = p1 And EnemyNumber = p2) Or (MyNumber = p2 And EnemyNumber = p1) Then
            Set GetCurrentFight = MeServer.Fight(i)
            Exit Function
        End If
    Next i
End Function

Private Function InBounds(Y As Long, X As Long) As Boolean
    InBounds = Y >= 0                                And _
               Y =< MeServer.GameMap.Rows.Value      And _
               X >= 0                                And _
               X =< MeServer.GameMap.Columns.Value
End Function

Private Function Traverseable(Y As Long, X As Long) As Boolean
    If InBounds(Y, X) Then
        If MeServer.GameMap.GetTile(Y, X).Speed          = +0 Then Exit Function
        If MeServer.GameMap.GetPlayer(Y, X).Number.Value = -1 Then Exit Function
    End If
    Traverseable = True
End Function