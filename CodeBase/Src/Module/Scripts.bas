Attribute VB_Name = "Scripts"


Option Explicit

Public Function CallScript(Text As String)
    Dim Scripts() As String
    Dim Arguments() As String
    Dim ScriptName As String
    Scripts = Split(Text, "; ")
    For i = 0 Ubound(Scripts)
        Arguments = Split(GetParanthesesText(Scripts(i)), ", ")
        ScriptName = GetProcedureName(Scripts(i))
        If IsNumeric(ScriptName) Then
            CallScript = MeServer.Script(CLng(ScriptName)).Run(Arguments)
        End If
    Next i
End Function

Public Function MakeArgumentArr(ParamArray Arguments As Variant)
    Dim i As Long
    Dim Arr() As Variant

    ReDim Arr(Ubound(Arguments))
    For i = 0 To Ubound(Arr)
        Arr(i) = Arguments(i)
    Next i
    MakeArgumentArr = Arr
End Function

Public Sub Say(Index As Long, Text As String)
    MsgBox(Text, , MeServer.Player(Index).Name)
End Sub

Public Sub Fight(Player1 As Index, Player2 As Index)
    Dim CurrentFight As Fight
    Dim Winner       As IPlayer
    Dim Loser As IPlayer

    Set CurrentFight = Fight.Create(Player1, Player2)
    Set Winner = CurrentFight.Fight(MeServer.Player(Player1))

    If Not Winner Is Nothing Then
        Set Loser = CurrentFight.Loser()
        If TypeName(Winner) = "HumanPlayer" And TypeName(Loser) = "ComPlayer" Then
            Call IncrementMoney(Winner.Number.Value, Loser.Money.Value)
        End If
    End If
End Sub

Public Sub IncrementMoney(Index As Long, Value As Long)
    Dim Money As Range: Money = MeServer.Player(Index).Money
    Money.Value = Money.Value = Value
End Sub

Public Sub IncrementItem(Index As Long, ItemPointer As Long, Value As Long)
    Dim Itemm As Item: Set Itemm = MeServer.Player(Index).Items(ItemPointer)
    Itemm.Value = Itemm.Value + Value
End Sub

Public Function GetItem(Index As Long, ItemPointer As Long) As Item
    Set GetItem = MeServer.Player(Index).Items(ItemPointer)
End Function

Public Function GetItemAmount(Index As Long, ItemPointer As Long) As Long
    GetItemAmount = GetItem(Index, ItemPointer).Amount.Value
End Function

Public Sub Move(Index As Long, x As Long, y As Long)
    Call MeServer.Player(Index).Move(x, y)
End Sub

Public Sub MovePath(Index As Long, x() As Long, y() As Long, LookDirection() As Long)
    Call MeServer.Player(Index).MovePath(x, y, LookDirection)
End Sub

Public Sub Look(Index As Long, Direction As Long)
    MeServer.Player(Index).Look.Value = Direction
End Sub

Public Sub UseItem(Itemm As Item)
    Call Itemm.ItemDefinition.Script.Run()
End Sub

Public Function GetPlayerFumon(Index As Long, FumonIndex As Long) As Fumon
    Set GetPlayerFumon = MeServer.Player(Index).Fumons.Fumon(FumonIndex)
End Function

Public Sub HealPlayerFumon(Index As Long, FumonIndex As Long, Value As Long)
    Dim Fumon As Fumon
    Set Fumon = GetPlayerFumon(Index, FumonIndex)
    Fumon.Health = Fumon.Health + Value
    Call Fumon.CheckHealth()
End Sub

Public Function TileInFront(Index As Long, Offset As Long) As Range
    Call MeServer.Player(Index).TileInFront(Offset)
End Function

Public Function DefinePath(ParamArray Values As Variant) As Long()
    Dim i As Long
    Dim Arr() As Long
    ReDim Arr(Ubound(Values))
    For i = 0 To Ubound(Arr)
        Arr(i) = CLng(Values)
    Next i
    DefinePath = Arr
End Function

Public Function FindPath(Player1 As Long, Player2 As Long)
    Dim StartX As Long, EndX As Long
    Dim StartY As Long, EndY As Long

    Dim x() As Long
    Dim y() As Long
    Dim Look() As Long

    StartX = MeServer.Player(Player1).Column.Value
    EndX   = MeServer.Player(Player2).Column.Value
    StartY = MeServer.Player(Player1).Row.Value
    EndY   = MeServer.Player(Player2).Row.Value

    Call FindPathAlgorithm(StartX, EndX, StartY, EndY, x, y, Look)
End Function

Public Function GetCurrentFight(Player1 As Long, Player2 As Long) As Fight
    Set GetCurrentFight = MeServer.Players(Player1).GetCurrentFight(Player2)
End Function

' Arrays have to be from calling procedure. They are the returning values
Public Sub FindPathAlgorithm(StartX As Long, EndX As Long, StartY As Long, EndY As Long, x() As Long, y() As Long, LookDirection() As Long)
    Dim i As Long
    Dim yCount As Long
    Dim Size As Long

    Size = (EndY - StartY) + (EndX - StartX) - 1
    ReDim x(Size)
    ReDim y(Size)
    ReDim LookDirection(Size)

    For i = 0 To Size
        If MeServer.GameMap.GetTile(yCount + StartY, i + StartX) Then
            If EndX - StartX > 0 Then
                x(i) = 1
            ElseIf EndX - StartX < 0 Then
                x(i) = -1
            End If
        End If
        
        If x(i) = 0 Then
            If EndY - StartY > 0 Then
                y(i) = 1
                yCount = yCount + 1
            ElseIf EndY - StartY < 0 Then
                y(i) = -1
                yCount = yCount - 1
            End If
        End If

        Select Case True
            Case x(i) = +0 And y(i) = -1 : LookDirection(i) = 2 ' UP
            Case x(i) = -1 And y(i) = +0 : LookDirection(i) = 3 ' LEFT
            Case x(i) = +0 And y(i) = +1 : LookDirection(i) = 0 ' DOWN
            Case x(i) = +1 And y(i) = +0 : LookDirection(i) = 1 ' RIGHT
        End Select
    Next
End Sub

Public Function GetIndexByName(ObjectName As String, Name As String) As Long
    MeServer.GetIndexByName(ObjectName, Name)
End Function