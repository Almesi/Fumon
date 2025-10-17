Attribute VB_Name = "Scripts"


Option Explicit



Public Sub Say(ByVal Name As String, ByVal Message As String) 
    Dim PreviousInput As VBGLIInput
    Call UpdateMessage(Name, Message)
    Set PreviousInput = CurrentRenderObject.Inputt
    CurrentRenderObject.Inputt = MessageBoxInput
    Call CurrentRenderObject.AddDrawable(MessageBox)
    Do Until EscapeTextBox = True
        Call glutMainLoopEvent()
        Call CurrentRenderObject.Loopp
    Loop
    CurrentRenderObject.Inputt = PreviousInput
    Call CurrentRenderObject.RemoveDrawable()
    Call EscapeTextBox(True)
End Sub

Public Function EscapeTextBox(Optional ByVal Setter As Boolean = False) As Boolean
    Static Value As Boolean
    If Setter Then Value = Value Xor True
    EscapeTextBox = Value
End Function

Public Sub IncrementItem(ByVal Player As HumanPlayer, ByVal ItemPointer As Long, ByVal Value As Long)
    Dim Itemm As Item: Set Itemm = GetItem(Player, ItemPointer)
    Itemm.Value = Itemm.Value + Value
End Sub

Public Function GetItem(ByVal Player As HumanPlayer, ByVal ItemPointer As Long) As Item
    Set GetItem = Player.Items.Item(ItemPointer)
End Function

Public Function GetItemAmount(ByVal Player As HumanPlayer, ByVal ItemPointer As Long) As Long
    GetItemAmount = GetItem(Player, ItemPointer).Amount.Value
End Function

Public Function GetPlayer(ByVal Index As Long) As IPlayer
    Set GetPlayer = MeServer.Player(Index)
End Function

Public Function MakeArgumentArr(ByVal Text As String) As Variant()
    Dim i As Long
    Dim Arr() As String
    Dim ReturnArr() As Variant

    Arr = Split(Text, ", ")
    ReDim ReturnArr(USize(Arr))
    For i = 0 To USize(Arr)
        ReturnArr(i) = InterpretArgument(Arr(i))
    Next i
    MakeArgumentArr = ReturnArr
End Function

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

Public Function TilesInFront(ByVal Index As Long, ByVal Offset As Long) As Range
    Call MeServer.Player(Index).TileInFront(Offset)
End Function