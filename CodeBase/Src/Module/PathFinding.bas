Attribute VB_Name = "PathFinding"


Option Explicit

Private Type Node
    X         As Long
    Y         As Long
    StartCost As Long ' G
    Distance  As Long ' H
    TotalCost As Long ' F
    ParentX   As Long
    ParentY   As Long
    Closed    As Boolean
    Open      As Boolean
End Type

Public Sub CreateMovePath(ByRef xPoints() As Long, ByRef yPoints() As Long, ByRef x() As Long, ByRef y() As Long)
    Dim TempX() As Long
    Dim TempY() As Long
    Dim StartX As Long, StartY As Long
    Dim EndX   As Long, EndY   As Long
    Dim Index As Long
    Dim i As Long
    Dim Size As Long
    Size = USize(xPoints)
    For i = 0 To Size
        Index = Iif(i + 1 > Size, 0, i + 1)
        StartX = xPoints(i)
        StartY = yPoints(i)
        EndX   = xPoints(Index)
        EndY   = yPoints(Index)
        Call FindPathAlgorithm(StartX, EndX, StartY, EndY, TempX, TempY)
        Call VBGLArrayMerge(x, TempX)
        Call VBGLArrayMerge(y, TempY)
    Next i
End Sub

Public Sub DefinePath(ByRef x() As Long, ByRef y() As Long, ParamArray Points() As Variant)
    Dim Size As Long
    Size = ((USize(Points) + 1) / 2) - 1

    ReDim x(Size)
    ReDim y(Size)

    Dim i As Long
    For i = 0 To Size
        x(i) = Points(i * 2 + 0)
        y(i) = Points(i * 2 + 1)
    Next i
End Sub

Public Function FindPath(ByVal Player1 As Long, ByVal Player2 As Long, ByRef x() As Long, ByRef y() As Long)
    Dim StartX As Long : StartX = MeServer.Player(Player1).Column.Value
    Dim StartY As Long : StartY = MeServer.Player(Player1).Row.Value
    Dim EndX   As Long : EndX   = MeServer.Player(Player2).Column.Value
    Dim EndY   As Long : EndY   = MeServer.Player(Player2).Row.Value - 1 ' will always stop above below player, thats prone to bugs since it doesnt check if that is usable

    Call FindPathAlgorithm(StartX, EndX, StartY, EndY, x, y)
End Function

Public Function ReversePath(ByRef InX() As Long, ByRef InY() As Long, ByRef OutX() As Long, ByRef OutY() As Long)
    Dim i As Long
    OutX = InX
    OutY = InY

    For i = 0 To USize(InX)
        OutX = InX * -1
    Next i

    For i = 0 To USize(InY)
        OutY = InY * -1
    Next i

End Function

Public Sub FindPathAlgorithm(ByVal StartX As Long, ByVal EndX As Long, ByVal StartY As Long, ByVal EndY As Long, ByRef x() As Long, ByRef y() As Long)
    Dim i As Long
    Dim Rows       As Long : Rows    = MeServer.GameMap.Rows.Value
    Dim Columns    As Long : Columns = MeServer.GameMap.Columns.Value
    Dim Grid()     As Node : Grid    = InitializeGrid(Rows, Columns)
    Dim OpenList() As Node : ReDim OpenList(0)
    Dim CountOpen  As Long : CountOpen = 0
    
    ' Add start node
    With Grid(StartY, StartX)
        .Distance = GetDistance(StartX, StartY, EndX, EndY)
        .TotalCost = .Distance
        .Open = True
    End With
    OpenList(0) = Grid(StartY, StartX)
    CountOpen = 1
    
    Do While CountOpen > 0
        Dim BestIndex As Long : BestIndex = FindLowestNode(OpenList, CountOpen)
        Dim Current   As Node : Current = OpenList(BestIndex)
        Call CloseNode(Grid, Current, OpenList, BestIndex, CountOpen)
        Dim Found As Boolean  : Found = FoundGoal(Current, EndX, EndY)
        If Found Then Exit Do
        
        Dim dirX(3) As Long, dirY(3) As Long
        Call Neighbors(Current, dirX, dirY)
        
        For i = 0 To 3
            Dim nx As Long, ny As Long
            nx = dirX(i)
            ny = dirY(i)
            
            If MeServer.GameMap.Traverseable(ny, nx) Then
                Dim Neighbor As Node
                Neighbor = Grid(ny, nx)
                Call ProcessNode(Grid, OpenList, Current, Neighbor, nx, ny, EndX, EndY, CountOpen)
            End If
        Next i
    Loop
    
    If Found Then
        Call BuildPath(EndX, EndY, Grid, x, y)
    End If
End Sub

Private Function InitializeGrid(ByVal Rows As Long, ByVal Columns As Long) As Node()
    Dim i As Long
    Dim j As Long
    Dim ReturnArr() As Node
    ReDim ReturnArr(Rows, Columns)
    
    ' Initialize nodes
    For i = 0 To Rows
        For j = 0 To Columns
            With ReturnArr(i, j)
                .X = j
                .Y = i
                .StartCost = 0
                .Distance = 0
                .TotalCost = 0
                .ParentX = -1
                .ParentY = -1
                .Closed = False
                .Open = False
            End With
        Next j
    Next i
    InitializeGrid = ReturnArr
End Function

Private Sub BuildPath(ByVal EndX As Long, ByVal EndY As Long, ByRef Grid() As Node, ByRef x() As Long, ByRef y() As Long)
    Dim Path() As Node
    ReDim Path(0)
    Dim PathLen As Long
    Dim cx As Long : cx = EndX
    Dim cy As Long : cy = EndY
    
    Do While cx <> -1 And cy <> -1
        ReDim Preserve Path(PathLen)
        Path(PathLen) = Grid(cy, cx)
        PathLen = PathLen + 1
        
        Dim px As Long, py As Long
        px = Grid(cy, cx).ParentX
        py = Grid(cy, cx).ParentY
        cx = px
        cy = py
    Loop
    
    ' Path is from goal to start; reverse and compute offsets
    If PathLen < 2 Then Exit Sub ' no movement needed

    ReDim x(PathLen - 2) ' offsets = number of steps = nodes - 1
    ReDim y(PathLen - 2)
    Dim i As Long
    For i = 0 To PathLen - 2
        x(i) = Path(PathLen - 2 - i).X - Path(PathLen - 1 - i).X
        y(i) = Path(PathLen - 2 - i).Y - Path(PathLen - 1 - i).Y
    Next i
End Sub

Private Function FindLowestNode(ByRef OpenList() As Node, ByVal Count As Long) As Long
    Dim Index As Long
    Dim i As Long
    For i = 1 To Count - 1
        If OpenList(i).TotalCost < OpenList(Index).TotalCost Then Index = i
    Next i
    FindLowestNode = Index
End Function

Private Sub CloseNode(ByRef Grid() As Node, ByRef Current As Node, ByRef OpenList() As Node, ByVal Index As Long, ByRef CountOpen As Long)
    Dim i As Long
    If Index < CountOpen - 1 Then
        For i = Index To CountOpen - 2
            OpenList(i) = OpenList(i + 1)
        Next i
    End If
    CountOpen = CountOpen - 1
    Grid(Current.Y, Current.X).Closed = True
End Sub

Private Function FoundGoal(ByRef Current As Node, ByVal EndX As Long, ByVal EndY As Long) As Boolean
    FoundGoal = (Current.X = EndX And Current.Y = EndY)
End Function

Private Sub Neighbors(ByRef Current As Node, ByRef x() As Long, ByRef y() As Long)
    x(0) = Current.X + 1 : y(0) = Current.Y + 0
    x(1) = Current.X - 1 : y(1) = Current.Y + 0
    x(2) = Current.X + 0 : y(2) = Current.Y + 1
    x(3) = Current.X + 0 : y(3) = Current.Y - 1
End Sub

Private Sub ProcessNode(ByRef Grid() As Node, ByRef OpenList() As Node, ByRef Current As Node, ByRef Neighbor As Node, ByVal x As Long, ByVal y As Long, ByVal EndX As Long, ByVal EndY As Long, ByRef CountOpen As Long)
    Dim StartCost As Long
    StartCost = Current.StartCost + 1
    
    If Neighbor.Open = False Or StartCost < Neighbor.StartCost Then
        Grid(y, x).StartCost = StartCost
        Grid(y, x).Distance = GetDistance(x, y, EndX, EndY)
        Grid(y, x).TotalCost = Grid(y, x).StartCost + Grid(y, x).Distance
        Grid(y, x).ParentX = Current.X
        Grid(y, x).ParentY = Current.Y
        
        If Neighbor.Open = False Then
            CountOpen = AddNode(OpenList, Grid(y, x))
        End If
    End If
End Sub

Private Function GetDistance(ByVal X As Long, ByVal Y As Long, ByVal GoalX As Long, ByVal GoalY As Long) As Long
    GetDistance = Abs(GoalY - Y) + Abs(GoalX - X)
End Function

Private Function AddNode(ByRef Arr() As Node, ByRef Value As Node) As Long
    On Error Resume Next
    Dim Size As Long
    Size = UBound(Arr) + 1
    ReDim Preserve Arr(Size)
    Value.Open = True
    Arr(Size) = Value
    AddNode = Size + 1
End Function