VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParentRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private FirstRange    As Range
Private p_Rng         As Range
Private p_RowStart    As Long
Private p_ColumnStart As Long
Private p_RowEnd      As Long
Private p_ColumnEnd   As Long
Private Values()      As Variant
Private Formulas()    As Variant

Public Property Let Rng(ByVal      n_Rng            As Range) : Set p_Rng         = n_Rng         : End Property
Public Property Let RowStart(ByVal n_RowStart       As Long)  : Let p_RowStart    = n_RowStart    : End Property
Public Property Let ColumnStart(ByVal n_ColumnStart As Long)  : Let p_ColumnStart = n_ColumnStart : End Property
Public Property Let RowEnd(ByVal   n_RowEnd         As Long)  : Let p_RowEnd      = n_RowEnd      : End Property
Public Property Let ColumnEnd(ByVal   n_ColumnEnd   As Long)  : Let p_ColumnEnd   = n_ColumnEnd   : End Property

Public Property Get Rng()                           As Range  : Set Rng           = p_Rng         : End Property
Public Property Get RowStart()                      As Long   : Let RowStart      = p_RowStart    : End Property
Public Property Get ColumnStart()                   As Long   : Let ColumnStart   = p_ColumnStart : End Property
Public Property Get RowEnd()                        As Long   : Let RowEnd        = p_RowEnd      : End Property
Public Property Get ColumnEnd()                     As Long   : Let ColumnEnd     = p_ColumnEnd   : End Property

Public Sub LetValue(ByVal StartRow As Long, ByVal StartColumn As Long, ByVal n_Value As Variant)
    Dim Data() As Variant
    If IsArray(n_Value) = False Then
        ReDim Data(1 To 1, 1 To 1)
        Data(1, 1) = n_Value
    Else
        Data = n_Value
    End If
    Call WriteToSlice(Values, Data, StartRow, StartColumn, Ubound(Data, 1), Ubound(Data, 2))
End Sub
Public Function GetValue(ByVal StartRow As Long, ByVal StartColumn As Long, ByVal RowCount As Long, ByVal ColumnCount As Long) As Variant
    GetValue = GetSlice(Values, StartRow, StartColumn, RowCount, ColumnCount)
End Function

Public Sub LetFormula(ByVal StartRow As Long, ByVal StartColumn As Long, ByVal n_Value As Variant)
    Dim Data() As Variant
    If IsArray(n_Value) = False Then
        ReDim Data(1 To 1, 1 To 1)
        Data(1, 1) = n_Value
    Else
        Data = n_Value
    End If
    Call WriteToSlice(Formulas, Data, StartRow, StartColumn, Ubound(Data, 1), Ubound(Data, 2))
End Sub
Public Function GetFormula(ByVal StartRow As Long, ByVal StartColumn As Long, ByVal RowCount As Long, ByVal ColumnCount As Long) As Variant
    GetFormula = GetSlice(Formulas, StartRow, StartColumn, RowCount, ColumnCount)
End Function

Public Function RowCount() As Long
    RowCount = RowEnd - RowStart + 1
End Function
Public Function ColumnCount() As Long
    ColumnCount = ColumnEnd - ColumnStart + 1
End Function


Public Function Create(ByVal n_Rng As Range) As ParentRange
    Dim n_Values() As Variant
    Dim n_Formulas() As Variant
    If IsArray(n_Rng.Value) Then
        n_Values   = n_Rng.Value
        n_Formulas = n_Rng.Formula
    Else
        ReDim n_Values(1 To 1, 1 To 1)
        ReDim n_Formulas(1 To 1, 1 To 1)

        n_Values(1, 1) = n_Rng.Value
        n_Formulas(1, 1) = n_Rng.Formula
    End If

    Set Create = New ParentRange
    With Create
        .Rng         = n_Rng
        .RowStart    = .Rng.Row
        .ColumnStart = .Rng.Column
        .RowEnd      = .RowStart + .Rng.Rows.Count - 1
        .ColumnEnd   = .ColumnStart + .Rng.Columns.Count - 1
        Call .Init(n_Values, n_Formulas)
    End With
End Function

Public Sub Init(ByRef n_Values() As Variant, ByRef n_Formulas() As Variant)
    Values = n_Values
    Formulas = n_Formulas
    Set FirstRange = Rng.Parent.Cells(1, 1)
End Sub

Public Sub UpdateRange(ByVal n_RowStart As Long, ByVal n_ColumnStart As Long, ByVal n_RowEnd As Long, ByVal n_ColumnEnd As Long)
    Dim NewRowStart    As Long : NewRowStart     = IIf(n_RowStart    < RowStart   , n_RowStart   , RowStart)
    Dim NewColumnStart As Long : NewColumnStart  = IIf(n_ColumnStart < ColumnStart, n_ColumnStart, ColumnStart)
    Dim NewRowEnd      As Long : NewRowEnd       = IIf(n_RowEnd      > RowEnd     , n_RowEnd     , RowEnd)
    Dim NewColumnEnd   As Long : NewColumnEnd    = IIf(n_ColumnEnd   > ColumnEnd  , n_ColumnEnd  , ColumnEnd)

    If NewRowStart    = RowStart And _
       NewColumnStart = ColumnStart And _
       NewRowEnd      = RowEnd   And _
       NewColumnEnd   = ColumnEnd   Then
        Exit Sub
    End If

    RowStart    = NewRowStart
    ColumnStart = NewColumnStart
    RowEnd      = NewRowEnd
    ColumnEnd   = NewColumnEnd
    Rng         = SubRange(NewRowStart, NewColumnStart, NewRowEnd - NewRowStart + 1, NewColumnEnd - NewColumnStart + 1)
    Values      = Rng.Value
    Formulas    = Rng.Formula
End Sub


Private Function GetSlice(ByRef Data() As Variant, ByVal StartRow As Long, ByVal StartColumn As Long, ByVal RowCount As Long, ByVal ColumnCount As Long) As Variant()
    Dim y As Long, x As Long
    Dim Result() As Variant

    ReDim Result(1 To RowCount, 1 To ColumnCount)

    For y = 1 To RowCount
        For x = 1 To ColumnCount
            Result(y, x) = Data(GetRow(StartRow + y), GetColumn(StartColumn + x))
        Next x
    Next y
    GetSlice = Result
End Function

Private Sub WriteToSlice(ByRef Goal As Variant, ByRef Data() As Variant, ByVal StartRow As Long, ByVal StartColumn As Long, ByVal RowCount As Long, ByVal ColumnCount As Long)
    Dim y As Long, x As Long
    For y = 1 To RowCount
        For x = 1 To ColumnCount
            Goal(GetRow(StartRow + y), GetColumn(StartColumn + x)) = Data(y, x)
        Next x
    Next y
    SubRange(StartRow, StartColumn, RowCount, ColumnCount).Value = Data
End Sub

Private Function SubRange(ByVal RowStart As Long, ByVal ColumnStart As Long, ByVal RowCount As Long, ByVal ColCount As Long) As Range
    Dim Start As Range: Set Start = FirstRange.Offset(RowStart - 1, ColumnStart - 1)
    Dim EndP  As Range: Set EndP  = Start.Offset(RowCount - 1, ColCount - 1)
    Set SubRange = Range(Start, EndP)
End Function

Private Function GetRow(ByVal Index As Long) As Long
    GetRow = Index - RowStart
End Function
Private Function GetColumn(ByVal Index As Long) As Long
    GetColumn = Index - ColumnStart
End Function