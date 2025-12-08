VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FasterRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IRange

' Descibes this Instance´s Positions
Private p_Parent      As ParentRange
Private p_Row         As Long
Private p_Column      As Long
Private p_RowCount    As Long
Private p_ColumnCount As Long

Public Property Let Parent(ByVal       n_Parent      As ParentRange)  : Set p_Parent         = n_Parent      : End Property
Public Property Let Row(ByVal          n_Row         As Long)         : Let p_Row            = n_Row         : End Property
Public Property Let Column(ByVal       n_Column      As Long)         : Let p_Column         = n_Column      : End Property
Public Property Let RowCount(ByVal     n_RowCount    As Long)         : Let p_RowCount       = n_RowCount    : End Property
Public Property Let ColumnCount(ByVal  n_ColumnCount As Long)         : Let p_ColumnCount    = n_ColumnCount : End Property

Public Property Get Parent()                         As ParentRange   : Set Parent           = p_Parent      : End Property
Public Property Get Row()                            As Long          : Let Row              = p_Row         : End Property
Public Property Get Column()                         As Long          : Let Column           = p_Column      : End Property
Public Property Get RowCount()                       As Long          : Let RowCount         = p_RowCount    : End Property
Public Property Get ColumnCount()                    As Long          : Let ColumnCount      = p_ColumnCount : End Property

' Constructor helper: loads Values + Formulas into the array
Public Function CreateFromRange(ByVal n_Rng As Range) As FasterRange
    Set CreateFromRange = New FasterRange
    With CreateFromRange
        .Parent      = ParentRange.Create(n_Rng)
        .Row         = .Parent.RowStart
        .Column      = .Parent.ColumnStart
        .RowCount    = .Parent.Rng.Rows.Count
        .ColumnCount = .Parent.Rng.Columns.Count
    End With
End Function

' Constructor helper: creates a FasterRange that references an existing data Slice
Public Function CreateFromSlice(ByVal n_RowStart As Long, ByVal n_ColStart As Long, ByVal n_RowCount As Long, ByVal n_ColCount As Long) As FasterRange
    Set CreateFromSlice = New FasterRange
    With CreateFromSlice
        Call Parent.UpdateRange(n_RowStart, n_ColStart, n_RowStart + n_RowCount - 1, n_ColStart + n_ColCount - 1)
        .Parent      = Parent
        .Row         = n_RowStart
        .Column      = n_ColStart
        .RowCount    = n_RowCount
        .ColumnCount = n_ColCount
    End With
End Function

'===============================
' IRange IMPLEMENTATION
'===============================
Private Function IRange_GetSelf() As Object
    Set IRange_GetSelf = Me
End Function

Private Property Let IRange_Value(ByVal n_Value As Variant)
    Call Parent.LetValue(Row, Column, n_Value)
End Property

Private Property Get IRange_Value() As Variant
    Dim Result As Variant
    Result = Parent.GetValue(Row, Column, RowCount, ColumnCount)
    If Ubound(Result, 1) = 1 And Ubound(Result, 2) = 1 Then
        IRange_Value = Result(1, 1)
    Else
        IRange_Value = Parent.GetValue(Row, Column, RowCount, ColumnCount)
    End If
End Property

Private Property Let IRange_Formula(ByVal n_Formula As Variant)
    Call Parent.LetFormula(Row, Column, n_Formula)
End Property

Private Property Get IRange_Formula() As Variant
    Dim Result As Variant
    Result = Parent.GetFormula(Row, Column, RowCount, ColumnCount)
    If Ubound(Result, 1) = 1 And Ubound(Result, 2) = 1 Then
        IRange_Formula = Result(1, 1)
    Else
        IRange_Formula = Parent.GetFormula(Row, Column, RowCount, ColumnCount)
    End If
End Property

Private Function IRange_Offset(ByVal y As Long, ByVal x As Long) As IRange
    Set IRange_Offset = CreateFromSlice(Row + y, Column + x, RowCount, ColumnCount)
End Function