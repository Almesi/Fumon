VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FumonBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_FumonString As IRange
Private p_Selected As Long

Public Property Get Count() As Long
    Dim Splits() As String
    Splits = Split(FumonString.Value, ",")
    Count = USize(Splits)
End Property

Public Property Let FumonString(ByVal n_FumonString As IRange)
    Set p_FumonString = n_FumonString
End Property
Public Property Get FumonString() As IRange
    Set FumonString = p_FumonString
End Property

Public Property Get Selected() As Long
    Selected = p_Selected
End Property
Public Property Let Selected(ByVal n_Selected As Long)
    p_Selected = n_Selected
End Property

Public Function Create(ByVal Rng As IRange) As FumonBox
    Set Create = New FumonBox
    With Create
        .FumonString = Rng
    End With
End Function

Public Function Add(ByVal Index As Long, ByVal Level As Long) As Long
    Dim Text As String
    Text = Index & " | " & Level

    If FumonString.Value = Empty Then
        FumonString.Value = Text
    Else
        FumonString.Value = "," & Text
    End If
    Add = Count
End Function

Public Sub Swap(ByVal SwapTo As Fumon, ByVal Index As Long, ByVal FumonDefinitions As PropCollection)
    Dim FumonIndex As Long, Level As Long
    Call GetInformation(Index, FumonIndex, Level)

    SwapTo.Definition = FumonDefinitions(FumonIndex)
    SwapTo.Number     = FumonIndex
    SwapTo.Level      = Level
    SwapTo.Expirience = 0
    Call SwapTo.HealResetAndRemove()
End Sub

Private Sub GetInformation(ByVal Index As Long, ByRef FumonIndex As Long, ByRef FumonLevel As Long)
    Dim Fumons()       As String : Fumons       = Split(FumonString.Value, ",")
    Dim Informations() As String : Informations = Split(Fumons(Index), " | ")
    FumonIndex   = Informations(0)
    FumonLevel   = Informations(1)
End Sub