VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit



Private p_Definition As ItemDefinition
Private p_Amount As IRange

Public Property Let Definition(n_Definition As ItemDefinition) : Set p_Definition = n_Definition : End Property
Public Property Let Amount(n_Amount         As IRange)          : Set p_Amount     = n_Amount     : End Property

Public Property Get Definition()            As ItemDefinition  : Set Definition   = p_Definition : End Property
Public Property Get Amount()                As IRange           : Set Amount       = p_Amount     : End Property




Public Function Create(ByVal Definition As ItemDefinition, ByVal Amount As IRange) As Item
    Set Create = New Item
    With Create
        .Definition = Definition
        .Amount     = Amount
    End With
End Function

Public Function Use() As Boolean
    If Amount.Value > 0 Then
        Amount.Value = Amount.Value - 1
        Call Definition.Command.Run()
        Use = True
    End If
End Function

Public Function FullName() As String
    FullName = Definition.Name & " : " & Amount.Value()
End Function