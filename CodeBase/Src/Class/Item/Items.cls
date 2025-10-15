VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private Const p_ItemCount As Long = 2
Private p_Items() As Item
Private p_Selected As Long

Public Property Get ItemCount() As Long
    ItemCount = p_ItemCount
End Property

Public Property Get Items() As Variant
    Items = p_Items
End Property
Public Property Let Items(n_Items As Variant)
    p_Items = n_Items
End Property

Public Property Get Item(Index As Long) As Item
    Set Item = p_Items(Index)
End Property
Public Property Let Item(Index As Long, n_Item As Item)
    Set p_Items(Index) = n_Item
End Property

Public Property Get Selected() As Long
    Selected = p_Selected
End Property
Public Property Let Selected(n_Value As Long)
    p_Selected = n_Value
End Property

Public Function Count() As Long
    Count = USize(p_Items)
End Function

Public Property Get SelectedItem() As Item
    SelectedItem = Item(Selected)
End Property

Public Sub Increment(ByVal Value As Long)
    Selected = Selected + Value
    If Selected < 0 Then
        Selected = 0
    ElseIf Selected > MePlayer.Items.Count Then
        Selected = 0
    End If
End Sub

Public Sub AddItem(ByVal n_Item As Item)
    Call VBGLAdd(p_Items, n_Item)
End Sub

Public Function Create(Rng As Range) As Items
    Dim i As Long
    Set Create = New Items
    With Create
        Do Until Rng.Offset(0, i).Formula = Empty
            Call .AddItem(Factoryitem.Create(FactoryServer.ItemDefinition(Rng.offset(0, i).Value), Rng.Offset(0, i + 1)))
            i = i + ItemCount
        Loop
    End With
End Function