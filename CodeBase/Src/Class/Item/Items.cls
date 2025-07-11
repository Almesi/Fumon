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
Private p_Items(p_ItemCount) As Item

Public Property Get ItemCount() As Long
    ItemCount = p_ItemCount
End Property

Public Property Get Items() As Item()
    Items = p_Items
End Property
Public Property Let Items(n_Items() As Item)
    p_Items = n_Items
End Property

Public Property Get Item(Index As Long) As Item
    Set Item = p_Items(Index)
End Property
Public Property Let Item(Index As Long, n_Item As Item)
    Set p_Items(Index) = n_Item
End Property

Public Function Create(Rng As Range) As Items
    Dim i As Long
    Set Create = New Items
    With Create
        Do Until Rng.Offset(0, i).Formula = Empty
            .Item(i) = FumonGame.Item.Create(FumonGame.FactoryServer.ItemDefinition(Rng.offset(0, i).Value), Rng.Offset(0, i + 1))
            i = i + ItemCount
        Loop
    End With
End Function