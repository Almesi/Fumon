VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit



Private Sub Worksheet_Change(ByVal Target As Range)
    Select Case Target.Value
        Case "Fire"     : Target.Interior.Color = RGB(255, 000, 000)
        Case "Water"    : Target.Interior.Color = RGB(000, 000, 255)
        Case "Plant"    : Target.Interior.Color = RGB(000, 255, 000)
        Case "Ghost"    : Target.Interior.Color = RGB(255, 000, 255)
        Case "Psycho"   : Target.Interior.Color = RGB(255, 200, 200)
        Case "Dark"     : Target.Interior.Color = RGB(050, 050, 050)
        Case "Stone"    : Target.Interior.Color = RGB(128, 128, 128)
        Case "Ground"   : Target.Interior.Color = RGB(200, 000, 000)
        Case "Steel"    : Target.Interior.Color = RGB(200, 200, 200)
        Case "Poison"   : Target.Interior.Color = RGB(255, 000, 128)
        Case "Fly"      : Target.Interior.Color = RGB(050, 050, 128)
        Case "Bug"      : Target.Interior.Color = RGB(000, 200, 000)
        Case "Fight"    : Target.Interior.Color = RGB(255, 128, 128)
        Case "Salt"     : Target.Interior.Color = RGB(255, 255, 255)
        Case "Elektro"  : Target.Interior.Color = RGB(255, 255, 000)
        Case "Ice"      : Target.Interior.Color = RGB(000, 255, 255)
        Case "Normal"   : Target.Interior.Color = RGB(150, 150, 150)
        Case "Dragon"   : Target.Interior.Color = RGB(000, 000, 200)
        Case "Ancient"  : Target.Interior.Color = RGB(128, 000, 000)
        Case Else
    End Select
End Sub