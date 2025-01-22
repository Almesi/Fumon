Attribute VB_Name = "GeneralMethods"


Option Explicit


Public Function DEFINEPOINT(Optional Tile As String = "0", Optional NPC As String = "0", Optional Script As String = "0") As String
    If IsNumeric(Tile)   = False Then Tile   = GetData(Tile  , "Tiles"  , "Name", 0)
    If IsNumeric(NPC)    = False Then NPC    = GetData(NPC   , "NPCs"   , "Name", 0)
    If IsNumeric(Script) = False Then Script = GetData(Script, "Scripts", "Name", 0)
    DEFINEPOINT = Tile & "," & NPC & "," & Script
End Function

Public Function GetData(SearchValue As String, Sheet As String, SearchIn As String, ReturnColumn As String) As String
    Dim Offset As Long
    Dim CurrentColumn As Long
    Dim Rng As Range
    If IsNumeric(ReturnColumn) Then 
        Offset = CLng(ReturnColumn)
    Else
        Offset = ThisWorkbook.Sheets(Sheet).Rows(1).Find(ReturnColumn).Column
    End If

    If SearchIn = "0" Then
        Set Rng = ThisWorkbook.Sheets(Sheet).Range("A1").Offset(CLng(SearchValue) - 1, 0)
    Else
        Set Rng = ThisWorkbook.Sheets(Sheet).Find(SearchValue)
    End If
    CurrentColumn = Rng.Column
    If ReturnColumn = 0 Then
        GetData = Rng.Row
    Else
        GetData = Rng.Offset(0, Offset - CurrentColumn).Value
    End If
End Function

Public Function ExtractPoint(DefinedPoint As String, ExtractWhat As String) As String
    Dim Data() As String
    Data = Split(DefinedPoint, ",")
    Select Case ExtractWhat
        Case "Tile"   : ExtractPoint = Data(0)
        Case "NPC"    : ExtractPoint = Data(1)
        Case "Script" : ExtractPoint = Data(2)
    End Select
End Function