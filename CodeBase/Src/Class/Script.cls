VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private Const CommandSeperator As String = vbCrLf
Private Const ArgSeperator     As String = "; "
Private       p_Number         As Long
Private       p_Name           As String
Private       p_Commands()     As String

Public Property Let Number(n_Number       As Long)   : p_Number    = n_Number   : End Property
Public Property Let Name(n_Name           As String) : p_Name      = n_Name     : End Property
Public Property Let Commands(n_Commands() As String) : p_Commands  = n_Commands : End Property

Public Property Get Number()              As Long    : Number      = p_Number   : End Property
Public Property Get Name()                As String  : Name        = p_Name     : End Property
Public Property Get Commands()            As String(): Commands    = p_Commands : End Property


Public Function Create(Rng As Range) As Script
    Dim Temp() As String
    Temp = Split(Rng.Offset(0, 2).Value, CommandSeperator)
    Set Create = New Script
    With Script
        .Number   = Rng.Offset(0, 0).Value
        .Name     = Rng.Offset(0, 1).Value
        .Commands = Temp
    End With
End Function

Public Function Run(Arguments As Variant) As Variant
    Dim i As Long
    Dim Arguments() As Variant
    For i = 0 To ArraySize(Commands)
        Arguments = GetProcedureArguments(Commands(i))
        If RunApplication(ReturnVariable, GetProcedureName(Commands(i)), Arguments) = False Then Exit Function
    Next i
    Run = ReturnVariable
End Function

Private Function RunApplication(ReturnVariable As Variant, Name As String, Arguments() As Variant) As Boolean
    On Error GoTo Error
    If IsMissing(Arguments) Then
        ReturnVariable = Application.Run(Name)
    Else
        Select Case ArraySize(Arguments)
            Case 00:   ReturnVariable = Application.Run(Name, Arguments(0))
            Case 01:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1))
            Case 02:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2))
            Case 03:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3))
            Case 04:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4))
            Case 05:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5))
            Case 06:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6))
            Case 07:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7))
            Case 08:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8))
            Case 09:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9))
            Case 10:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10))
            Case 11:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11))
            Case 12:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12))
            Case 13:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13))
            Case 14:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14))
            Case 15:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15))
            Case 16:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16))
            Case 17:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17))
            Case 18:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18))
            Case 19:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19))
            Case 20:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20))
            Case 21:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21))
            Case 22:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22))
            Case 23:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23))
            Case 24:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24))
            Case 25:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25))
            Case 26:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26))
            Case 27:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27))
            Case 28:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27), Arguments(28))
            Case 29:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27), Arguments(28), Arguments(29))
            Case Else
            MsgBox("Too many Arguments: " & ArraySize(Arguments))
        End Select
    End If
    If IsError(ReturnVariable) Then
        GoTo Error:
    Else
        RunApplication = True
        Exit Function
    End If
    Error:
    MsgBox("Could not run Procedure. Procedure might not exist: " & Name)
End Function

Public Function GetProcedureArguments(Text As String) As Variant()
    Dim Temp() As String
    Temp = Split(GetParanthesesText(Text), ArgSeperator)
    GetProcedureArguments = ArrayConvertVariant(Temp)
End Function