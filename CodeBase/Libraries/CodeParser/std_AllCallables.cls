VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_AllCallables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

#Const DebugMode = 1

Private Type CallableComp
    Name As String
    Typee As vbext_ComponentType
    Callables() As std_Callable
    Size As Long
End Type

Private Comps() As CallableComp
Private Size    As Long

Private TempComponent As std_TempCode
Private RegExp As Object

Public Function CreateCallable(ByVal Text As String, ParamArray Args() As Variant) As std_Callable
    Dim ArgsV() As Variant
    ArgsV = Args
    Set CreateCallable = CreateCallableArr(Text, ArgsV)
End Function

Public Function CreateCallableArr(ByVal Text As String, ByRef Args() As Variant) As std_Callable
    Dim Func As Collection
    Set Func = ParseFunctions(Text)
    Set CreateCallableArr = InterpretChain(Func(1), Nothing, Args)
End Function

'===================================
'==========Interpretetaion==========
'===================================


Private Function InterpretChain(ByVal CurrentArg As Collection, ByVal Parent As Object, ByRef Args() As Variant) As Variant
    Dim i As Long
    Dim CurrentParent As Object
    Dim Result As Variant
    
    Set CurrentParent = Parent
    For i = 1 To CurrentArg.Count
        If TypeName(CurrentArg(i)) = "Collection" Then
            Set Result = InterpretFunction(CurrentArg(i), CurrentParent, Args)
            Set CurrentParent = Result
        Else
            Call InterpretArgument(Result, CurrentArg(i), Args)
        End If
    Next i
    
    Call Assign(InterpretChain, Result)
End Function

Private Function InterpretFunction(ByVal Func As Collection, ByVal Parent As Object, ByRef Args() As Variant) As Object
    Dim FuncName As String       : Let FuncName = Func(1)
    Dim FuncArgs As Collection   : Set FuncArgs = Func(2)
    Dim Found As std_Callable
    

    ' Might have to fix this part
    Dim Index As Long
    Index = ExtractArgAssign(FuncName)
    If Index <> -1 Then
        Set InterpretFunction = Args(Index)
        Exit Function
    Else
        Set Found = CopyFromCallable(FuncName, Parent, Args, FuncArgs.Count - 1)
    End If
    
    ' Interpret the arguments for this function and bind them
    Dim MyArgs() As Variant
    MyArgs = InterpretArguments(FuncArgs, Parent, Args)
    Call Found.BindArr(MyArgs)
    
    Set InterpretFunction = Found
End Function

Private Function InterpretArguments(ByVal Coll As Collection, ByVal Parent As Object, ByRef Args() As Variant) As Variant()
    Dim i As Long
    Dim Arr() As Variant
    Dim Size As Long
    Dim Item As Variant
    
    Size = Coll.Count - 1
    If Size = -1 Then Exit Function
    
    ReDim Arr(Size)
    For i = 0 To Size
        Call Assign(Item, Coll(i + 1))        
        Call Assign(Arr(i), InterpretChain(Item, Parent, Args))
    Next i
    
    InterpretArguments = Arr
End Function

Private Sub InterpretArgument(ByRef Goal As Variant, ByVal Text As String, ByRef Args() As Variant)
    Dim Index As Long
    Index = ExtractArgAssign(Text)
    If Index <> -1 Then
        Call Assign(Goal, Args(Index))
    ElseIf UCase(Text) Like "TRUE" Then
        Call Assign(Goal, True)
    ElseIf UCase(Text) Like "FALSE" Then
        Call Assign(Goal, False)
    Else
        Call Assign(Goal, Text)
    End If
End Sub

Private Function CopyFromCallable(ByVal Expression As String, ByVal Parent As Object, ByRef Args() As Variant, ByVal ArgSize As Long) As std_Callable
    Dim Found  As std_Callable
    Set Found = GetCallable(Expression, Parent, Args, ArgSize)
    Dim Result As std_Callable
    If Not Found Is Nothing Then
        Set Result = Found.Clone()
        Result.Object = Parent
    End If
    Set CopyFromCallable = Result
End Function

Private Function GetReturnType(ByVal Var As Variant) As String
    Dim Result As String
    Result = TypeName(Var)
    If Result = "std_Callable" Then
        Result = Var.ReturnType
    End If
    If Result = "Nothing" Then
        Result = Empty
    End If
    GetReturnType = Result
End Function

Private Function ExtractArgAssign(ByVal Text As String) As Long
    Dim StartPoint As Long
    StartPoint = InStr(1, Text, "$")
    If StartPoint <> 0 Then
        Dim i As Long
        Dim Result As String
        Dim Char As String
        For i = StartPoint + 1 To Len(Text)
            Char = Mid(Text, i, 1)
            If Mid(Text, i, 1) Like "[0123456789]" Then
                Result = Result & Char
            End If
        Next i
        If IsNumeric(Result) Then
            ExtractArgAssign = CLng(Result)
            Exit Function
        End If
    End If
    ExtractArgAssign = -1
End Function

Private Sub Assign(ByRef Goal As Variant, ByVal Value As Variant)
    If IsObject(Value) Then
        Set Goal = Value
    Else
        Let Goal = Value
    End If
End Sub

'===================================
'==========Find Callable============
'===================================

Private Function GetCallable(ByVal Name As String, ByVal Object As Object, ByRef Args() As Variant, ByVal ArgSize As Long) As std_Callable
    Dim Index As Long
    Index = ExtractArgAssign(Name)
    If Index <> -1 Then
        Name = GetReturnType(Args(Index))
    End If

    Dim ClassName As String
    ClassName = GetReturnType(Object)

    Dim Result As std_Callable
    Set Result = Find(ClassName, Name, ArgSize)
    If Not Result Is Nothing Then
        Set GetCallable = Result
    End If
End Function

Private Function Find(ByVal ClassName As String, ByVal MethodName As String, ByVal ArgSize As Long)
    Dim i As Long
    For i = 0 To Size
        If (ClassName = Empty And Comps(i).Typee = vbext_ct_StdModule) Or UCase(ClassName) = UCase(Comps(i).Name) Then
            Set Find = FindProcedure(Comps(i), MethodName, ArgSize)
            If Not Find Is Nothing Then Exit For
        End If
    Next i
End Function

Private Function FindProcedure(ByRef Comp As CallableComp, ByVal Name As String, ByVal ArgSize As Long) As std_Callable
    Dim i As Long
    For i = 0 To Comp.Size
        If UCase(Comp.Callables(i).Name) = UCase(Name) Then
            If Comp.Callables(i).CallType <> vbMethod Then
                If ArgSize = Comp.Callables(i).ArgCount Then
                    Set FindProcedure = Comp.Callables(i)
                    Exit Function  
                End If
            Else 
                Set FindProcedure = Comp.Callables(i)
                Exit Function
            End If
        End If
    Next i
End Function

'===================================
'==========From Component===========
'===================================

Public Function CreateFromProject(ByVal Proj As VBIDE.VBProject) As std_AllCallables
    Set CreateFromProject = New std_AllCallables
    With CreateFromProject
        Call .CreateComponents(Proj)
    End With
End Function

Public Sub CreateComponents(ByVal Proj As VBIDE.VBProject)
    Dim Comp    As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim Index   As Long
    Dim Code    As String

    Set TempComponent = std_TempCode.Create(ThisWorkbook.VBProject)
    For Each Comp in Proj.VBComponents
        Set CodeMod = Comp.CodeModule
        Index = AddComp(Comp.Type, Comp.Name)
        If CodeMod.CountOfLines > 0 Then
            Code  = CodeMod.Lines(1, CodeMod.CountOfLines)
            Call AddCode(Comps(Index), Code)
        End If
    Next Comp
    Set TempComponent = Nothing
End Sub


'===================================
'==============From File============
'===================================

Public Function CreateFromFolder(ByVal FolderPath As String) As std_AllCallables
    Set CreateFromFolder = New std_AllCallables
    With CreateFromFolder
        Call .CreateComponentsFromFolder(FolderPath)
    End With
End Function

Public Sub CreateComponentsFromFolder(ByVal FolderPath As String)
    Dim FSO As Object
    Dim Folder As Object
    Dim SubFolder As Object
    Dim File As Object
    Dim Index As Long
    Dim Code As String

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(FolderPath)

    Dim ThisOne As Boolean
    If TempComponent Is Nothing Then
        Set TempComponent = std_TempCode.Create(ThisWorkbook.VBProject)
        ThisOne = True
    End If
    For Each File In Folder.Files
        Index = AddComp(GetTypeFromName(File.Name), GetFileName(File.Name))
        Code  = ReadFile(File.Path)
        Call AddCode(Comps(Index), Code)
    Next File
    For Each SubFolder In Folder.SubFolders
        Call CreateComponentsFromFolder(SubFolder.Path)
    Next SubFolder
    If ThisOne Then Set TempComponent = Nothing
End Sub

Private Function ReadFile(ByVal FilePath As String) As String
    Dim FileNum As Integer
    Dim FileContent As String
    
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    FileContent = Input$(LOF(FileNum), FileNum)
    Close #FileNum
    
    ReadFile = FileContent
End Function

Private Function GetFileName(ByVal Name As String)
    Dim Found As Long
    Found = InStr(1, Name, ".")
    If Found = 0 Then
        Found = Len(Name)
    Else
        Found = Found - 1
    End If
    GetFileName = Mid(Name, 1, Found)
End Function


'===================================
'==============PARSING==============
'===================================

Private Function ParseFunctions(ByRef Text As String, Optional ByRef Pos As Long = 1) As Collection
    Dim Result As New Collection
    Dim CurrentArg As Collection      ' <- NEW: every argument is itself a Collection
    Dim Token As String
    Dim Char As String
    Dim FuncCollection As Collection
    Dim ArgList As Collection
    
    Set CurrentArg = New Collection   ' start first argument

    Do While Pos <= Len(Text)
        Char = Mid(Text, Pos, 1)
        
        Select Case Char
            Case "("
                Dim Name As String: Name = Trim(Token)
                Token = Empty

                Pos = Pos + 1
                Set ArgList = ParseFunctions(Text, Pos)

                Set FuncCollection = New Collection
                FuncCollection.Add Name
                If ArgList(1).Count = 0 Then
                    FuncCollection.Add New Collection
                Else
                    FuncCollection.Add ArgList
                End If
                CurrentArg.Add FuncCollection
            Case ")"
                If Trim(Token) <> Empty Then
                    CurrentArg.Add Trim(Token)
                    Token = Empty
                End If
                Result.Add CurrentArg
                GoTo Ending
            Case ","
                If Trim(Token) <> Empty Then
                    CurrentArg.Add Trim(Token)
                    Token = Empty
                End If
                Result.Add CurrentArg
                Set CurrentArg = New Collection
            Case "."
                If Trim(Token) <> Empty Then
                    Set FuncCollection = New Collection
                    FuncCollection.Add Trim(Token)        ' function name
                    FuncCollection.Add New Collection     ' empty arg list
                    CurrentArg.Add FuncCollection
                    Token = Empty
                End If
            Case Else
                Token = Token & Char
        End Select
        Pos = Pos + 1
    Loop

    If Trim(Token) <> Empty Then
        CurrentArg.Add Trim(Token)
    End If
    Result.Add CurrentArg

    Ending:
    Set ParseFunctions = Result
End Function



'===================================
'==============Creation=============
'===================================

Private Function AddComp(ByVal CompType As vbext_ComponentType, ByVal Name As String) As Long
    Size = Size + 1
    ReDim Preserve Comps(Size)
    Comps(Size).Name = Name
    Comps(Size).Typee = CompType
    Comps(Size).Size = -1
    AddComp = Size
End Function

Private Sub AddCode(ByRef Comp As CallableComp, ByVal Code As String)
    Dim CleanCode As String
    Dim Matches As Object
    Dim i As Long
    CleanCode = JoinContinuationLines(Code)
    Set Matches = RegExp.Execute(CleanCode)
    For i = 0 To Matches.Count - 1
        Call AddCallable(Comp.Typee, Comp.Name, Comp.Callables, Comp.Size, Matches(i).Value)
    Next i
End Sub

Private Sub AddCallable(ByVal CompType As Long, ByVal CompName As String, ByRef Callables() As std_Callable, ByRef Size As Long, ByVal Code As String)
    Dim Result As std_Callable
    Set Result = std_Callable.CreateFromText(CreateNewObject(ThisWorkbook.VBProject, CompType, CompName), Code)
    Size = Size + 1
    ReDim Preserve Callables(Size)
    Set Callables(Size) = Result
End Sub

Private Function GetTypeFromName(ByVal Name As String) As vbext_ComponentType
    Dim Parts() As String
    Parts = Split(Name, ".")
    Select Case Parts(1)
        Case "cls" : GetTypeFromName = vbext_ComponentType.vbext_ct_ClassModule
        Case "frm" : GetTypeFromName = vbext_ComponentType.vbext_ct_MSForm
        Case "bas" : GetTypeFromName = vbext_ComponentType.vbext_ct_StdModule
    End Select
End Function

Private Function CreateNewObject(ByVal Proj As VBIDE.VBProject, ByVal CompType As Long, ByVal ClassName As String) As Object
    'Currently disabled, as it is not debuggable
    #If DebugMode = 1 Then
    #Else
        If CompType = vbext_ComponentType.vbext_ct_ClassModule Then
            Dim Name As String : Name = "CreateTEMP" & ClassName
            Dim Code As String : Code = "Public Function " & Name & "() As Object: Set " & Name & " = New " & ClassName & " : End Function"
            Call TempComponent.AddFunction(Name, Code)
            Set CreateNewObject = TempComponent.CallFunction(Name)
        End If
    #End If
End Function

Private Function JoinContinuationLines(ByVal Code As String) As String
    Dim Lines() As String
    Dim i As Long
    Dim Result As String
    
    Lines = Split(Code, vbCrLf)
    i = 0
    
    Do While i <= UBound(Lines)
        Dim Line As String
        Line = Lines(i)
        
        ' Remove trailing underscore and concatenate with next line
        Do While Right(Line, 1) = "_"
            Line = Mid(Line, 1, Len(Line) - 1) & " " & Lines(i + 1)
            i = i + 1
            If i >= UBound(Lines) Then Exit Do
        Loop
        
        Result = Result & Line & vbCrLf
        i = i + 1
    Loop
    
    JoinContinuationLines = Result
End Function



Private Sub Class_Initialize()
    Size = -1
    Set RegExp = CreateObject("VBScript.RegExp")
    With RegExp
        .pattern = "(Public|Private|Friend)?\s*(Static\s*)?(Function|Sub|Property Get|Property Let|Property Set)\s+([A-Za-z_][A-Za-z0-9_]*)\s*\((.*?)(?=\))\)(.*?)(?=$|\r\n|\s*\:)"
        .IgnoreCase = True
        .Global = True
    End With
End Sub