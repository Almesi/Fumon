Attribute VB_Name = "LoadAll"


Option Explicit

Public NewErrorHandler As std_ErrorHandler


Public Sub Test()
    Dim Path As String
    Path = ThisWorkbook.Path

    Dim FoldersToIgnore As Variant
    FoldersToIgnore = IgnoreFolders()

    Dim Shower As IDestination: Set Shower = Nothing
    Dim Logger As IDestination: Set Logger = std_ImmiedeateDestination.Create()
    
    Set NewErrorHandler = std_ErrorHandler.Create(Shower, Logger)
    
    Dim Proj As std_VBProject
    Set Proj = std_VBProject.Create(ThisWorkbook.VBProject, NewErrorHandler)

    If Proj.IncludeFolderArr("C:\Users\deallulic\Documents\GitHub\VBA_StandardLibrary\Src\Class\Generics\CodeParser", NormalInclude, True, False, FoldersToIgnore) <> Proj.Handler.IS_ERROR Then
        If Proj.IncludeFolderArr("C:\Users\deallulic\Documents\GitHub\VBGL\Code\Src", NormalReplace, True, False, FoldersToIgnore) <> Proj.Handler.IS_ERROR Then
            If Proj.IncludeFolderArr("C:\Users\deallulic\Documents\GitHub\Fumon\CodeBase\Src", NormalReplace, True, False, FoldersToIgnore) <> Proj.Handler.IS_ERROR Then
                If Proj.IncludeFolderArr("C:\Users\deallulic\Documents\GitHub\Fumon\UserCode\Src", NormalReplace, True, False, FoldersToIgnore) <> Proj.Handler.IS_ERROR Then
                    If Proj.Include("C:\Users\deallulic\Documents\GitHub\stdVBA\src\stdLambda.cls", NormalReplace, True, False) <> Proj.Handler.IS_ERROR Then
                        If Proj.Include("C:\Users\deallulic\Documents\GitHub\stdVBA\src\stdICallable.cls", NormalReplace, True, False) <> Proj.Handler.IS_ERROR Then
                            Debug.Print Application.Run("InitializeGame")
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Function IgnoreFolders() As Variant
    Dim ReturnArr() As Variant
    ReDim ReturnArr(0)
    ReturnArr(0) = CVar("Errorhandling")
    IgnoreFolders = ReturnArr
End Function