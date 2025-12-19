Attribute VB_Name = "LoadAll"


Option Explicit

Public NewErrorHandler As std_ErrorHandler


Public Sub Test()
    Dim Path As String
    Path = WBPath

    Dim FoldersToIgnore As Variant
    FoldersToIgnore = IgnoreFolders()

    Dim Shower As IDestination: Set Shower = Nothing
    Dim Logger As IDestination: Set Logger = std_ImmiedeateDestination.Create()
    
    Set NewErrorHandler = std_ErrorHandler.Create(Shower, Logger)
    
    Dim Proj As std_VBProject
    Set Proj = std_VBProject.Create(ThisWorkbook.VBProject, NewErrorHandler)

    If Proj.IncludeFolderArr(Path & "CodeBase\Libraries", NormalInclude, True, False, FoldersToIgnore) <> Proj.Handler.IS_ERROR Then
        If Proj.IncludeFolderArr(Path & "CodeBase\Src", NormalReplace, True, False, FoldersToIgnore) <> Proj.Handler.IS_ERROR Then
            If Proj.IncludeFolderArr(Path & "UserCode\Src", NormalReplace, True, False, FoldersToIgnore) <> Proj.Handler.IS_ERROR Then
                Debug.Print Application.Run("InitializeGame", Path)
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