Attribute VB_Name = "RenderMethods"


Option Explicit


'===========================
'=======RenderObjects=======
'===========================
Public RenderStack           As std_Stack
Public InputStack            As std_Stack

Public StartRenderObject     As VBGLRenderObject
Public OverWorldRenderObject As VBGLRenderObject
Public MapRenderObject       As VBGLRenderObject
Public InventoryRenderObject As VBGLRenderObject
Public FightRenderObject     As VBGLRenderObject
Public FumonRenderObject     As VBGLRenderObject
Public AttackRenderObject    As VBGLRenderObject
Public OptionsRenderObject   As VBGLRenderObject
Public MessageRenderObject   As VBGLRenderObject
Public QuestionRenderObject  As VBGLRenderObject

'===========================
'=========Factories=========
'===========================
Public FactoryTextBoxProperties As VBGLProperties
Public FactoryTextBox           As VBGLTextBox
Public UsedFont                 As VBGLFontLayout
Public GameTextures             As PropCollection

Public FarLeftSideFrame         As VBGLFrame
Public LeftSideFrame            As VBGLFrame

'===========================
'=========DrawStack=========
'===========================
Public Sub AddRenderObject(ByVal Obj As VBGLRenderObject)
    If Obj Is CurrentRenderObject Then Exit Sub
    Call RenderStack.Add(Obj)
    Call InputStack.Add(Obj.UserInput)
    Set CurrentRenderObject = Obj
End Sub

Public Sub RemoveRenderObject()
    Call RenderStack.Delete()
    Set CurrentRenderObject = RenderStack.Value
    Call InputStack.Delete()
End Sub

Public Sub AddDrawableToRenderObject(ByVal Drawable As VBGLDrawable, ByVal UserInput As VBGLIInput)
    'If Drawable Is CurrentRenderObject Then Exit Sub
    Call CurrentRenderObject.AddDrawable(Drawable)
    Call InputStack.Add(UserInput)
    CurrentRenderObject.UserInput = UserInput
End Sub

Public Sub RemoveDrawableFromRenderObject()
    Call CurrentRenderObject.RemoveDrawable()
    Call InputStack.Delete()
    CurrentRenderObject.UserInput = InputStack.Value
End Sub