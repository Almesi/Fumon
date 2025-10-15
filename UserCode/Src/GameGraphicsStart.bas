Attribute VB_Name = "GameGraphicsStart"


Option Explicit

Private RenderStack          As std_Stack

Public CurrentRenderObject   As New RenderObject
Public StartRenderObject     As New RenderObject
Public OverWorldRenderObject As New RenderObject
Public MapRenderObject       As New RenderObject
Public InventoryRenderObject As New RenderObject
Public FightRenderObject     As New RenderObject
Public FumonRenderObject     As New RenderObject
Public AttackRenderObject    As New RenderObject
Public OptionsRenderObject   As New RenderObject

Public UsedFont As VBGLFontLayout



Public Function StartGame() As Boolean
    Dim FreetypePath As String: FreetypePath = "C:\Users\deallulic\Documents\GitHub\VBGL\Code\Src\Externals"
    Dim FontPath     As String: FontPath     = "C:\Users\deallulic\Documents\GitHub\VBGL\Code\Res\Fonts\Consolas.ttf"

    Set UsedFont = VBGLFontLayout.Create(FreetypePath, FontPath, 48)
    Dim TextBox As VBGLTextBox
    Set TextBox = CreateTextBox(FreetypePath, FontPath)
    Set StartRenderObject = RenderObject.Create(CreateInput())
    Call StartRenderObject.AddDrawable(TextBox)
    Set RenderStack = New std_Stack
    Call RenderStack.Add(StartRenderObject)

    Application.EnableCancelKey = xlDisabled
    Call SetForegroundWindow(CurrentContext.CurrentWindow.ID)

    Call SetUpOverWorldGraphics()
    Call SetUpMapGraphics()
    Call SetUpAttackGraphics()
    Call SetUpFightGraphics()
    Call SetUpInventoryGraphics()
    Call SetUpFumonGraphics()
    Call SetUpOptionsGraphics()
    Call SetUpMessageGraphics()
    With CurrentContext
        Call .SetDisplayFunc(AddressOf       DisplayFuncStack)
        Call .SetIdleFunc(AddressOf          IdleFuncStack)
        Call .SetKeyboardFunc(AddressOf      KeyboardFuncStack)
        Call .SetKeyboardUpFunc(AddressOf    KeyboardFuncUpStack)
        Call .SetPassiveMotionFunc(AddressOf PassiveMotionFuncStack)
        Call .SetMouseWheelFunc(AddressOf    MouseWheelFuncStack)
        Set CurrentRenderObject = StartRenderObject
        Call .MainLoop()
    End With
    MeServer.GameMap.ServerStarter.Formula = Empty
    StartGame = True
End Function

Private Function CreateTextBox(ByVal LoadFilePath As String, ByVal FilePath As String) As VBGLTextBox
    Dim Temp As VBGLProperties
    Set Temp = FactoryTextBoxProperties.Clone()
    Call Temp.LetValueFamily("TopLeft*"     , -1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("TopRight*"    , +1.0!, +1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomLeft*"  , -1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("BottomRight*" , +1.0!, -1.0!, +0.0!)
    Call Temp.LetValueFamily("Color*"       , +1.0!, +1.0!, +1.0!, +0.0!)
    Set CreateTextBox = VBGLTextBox.CreateFromText(Temp, _
                                                   "Welcome to Fumon"            & vbCrLf & _
                                                   "To start the Game press [s]" & vbCrLf & _
                                                   "To View the Options [o]"     & vbCrLf & _
                                                   "To cancel Press [ESC]"       & vbCrLf, UsedFont)
End Function

Private Function CreateInput() As VBGLIInput
    Dim Temp As VBGLGeneralInput
    Set Temp = New VBGLGeneralInput
    Call Temp.AddKeyUp(Asc("s"), VBGLCallable.Create(Nothing, "AddRenderObject"   , vbMethod, 0, OverWorldRenderObject))
    Call Temp.AddKeyUp(Asc("o"), VBGLCallable.Create(Nothing, "AddRenderObject"   , vbMethod, 0, OptionsRenderObject))
    Call Temp.AddKeyUp(27      , VBGLCallable.Create(Nothing, "glutLeaveMainLoop" , vbMethod, -1))
    Set CreateInput = Temp
End Function

Public Sub AddRenderObject(Obj As RenderObject)
    Call RenderStack.Add(Obj)
    Set CurrentRenderObject = Obj
End Sub

Public Sub RemoveRenderObject()
    Call RenderStack.Delete()
    Set CurrentRenderObject = RenderStack.Value
End Sub

Private Sub CreateFontLayout(ByVal LoadFilePath As String, ByVal FilePath As String)
    Set UsedFont = VBGLFontLayout.Create(LoadFilePath, FilePath, 48, "Consolas")
End Sub

' Callback Functions, but stackable
Public Sub DisplayFuncStack()
    Call MeServer.Update()
    Call CurrentRenderObject.Loopp()
End Sub
Public Sub IdleFuncStack()
    Call MeServer.Update()
    Call CurrentRenderObject.Loopp()
End Sub
Public Sub KeyboardFuncStack(ByVal Char As Byte, ByVal x As Long, ByVal y As Long)
    Call CurrentRenderObject.KeyBoard(Char, x, y)
End Sub
Public Sub KeyboardFuncUpStack(ByVal Char As Byte, ByVal x As Long, ByVal y As Long)
    Call CurrentRenderObject.KeyBoardUp(Char, x, y)
End Sub
Public Sub PassiveMotionFuncStack(ByVal x As Long, ByVal y As Long)
    Call CurrentRenderObject.MouseMove(x, y)
End Sub
Public Sub MouseWheelFuncStack(ByVal wheel As Long, ByVal direction As Long, ByVal x As Long, ByVal y As Long)
    Call CurrentRenderObject.MouseWheel(wheel, direction, x, y)
End Sub