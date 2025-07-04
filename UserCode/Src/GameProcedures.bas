Attribute VB_Name = "GameProcedures"


Option Explicit


Public FumonRenderStack As std_Stack

Private CurrentSelectedItem   As Long
Private CurrentSelectedFumon  As Long
Private CurrentSelectedAttack As Long
Public CurrentSelectedPreviousAttack As String

Public CurrentFightMove  As FumonGame.FightMove
Public CurrentFightValue As Variant

Public Enum FumonRenderType
    NoRender        = 0
    OverWorld       = 1
    Map             = 2
    Fight           = 3
    Inventory       = 4
    Fumons          = 5
    Attacks         = 6
    PreviousAttacks = 7
End Enum

Public Enum FumonLookDirection
    Up    = 0
    Left  = 1
    Down  = 2
    Right = 3
End Enum

Public Sub CallBackKeyBoard(ByVal Char As Byte, ByVal x As Long, ByVal y As Long)
    Select Case CurrentFumonRender
        Case FumonRenderType.NoRender
            Select Case Char
                Case Asc("ENTER") : Call FumonRenderStack.Add(FumonRenderType.OverWorld)' Start Game
            End Select
        Case FumonRenderType.OverWorld
            Select Case Char
                Case Asc("w")    : Call MePlayer.MovePath(+0, -1)                                                                 ' Move Up
                Case Asc("a")    : Call MePlayer.MovePath(-1, +0)                                                                 ' Move Left
                Case Asc("s")    : Call MePlayer.MovePath(+0, +1)                                                                 ' Move Down
                Case Asc("d")    : Call MePlayer.MovePath(+1, +0)                                                                 ' Move Right
                Case Asc("UP")   : Call MePlayer.Look(FumonLookDirection.Up)                                                      ' Look Up
                Case Asc("LEFT") : Call MePlayer.Look(FumonLookDirection.Left)                                                    ' Look Left
                Case Asc("DOWN") : Call MePlayer.Look(FumonLookDirection.Down)                                                    ' Look Down
                Case Asc("RIGHT"): Call MePlayer.Look(FumonLookDirection.Right)                                                   ' Look Right
                Case Asc("SPACE"): Call MePlayer.Interact(MePlayer.TileInfront(1))                                                ' Interact with block in front of you
                Case Asc("i")    : Call FumonRenderStack.Add(FumonRenderType.Inventory)  ' Open Inventory menu
                Case Asc("f")    : Call FumonRenderStack.Add(FumonRenderType.Fumons)     ' Open Fumon menu
                Case Asc("m")    : Call FumonRenderStack.Add(FumonRenderType.Map)        ' Open Map menu
                Case Asc("ESC")  : Call FumonRenderStack.AfterDelete                     ' Go Back to previous Render
            End Select
        Case FumonRenderType.Map
            Select Case 
                Case Asc("ESC")  : Call FumonRenderStack.AfterDelete                     ' Go Back to previous Render
            End Select
        Case FumonRenderType.Fight
            Select Case 
                Case Asc("f")    : Call FumonRenderStack.Add(FumonRenderType.Fumon)     ' Open Fumons menu
                Case Asc("i")    : Call FumonRenderStack.Add(FumonRenderType.Inventory) ' Open Inventory menu
                Case Asc("r")    : CurrentFightMove = FumonGame.FightMove.Flee
                Case Asc("a")    : Call FumonRenderStack.Add(FumonRenderType.Attacks)   ' Open Attacks menu
            End Select
        Case FumonRenderType.Inventory
            Select Case 
                Case Asc("w")     : CurrentSelectedItem = CurrentSelectedItem - 1: If CurrentSelectedItem < 0 Or CurrentSelectedItem > Ubound(MePlayer.Items) Then CurrentSelectedItem = 0 ' Select Item UP
                Case Asc("s")     : CurrentSelectedItem = CurrentSelectedItem + 1: If CurrentSelectedItem < 0 Or CurrentSelectedItem > Ubound(MePlayer.Items) Then CurrentSelectedItem = 0 ' Select Item DOWN
                Case Asc("SPACE") : CurrentFightValue = MePlayer.Items.Item(CurrentSelectedItem): CurrentFightMove = FumonGame.FightMove.UseItem                                                     ' Use selected Item
                Case Asc("ESC")   : Call FumonRenderStack.AfterDelete                                                                                                                      ' Go Back to previous Render
            End Select
        Case FumonRenderType.Fumons
            Select Case 
                Case Asc("1")     : CurrentSelectedFumon = 0                                                                  ' Select Fumon 1
                Case Asc("2")     : CurrentSelectedFumon = 1                                                                  ' Select Fumon 2
                Case Asc("3")     : CurrentSelectedFumon = 2                                                                  ' Select Fumon 3
                Case Asc("4")     : CurrentSelectedFumon = 3                                                                  ' Select Fumon 4
                Case Asc("5")     : CurrentSelectedFumon = 4                                                                  ' Select Fumon 5
                Case Asc("6")     : CurrentSelectedFumon = 5                                                                  ' Select Fumon 6
                Case Asc("7")     : CurrentSelectedFumon = 6                                                                  ' Select Fumon 7
                Case Asc("8")     : CurrentSelectedFumon = 7                                                                  ' Select Fumon 8
                Case Asc("SPACE") : CurrentFightValue = CurrentSelectedFumon: CurrentFightMove = FumonGame.FightMove.ChangeFumon        ' Switch selected Fumon with Fumon 1
                Case Asc("ESC")   : Call FumonRenderStack.AfterDelete                                                         ' Go Back to previous Render
                Case Asc("a")     : Call FumonRenderStack.Add(FumonRenderType.Attacks)                                        ' Open Attack menu
            End Select
        Case FumonRenderType.Attacks
                Case Asc("1")     : CurrentSelectedAttack = 0                                                                                ' Select Attack 1
                Case Asc("2")     : CurrentSelectedAttack = 1                                                                                ' Select Attack 2
                Case Asc("3")     : CurrentSelectedAttack = 2                                                                                ' Select Attack 3
                Case Asc("4")     : CurrentSelectedAttack = 3                                                                                ' Select Attack 4
                Case Asc("SPACE") : CurrentFightValue = MePlayer.Attacks.Attack(CurrentSelectedAttack): CurrentFightMove = FumonGame.FightMove.Attack  ' Use Selected Attack
                Case Asc("ENTER") : Call MePlayer.Attacks.Swap(0, CurrentSelectedAttack)                                                     ' Switch selected Attack with Attack 1
                Case Asc("ESC")   : Call FumonRenderStack.AfterDelete                                                                        ' Go Back to previous Render
        Case FumonRenderType.PreviousAttacks
                Case Asc("1"), Asc("2"), Asc("3"), Asc("4"), Asc("5"), Asc("6"), Asc("7"), Asc("8"), Asc("9"), Asc("0") : CurrentSelectedPreviousAttack = CurrentSelectedPreviousAttack & Chr(Char)              ' Increment Attack Number
                Case Asc("DELETE")                                                                                      : CurrentSelectedPreviousAttack = Empty                                                  ' Delete Attack Number
                Case Asc("SPACE")                                                                                       : Call MePlayer.Attacks.Attack(0) = MeServer.Attack(CLng(CurrentSelectedPreviousAttack)) ' Overwrite Attack1 with selected Attack Number
                Case Asc("ESC")                                                                                         : Call FumonRenderStack.AfterDelete                                                      ' Go Back to previous Render          
    End Select
End Sub