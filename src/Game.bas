Attribute VB_Name = "Game"
Option Explicit

' Pure game logic: menu, one tick of simulation and one render.
' No Excel session handling here, so the headless test harness can drive it.

Public Const GS_MENU As Long = 0       ' title screen
Public Const GS_FIGHT As Long = 1
Public Const GS_SELECT As Long = 2
Public Const GS_LEVEL As Long = 3
Public Const GS_STAGE As Long = 4
Public Const GS_LADDER As Long = 5     ' arcade: next opponent
Public Const GS_CONTINUE As Long = 6   ' arcade: lost -> continue countdown
Public Const GS_ENDING As Long = 7     ' arcade: cleared

Public Const MODE_TRAINING As Long = 0
Public Const MODE_VSCPU As Long = 1
Public Const MODE_SANDBOX As Long = 2      ' harness: no rounds, no refills
Public Const MODE_ARCADE As Long = 3
Public Const MODE_TAG As Long = 4

Public GameState As Long
Public CpuLevel As Long                 ' CpuAI.LV_EASY / LV_NORMAL / LV_HARD
Public HumanIndex As Long            ' which fighter the keyboard controls (0 = left)
Public P1Char As String              ' character names picked on the select screen
Public P2Char As String
Public P1Char2 As String             ' tag mode partners
Public P2Char2 As String
Public StageIndex As Long            ' Data stage (0 = mountain gate)
Public WatchMode As Boolean              ' CPU vs CPU
Public Multi As Boolean                  ' 2 players on one keyboard
Public GameMode As Long              ' MODE_TRAINING / MODE_VSCPU
Public QuitRequested As Boolean       ' title screen Esc -> GameLoop stops
Public HarnessSeed As Long            ' harness: seed for the next auto-restarted match (varies every match)
Public AutoRestart As Boolean        ' harness: start a new match after KO instead of the menu

Private mPaused As Boolean
Private mPauseSel As Long


Public Sub GameInit(ByVal headless As Boolean)
    Renderer.Headless = headless
    KeyInput.Init
    Data.LoadAll
    KeyConfig.Load
    Renderer.Init
    FX.Init
    Hud.Init
    CpuLevel = LV_NORMAL
    HumanIndex = 0
    P1Char = "danwoo"
    P2Char = "cheolsan"
    WatchMode = False
    Screens.Init
    StartMatch 0                       ' build fighters so the menu can show them
    GameState = GS_MENU
End Sub

Public Sub StartMatch(Optional ByVal seed As Long = 0)
    Renderer.SetStage StageIndex
    If GameMode = MODE_TAG Then
        Tag.Begin P1Char, P2Char, P1Char2, P2Char2, HumanIndex
    Else
        Fighter.Init P1Char, P2Char, HumanIndex
        Tag.Off
    End If
    Combat.Init
    Projectile.Init
    FX.Init
    mPaused = False
    CpuAI.Init CpuLevel, seed
    Dim rounds As Boolean, training As Boolean
    rounds = (GameMode = MODE_VSCPU Or GameMode = MODE_ARCADE Or GameMode = MODE_TAG)
    training = (GameMode = MODE_TRAINING)
    Match.BeginMatch rounds, training
    GameState = GS_FIGHT
End Sub

Public Sub GameTick()
    Dim i As Long, frozen As Boolean
    KeyInput.Poll
    If GameState = GS_FIGHT Then
        If KeyInput.IsPressed(BTN_QUIT) Then                ' Esc = pause menu
            mPaused = Not mPaused
            mPauseSel = 0
            Exit Sub
        End If
        If KeyInput.FocusLost() And Not mPaused Then        ' switched to another window: auto-pause
            mPaused = True
            mPauseSel = 0
            Exit Sub
        End If
        If mPaused Then
            PauseTick
            Exit Sub
        End If
    End If
    If GameState = GS_MENU Then
        Screens.TitleTick
        Exit Sub
    ElseIf GameState = GS_SELECT Then
        Screens.SelectTick
        Exit Sub
    ElseIf GameState = GS_LEVEL Then
        Screens.LevelTick
        Exit Sub
    ElseIf GameState = GS_STAGE Then
        Screens.StageTick
        Exit Sub
    ElseIf GameState >= GS_LADDER Then
        Arcade.ScreenTick
        Exit Sub
    End If

    For i = 0 To 1
        If Not Match.InputAllowed Then
            Fighter.FeedInput i, 5, 0, 0                ' intro / round over: nobody moves
        ElseIf Multi Then
            Fighter.FeedHuman i, 1 + i                  ' 1P / 2P key sets
        ElseIf IsCpu(i) Then
            CpuAI.Observe i
            CpuAI.Think i
        ElseIf i = Fighter.Controlled Then
            Fighter.FeedHuman i
        Else
            Fighter.FeedDummy i
        End If
    Next i

    Combat.TickTimers
    If Combat.SuperFreeze > 0 Then
        Combat.SuperFreeze = Combat.SuperFreeze - 1
    ElseIf Combat.HitStop > 0 Then
        Combat.HitStop = Combat.HitStop - 1
    Else
        Tag.Inputs Match.InputAllowed
        Tag.Tick
        Fighter.Update
        Projectile.Update
        Combat.Resolve
    End If
    FX.Update
    Hud.Update
    frozen = (Combat.SuperFreeze > 0 Or Combat.HitStop > 0)
    Match.Tick frozen
End Sub

Public Function IsCpu(ByVal i As Long) As Boolean
    If Multi Then
        IsCpu = False
    ElseIf WatchMode Then
        IsCpu = True
    ElseIf i <> Fighter.Controlled Then
        IsCpu = (Fighter.DummyMode = DUMMY_CPU)
    End If
End Function

Public Sub GameRender()
    Dim dark As Boolean
    If GameState = GS_MENU Then
        Screens.TitleRender
    ElseIf GameState = GS_SELECT Then
        Screens.SelectRender
    ElseIf GameState = GS_LEVEL Then
        Screens.LevelRender
    ElseIf GameState = GS_STAGE Then
        Screens.StageRender
    ElseIf GameState >= GS_LADDER Then
        Arcade.ScreenRender
    Else
        dark = (Combat.SuperFreeze > 0)
        Renderer.BeginFrame dark
        Fighter.Draw
        Projectile.Draw
        FX.Draw
        Hud.Draw
        Match.Draw
        If mPaused Then PauseRender
    End If
    Renderer.Flush
End Sub


' ---------------------------------------------------------------- pause menu (Esc during a fight)

' Start the same match again (training keeps its idle dummy).
Public Sub Restart()
    StartMatch
    If GameMode = MODE_TRAINING Then Fighter.SetDummy DUMMY_IDLE
End Sub

Private Sub PauseTick()
    If KeyInput.IsPressed(BTN_UP) Then mPauseSel = (mPauseSel + 3) Mod 4
    If KeyInput.IsPressed(BTN_DOWN) Then mPauseSel = (mPauseSel + 1) Mod 4
    If KeyInput.IsPressed(BTN_HP) Then mPaused = False             ' S = resume
    If KeyInput.IsPressed(BTN_LP) Then
        mPaused = False
        Select Case mPauseSel
        Case 1: Restart
        Case 2: Screens.SelectBegin
        Case 3: GameState = GS_MENU
        End Select
    End If
End Sub

Private Sub PauseRender()
    Dim k As Long, y As Long, c As Long, lbl As String, cx As Long
    cx = SCREEN_W \ 2
    Renderer.FillRect cx - 40, 22, 80, 58, RGB(14, 10, 22)
    Renderer.DrawRectOutline cx - 40, 22, 80, 58, RGB(230, 180, 80)
    PixelFont.DrawBig "PAUSE", cx, 26, RGB(255, 204, 92), 1
    For k = 0 To 3
        y = 39 + k * 9
        Select Case k
        Case 0: lbl = "RESUME"
        Case 1: lbl = "RESTART"
        Case 2: lbl = "CHARACTER SELECT"
        Case Else: lbl = "MAIN MENU"
        End Select
        If k = mPauseSel Then
            Renderer.FillRect cx - 36, y - 2, 72, 9, RGB(60, 40, 20)
            c = RGB(255, 236, 180)
        Else
            c = RGB(140, 130, 160)
        End If
        PixelFont.DrawText lbl, cx, y, c, 1, 1
    Next k
    PixelFont.DrawText "ESC OR S = RESUME", cx, 74, RGB(110, 100, 130), 1, 1
End Sub
