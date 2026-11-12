Attribute VB_Name = "Game"
Option Explicit

' Pure game logic: menu, one tick of simulation and one render.
' No Excel session handling here, so the headless test harness can drive it.

Public Const GS_MENU As Long = 0
Public Const GS_FIGHT As Long = 1

Public GameState As Long
Public CpuLevel As Long                 ' CpuAI.LV_EASY / LV_NORMAL / LV_HARD
Public HumanIndex As Long            ' 0 = Danwoo, 1 = Cheolsan
Public WatchMode As Boolean              ' CPU vs CPU
Public AutoRestart As Boolean        ' harness: start a new match after KO instead of the menu

Private mCursor As Long
Private mMenuTick As Long

Public Sub GameInit(ByVal headless As Boolean)
    Renderer.Headless = headless
    KeyInput.Init
    Data.LoadAll
    Renderer.Init
    FX.Init
    Hud.Init
    CpuLevel = LV_NORMAL
    HumanIndex = 0
    WatchMode = False
    StartMatch 0                       ' build fighters so the menu can show them
    GameState = GS_MENU
End Sub

Public Sub StartMatch(Optional ByVal seed As Long = 0)
    Fighter.Init HumanIndex
    Combat.Init
    Projectile.Init
    FX.Init
    CpuAI.Init CpuLevel, seed
    GameState = GS_FIGHT
End Sub

Public Sub GameTick()
    Dim i As Long
    KeyInput.Poll
    If GameState = GS_MENU Then
        MenuTick
        Exit Sub
    End If

    For i = 0 To 1
        If IsCpu(i) Then
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
        Fighter.Update
        Projectile.Update
        Combat.Resolve
    End If
    FX.Update
    Hud.Update

    If Combat.KoFinished Then
        If AutoRestart Then
            StartMatch
        Else
            GameState = GS_MENU
        End If
    End If
End Sub

Public Function IsCpu(ByVal i As Long) As Boolean
    If WatchMode Then
        IsCpu = True
    ElseIf i <> Fighter.Controlled Then
        IsCpu = (Fighter.DummyMode = DUMMY_CPU)
    End If
End Function

Public Sub GameRender()
    Dim dark As Boolean
    If GameState = GS_MENU Then
        MenuRender
    Else
        dark = (Combat.SuperFreeze > 0)
        Renderer.BeginFrame dark
        Fighter.Draw
        Projectile.Draw
        FX.Draw
        Hud.Draw
    End If
    Renderer.Flush
End Sub

' ---------------------------------------------------------------- menu

Private Sub MenuTick()
    mMenuTick = mMenuTick + 1
    If KeyInput.IsPressed(BTN_UP) Then mCursor = (mCursor + 2) Mod 3
    If KeyInput.IsPressed(BTN_DOWN) Then mCursor = (mCursor + 1) Mod 3
    If KeyInput.IsPressed(BTN_LEFT) Or KeyInput.IsPressed(BTN_RIGHT) Then
        Select Case mCursor
        Case 0
            HumanIndex = 1 - HumanIndex
        Case 1
            If KeyInput.IsPressed(BTN_RIGHT) Then CpuLevel = (CpuLevel + 1) Mod 3 Else CpuLevel = (CpuLevel + 2) Mod 3
        Case Else
            WatchMode = Not WatchMode
        End Select
    End If
    If KeyInput.IsPressed(BTN_LP) Then StartMatch
End Sub

Private Sub MenuRender()
    Dim y As Long, k As Long, c As Long, v As String, lbl As String, sel As Long
    Renderer.BeginFrame True
    ' both fighters idle on the sides
    Renderer.DrawSprite Data.AnimFrame(Fighter.F(0).aIdle, mMenuTick), 13, GROUND_Y, False
    Renderer.DrawSprite Data.AnimFrame(Fighter.F(1).aIdle, mMenuTick), SCREEN_W - 13, GROUND_Y, True
    Renderer.FillRect 56, 32, 80, 29, RGB(18, 14, 30)
    Renderer.DrawRectOutline 56, 32, 80, 29, RGB(90, 70, 120)
    PixelFont.DrawText "EXCEL MUHYEOP", SCREEN_W \ 2, 5, RGB(255, 214, 120), 2, 1
    For k = 0 To 2
        y = 35 + k * 9
        Select Case k
        Case 0
            lbl = "FIGHTER"
            If HumanIndex = 0 Then v = "DANWOO" Else v = "CHEOLSAN"
        Case 1
            lbl = "LEVEL"
            v = (CpuLevel + 1) & " " & CpuAI.LevelName(CpuLevel)
        Case Else
            lbl = "MODE"
            If WatchMode Then v = "WATCH CPU" Else v = "VS CPU"
        End Select
        If k = mCursor Then c = RGB(255, 255, 255) Else c = RGB(140, 130, 160)
        PixelFont.DrawText lbl, 60, y, c
        If k = mCursor Then v = "<" & v & ">"
        PixelFont.DrawText v, 110, y, c, 1, 1
    Next k
    If (mMenuTick \ 8) Mod 2 = 0 Then PixelFont.DrawText "A = START", SCREEN_W \ 2, 70, RGB(255, 200, 90), 1, 1
End Sub
