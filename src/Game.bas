Attribute VB_Name = "Game"
Option Explicit

' Pure game logic: menu, one tick of simulation and one render.
' No Excel session handling here, so the headless test harness can drive it.

Public Const GS_MENU As Long = 0       ' title screen
Public Const GS_FIGHT As Long = 1

Public Const MODE_TRAINING As Long = 0
Public Const MODE_VSCPU As Long = 1

Public GameState As Long
Public CpuLevel As Long                 ' CpuAI.LV_EASY / LV_NORMAL / LV_HARD
Public HumanIndex As Long            ' 0 = Danwoo, 1 = Cheolsan
Public WatchMode As Boolean              ' CPU vs CPU
Public GameMode As Long              ' MODE_TRAINING / MODE_VSCPU
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
        TitleTick
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
        TitleRender
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

' ---------------------------------------------------------------- title screen

Private Sub TitleTick()
    mMenuTick = mMenuTick + 1
    If KeyInput.IsPressed(BTN_UP) Or KeyInput.IsPressed(BTN_DOWN) Then mCursor = 1 - mCursor
    If KeyInput.IsPressed(BTN_LP) Then
        ' character / level / stage select come next; for now start with defaults
        GameMode = mCursor
        WatchMode = False
        StartMatch
        If GameMode = MODE_TRAINING Then Fighter.SetDummy DUMMY_IDLE
    End If
End Sub

Private Sub TitleRender()
    Dim k As Long, y As Long, c As Long, lbl As String, cx As Long
    cx = SCREEN_W \ 2
    Renderer.BeginFrame True
    ' fighters facing each other on both sides
    Renderer.DrawSprite Data.AnimFrame(Fighter.F(0).aIdle, mMenuTick), 30, GROUND_Y, False
    Renderer.DrawSprite Data.AnimFrame(Fighter.F(1).aIdle, mMenuTick), SCREEN_W - 30, GROUND_Y, True
    ' title
    PixelFont.DrawBig "EXCEL", cx, 5, RGB(230, 220, 200), 2
    PixelFont.DrawBig "MUHYEOP", cx, 21, RGB(255, 204, 92), 3
    Renderer.FillRect cx - 50, 46, 100, 1, RGB(200, 150, 60)
    Renderer.FillRect cx - 38, 48, 76, 1, RGB(140, 100, 40)
    Renderer.FillRect cx - 52, 45, 3, 3, RGB(200, 150, 60)
    Renderer.FillRect cx + 49, 45, 3, 3, RGB(200, 150, 60)
    ' menu
    For k = 0 To 1
        y = 57 + k * 13
        If k = 0 Then lbl = "TRAINING" Else lbl = "VS CPU"
        If k = mCursor Then
            Renderer.FillRect cx - 30, y - 3, 60, 11, RGB(40, 28, 20)
            Renderer.DrawRectOutline cx - 30, y - 3, 60, 11, RGB(230, 180, 80)
            c = RGB(255, 236, 180)
            If (mMenuTick \ 6) Mod 2 = 0 Then
                PixelFont.DrawText ">", cx - 26, y, RGB(255, 200, 90)
                PixelFont.DrawText "<", cx + 24, y, RGB(255, 200, 90)
            End If
        Else
            Renderer.FillRect cx - 30, y - 3, 60, 11, RGB(20, 16, 26)
            c = RGB(130, 120, 150)
        End If
        PixelFont.DrawText lbl, cx, y, c, 1, 1
    Next k
    PixelFont.DrawText "UP DOWN = SELECT    A = OK    ESC = QUIT", cx, 92, RGB(150, 140, 170), 1, 1
End Sub
