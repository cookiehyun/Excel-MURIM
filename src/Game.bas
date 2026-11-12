Attribute VB_Name = "Game"
Option Explicit

' Pure game logic: menu, one tick of simulation and one render.
' No Excel session handling here, so the headless test harness can drive it.

Public Const GS_MENU As Long = 0       ' title screen
Public Const GS_FIGHT As Long = 1
Public Const GS_SELECT As Long = 2
Public Const GS_LEVEL As Long = 3
Public Const GS_STAGE As Long = 4

Public Const MODE_TRAINING As Long = 0
Public Const MODE_VSCPU As Long = 1

Public GameState As Long
Public CpuLevel As Long                 ' CpuAI.LV_EASY / LV_NORMAL / LV_HARD
Public HumanIndex As Long            ' which fighter the keyboard controls (0 = left)
Public P1Char As String              ' character names picked on the select screen
Public P2Char As String
Public StageIndex As Long            ' Data stage (0 = mountain gate)
Public WatchMode As Boolean              ' CPU vs CPU
Public GameMode As Long              ' MODE_TRAINING / MODE_VSCPU
Public AutoRestart As Boolean        ' harness: start a new match after KO instead of the menu


Public Sub GameInit(ByVal headless As Boolean)
    Renderer.Headless = headless
    KeyInput.Init
    Data.LoadAll
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
    Fighter.Init P1Char, P2Char, HumanIndex
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
        Screens.TitleRender
    ElseIf GameState = GS_SELECT Then
        Screens.SelectRender
    ElseIf GameState = GS_LEVEL Then
        Screens.LevelRender
    ElseIf GameState = GS_STAGE Then
        Screens.StageRender
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

