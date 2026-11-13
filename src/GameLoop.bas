Attribute VB_Name = "GameLoop"
Option Explicit

' Excel session + fixed-timestep loop. Logic runs at exactly LOGIC_HZ,
' rendering happens once per loop iteration if at least one tick ran.
' Entry point: StartGame. Esc = pause menu in a fight, quit on the title screen.
' If Excel ever gets stuck (no input), run RestoreExcel from the Immediate window.

Public Const LOGIC_HZ As Long = 20                 ' 1 tick = 1 "frame" for frame data
Private Const DT As Double = 1# / LOGIC_HZ
Private Const MAX_STEPS As Long = 5                ' max catch-up ticks per loop
Private Const MAX_FRAME_TIME As Double = 0.25

Private mRunning As Boolean
Private mAnim As Boolean               ' Excel's animation setting before the game started

Public Sub StartGame()
    If mRunning Then Exit Sub
    On Error GoTo Fail

    BeginSession
    Game.GameInit False
    Game.QuitRequested = False
    mRunning = True

    Dim prevT As Double, nowT As Double, frameT As Double, acc As Double, lastPump As Double
    Dim steps As Long, r0 As Double, r1 As Double, waitMs As Long, focused As Boolean, wasFocused As Boolean
    Dim statT As Double, statFrames As Long, statTicks As Long, statPaints As Long, statRender As Double

    prevT = WinApi.NowSec()
    statT = prevT: lastPump = prevT
    wasFocused = True
    Do While mRunning
        ' another window in front: keep ticking (the fight auto-pauses) but do not paint;
        ' give Windows plenty of time, and repaint everything once Excel is back
        focused = KeyInput.HasFocus()
        If focused And Not wasFocused Then Renderer.Invalidate
        wasFocused = focused
        nowT = WinApi.NowSec()
        frameT = nowT - prevT
        If frameT > MAX_FRAME_TIME Then frameT = MAX_FRAME_TIME
        If frameT < 0 Then frameT = 0
        prevT = nowT
        acc = acc + frameT

        steps = 0
        Do While acc >= DT And steps < MAX_STEPS
            Game.GameTick
            If Game.QuitRequested Then mRunning = False
            acc = acc - DT
            steps = steps + 1
        Loop
        If steps = MAX_STEPS Then acc = 0                ' too slow: drop time instead of spiraling
        statTicks = statTicks + steps

        If steps > 0 And mRunning And focused Then
            r0 = WinApi.NowSec()
            Game.GameRender
            r1 = WinApi.NowSec()
            statRender = statRender + (r1 - r0)
            statFrames = statFrames + 1
            statPaints = statPaints + Renderer.PaintCalls
            DoEvents                                     ' let Excel show the frame (once per frame, not per loop)
            lastPump = r1
        ElseIf nowT - lastPump > 0.015 Then
            DoEvents                                     ' keep Excel / Windows responsive
            lastPump = nowT
        End If
        If Not focused Then
            WinApi.SleepMs 20
            DoEvents
        End If

        ' stats once per second
        If nowT - statT >= 1 Then
            If statFrames > 0 Then
                Application.StatusBar = "FPS " & statFrames & " | ticks/s " & statTicks & _
                    " | paint calls/frame " & Format(statPaints / statFrames, "0.0") & _
                    " | render ms/frame " & Format(1000 * statRender / statFrames, "0.0") & _
                    " | LV " & CpuAI.LevelName(Game.CpuLevel) & " | " & Fighter.StatusText & " | Esc = pause"
            End If
            statT = nowT: statFrames = 0: statTicks = 0: statPaints = 0: statRender = 0
        End If

        ' sleep until the next tick is due (accurate thanks to timeBeginPeriod 1)
        waitMs = Int((DT - acc - (WinApi.NowSec() - nowT)) * 1000) - 1
        If waitMs > 0 Then WinApi.SleepMs waitMs
    Loop

    EndSession
    Exit Sub

Fail:
    Dim n As Long, d As String
    n = Err.Number: d = Err.Description
    mRunning = False
    EndSession
    If n <> 18 Then MsgBox "Error " & n & ": " & d, vbExclamation, "Excel Murim"   ' 18 = user break
End Sub

' Measures how fast this PC paints frames: plays a CPU vs CPU fight for a few
' seconds with each paint option and reports milliseconds per frame.
' Run: Alt+F8 -> RenderBenchmark (takes about 20 seconds).
Public Sub RenderBenchmark()
    Dim results As String, v As Long, f As Long, t0 As Double, tr As Double, calls As Long
    If mRunning Then Exit Sub
    On Error GoTo Fail
    BeginSession
    For v = 0 To 1
        Renderer.KeepScreenUpdating = (v = 1)
        Game.GameInit False
        Game.WatchMode = True
        Game.GameMode = MODE_VSCPU
        Game.Multi = False
        Game.StartMatch 1
        tr = 0: calls = 0
        For f = 1 To 150
            Game.GameTick
            t0 = WinApi.NowSec()
            Game.GameRender
            DoEvents
            tr = tr + (WinApi.NowSec() - t0)
            calls = calls + Renderer.PaintCalls
        Next f
        results = results & IIf(v = 0, "ScreenUpdating off/on per frame: ", "ScreenUpdating always on: ") & _
                  Format(1000 * tr / 150, "0.0") & " ms/frame, " & Format(calls / 150, "0.0") & " paint calls" & vbCrLf
    Next v
    Renderer.KeepScreenUpdating = False
    Game.WatchMode = False
    EndSession
    MsgBox results & vbCrLf & "Lower is better. Send these numbers to tune the renderer.", vbInformation, "Excel Murim"
    Exit Sub
Fail:
    Renderer.KeepScreenUpdating = False
    EndSession
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Excel Murim"
End Sub

Public Sub StopGame()
    mRunning = False
End Sub

Public Sub RestoreExcel()
    mRunning = False
    EndSession
End Sub

Private Sub BeginSession()
    Application.EnableCancelKey = xlErrorHandler
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.Interactive = False      ' block keystrokes from editing cells; GetAsyncKeyState still sees them
    On Error Resume Next
    mAnim = True
    mAnim = CallByName(Application, "EnableAnimations", VbGet)   ' late bound: not in every Excel version
    CallByName Application, "EnableAnimations", VbLet, False    ' no animated cell / scroll effects
    Application.DisplayStatusBar = True
    On Error GoTo 0
    WinApi.HighResTimer True             ' 1 ms sleep / timer resolution while playing
End Sub

Private Sub EndSession()
    On Error Resume Next
    WinApi.HighResTimer False
    CallByName Application, "EnableAnimations", VbLet, mAnim
    Application.Interactive = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableCancelKey = xlInterrupt
    Application.StatusBar = False
End Sub
