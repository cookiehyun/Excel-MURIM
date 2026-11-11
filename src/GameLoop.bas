Attribute VB_Name = "GameLoop"
Option Explicit

' Excel session + fixed-timestep loop. Logic runs at exactly LOGIC_HZ,
' rendering happens once per loop iteration if at least one tick ran.
' Entry point: StartGame. Esc quits.
' If Excel ever gets stuck (no input), run RestoreExcel from the Immediate window.

Public Const LOGIC_HZ As Long = 20                 ' 1 tick = 1 "frame" for frame data
Private Const DT As Double = 1# / LOGIC_HZ
Private Const MAX_STEPS As Long = 5                ' max catch-up ticks per loop
Private Const MAX_FRAME_TIME As Double = 0.25

Private mRunning As Boolean

Public Sub StartGame()
    If mRunning Then Exit Sub
    On Error GoTo Fail

    BeginSession
    Game.GameInit False
    mRunning = True

    Dim prevT As Double, nowT As Double, frameT As Double, acc As Double
    Dim steps As Long, r0 As Double, r1 As Double
    Dim statT As Double, statFrames As Long, statTicks As Long, statPaints As Long, statRender As Double

    prevT = Timer
    statT = prevT
    Do While mRunning
        nowT = Timer
        frameT = nowT - prevT
        If frameT < 0 Then frameT = frameT + 86400#    ' Timer wraps at midnight
        If frameT > MAX_FRAME_TIME Then frameT = MAX_FRAME_TIME
        prevT = nowT
        acc = acc + frameT

        steps = 0
        Do While acc >= DT And steps < MAX_STEPS
            Game.GameTick
            If KeyInput.IsPressed(BTN_QUIT) Then mRunning = False
            acc = acc - DT
            steps = steps + 1
        Loop
        If steps = MAX_STEPS Then acc = 0                ' too slow: drop time instead of spiraling
        statTicks = statTicks + steps

        If steps > 0 And mRunning Then
            r0 = Timer
            Game.GameRender
            r1 = Timer
            If r1 >= r0 Then statRender = statRender + (r1 - r0)
            statFrames = statFrames + 1
            statPaints = statPaints + Renderer.PaintCalls
        End If

        ' stats once per second
        If nowT - statT >= 1 Or nowT < statT Then
            If statFrames > 0 Then
                Application.StatusBar = "FPS " & statFrames & " | ticks/s " & statTicks & _
                    " | paint calls/frame " & Format(statPaints / statFrames, "0.0") & _
                    " | render ms/frame " & Format(1000 * statRender / statFrames, "0") & _
                    " | LV " & CpuAI.LevelName(Game.CpuLevel) & " | " & Fighter.StatusText & " | Esc = quit"
            End If
            statT = nowT: statFrames = 0: statTicks = 0: statPaints = 0: statRender = 0
        End If

        DoEvents
        If acc < DT Then WinApi.SleepMs 1
    Loop

    EndSession
    Exit Sub

Fail:
    Dim n As Long, d As String
    n = Err.Number: d = Err.Description
    mRunning = False
    EndSession
    If n <> 18 Then MsgBox "Error " & n & ": " & d, vbExclamation, "Excel Fighter"   ' 18 = user break
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
End Sub

Private Sub EndSession()
    On Error Resume Next
    Application.Interactive = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableCancelKey = xlInterrupt
    Application.StatusBar = False
End Sub
