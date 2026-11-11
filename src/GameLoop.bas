Attribute VB_Name = "GameLoop"
Option Explicit

' Fixed-timestep loop: logic runs at exactly LOGIC_HZ, rendering happens
' once per loop iteration if at least one logic tick ran.
' Entry point: StartGame. Esc quits.
' If Excel ever gets stuck (no input), run RestoreExcel from the Immediate window.

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Const LOGIC_HZ As Long = 20                 ' 1 tick = 1 "frame" for frame data
Private Const DT As Double = 1# / LOGIC_HZ
Private Const MAX_STEPS As Long = 5                ' max catch-up ticks per loop
Private Const MAX_FRAME_TIME As Double = 0.25

Private mRunning As Boolean
Private mTick As Long

Public Sub StartGame()
    If mRunning Then Exit Sub
    On Error GoTo Fail

    BeginSession
    KeyInput.Init
    Data.LoadAll
    Renderer.Init
    Fighter.Init
    mRunning = True
    mTick = 0

    Dim prevT As Double, nowT As Double, frameT As Double, acc As Double
    Dim steps As Long
    Dim statT As Double, statFrames As Long, statTicks As Long, statPaints As Long

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
            KeyInput.Poll
            If KeyInput.IsPressed(BTN_QUIT) Then mRunning = False
            Fighter.Update
            mTick = mTick + 1
            acc = acc - DT
            steps = steps + 1
        Loop
        If steps = MAX_STEPS Then acc = 0                ' too slow: drop time instead of spiraling
        statTicks = statTicks + steps

        If steps > 0 And mRunning Then
            Render
            statFrames = statFrames + 1
            statPaints = statPaints + Renderer.PaintCalls
        End If

        ' stats once per second
        If nowT - statT >= 1 Or nowT < statT Then
            If statFrames > 0 Then
                Application.StatusBar = "FPS " & statFrames & " | ticks/s " & statTicks & _
                    " | paint calls/frame " & Format(statPaints / statFrames, "0.0") & _
                    " | " & Fighter.StatusText & " | Esc = quit"
            End If
            statT = nowT: statFrames = 0: statTicks = 0: statPaints = 0
        End If

        DoEvents
        If acc < DT Then Sleep 1
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

Private Sub Render()
    Renderer.BeginFrame
    Fighter.Draw
    DrawInputDisplay
    Renderer.Flush
End Sub

' Tiny input monitor in the top-left corner: one 2x2 square per button.
Private Sub DrawInputDisplay()
    Dim b As Long, c As Long
    For b = BTN_LEFT To BTN_HK          ' direction + 4 attack buttons
        If KeyInput.IsDown(b) Then c = RGB(255, 255, 255) Else c = RGB(30, 40, 70)
        Renderer.FillRect 1 + b * 3, 1, 2, 2, c
    Next b
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
