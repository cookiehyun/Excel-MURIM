Attribute VB_Name = "Game"
Option Explicit

' Pure game logic: one tick of simulation and one render.
' No Excel session handling here, so the headless test harness can drive it.

Public Sub GameInit(ByVal headless As Boolean)
    Renderer.Headless = headless
    KeyInput.Init
    Data.LoadAll
    Renderer.Init
    Fighter.Init
    Combat.Init
    FX.Init
    Hud.Init
End Sub

Public Sub GameTick()
    KeyInput.Poll
    Fighter.BufferInput          ' record button presses even during hit stop
    If Combat.HitStop > 0 Then
        Combat.HitStop = Combat.HitStop - 1
    Else
        Fighter.Update
        Combat.Resolve
    End If
    FX.Update
    Hud.Update
End Sub

Public Sub GameRender()
    Renderer.BeginFrame
    Fighter.Draw
    FX.Draw
    Hud.Draw
    Renderer.Flush
End Sub
