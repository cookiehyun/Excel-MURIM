Attribute VB_Name = "Hud"
Option Explicit

' Health bars (with red damage trail), KO banner, input monitor.

Private Const BAR_W As Long = 50
Private Const BAR_Y As Long = 2
Private Const P1_X As Long = 4                       ' left bar outer x
Private Const P2_X As Long = 64                     ' right bar outer x (SCREEN_W - 4 - BAR_W - 2)
Private Const TRAIL_SPEED As Double = 6              ' hp per tick

Private cBorder As Long, cEmpty As Long, cHp As Long, cHpDark As Long, cTrail As Long, cLow As Long

Public Sub Init()
    cBorder = RGB(16, 14, 24)
    cEmpty = RGB(48, 40, 56)
    cHp = RGB(244, 200, 72)
    cHpDark = RGB(200, 140, 40)
    cTrail = RGB(200, 48, 48)
    cLow = RGB(240, 90, 60)
End Sub

Public Sub Update()
    Dim i As Long
    For i = 0 To 1
        With Fighter.F(i)
            If .trailDelay > 0 Then
                .trailDelay = .trailDelay - 1
            ElseIf .healthShown > .health Then
                .healthShown = .healthShown - TRAIL_SPEED
                If .healthShown < .health Then .healthShown = .health
            ElseIf .healthShown < .health Then
                .healthShown = .health        ' refill after reset
            End If
        End With
    Next i
End Sub

Public Sub Draw()
    DrawBar 0, P1_X, False
    DrawBar 1, P2_X, True
    If Combat.IsKO Then DrawKO
    DrawInputs
End Sub

' fromRight = True: bar is anchored on its right edge and drains toward the right (P2 mirror).
Private Sub DrawBar(ByVal i As Long, ByVal ox As Long, ByVal fromRight As Boolean)
    Dim wHp As Long, wTrail As Long, top As Long, low As Boolean
    With Fighter.F(i)
        wHp = (BAR_W * .health + .maxHealth - 1) \ .maxHealth
        wTrail = (BAR_W * Int(.healthShown) + .maxHealth - 1) \ .maxHealth
        low = (.health * 4 < .maxHealth)
    End With
    Renderer.FillRect ox, BAR_Y, BAR_W + 2, 4, cBorder
    Renderer.FillRect ox + 1, BAR_Y + 1, BAR_W, 2, cEmpty
    top = IIf(low, cLow, cHp)
    If fromRight Then
        Renderer.FillRect ox + 1 + BAR_W - wTrail, BAR_Y + 1, wTrail, 2, cTrail
        Renderer.FillRect ox + 1 + BAR_W - wHp, BAR_Y + 1, wHp, 1, top
        Renderer.FillRect ox + 1 + BAR_W - wHp, BAR_Y + 2, wHp, 1, cHpDark
    Else
        Renderer.FillRect ox + 1, BAR_Y + 1, wTrail, 2, cTrail
        Renderer.FillRect ox + 1, BAR_Y + 1, wHp, 1, top
        Renderer.FillRect ox + 1, BAR_Y + 2, wHp, 1, cHpDark
    End If
End Sub

' "KO" in 5x5 pixel letters, scaled x2, centered.
Private Sub DrawKO()
    Dim rows As Variant, y As Long, x As Long, ch As String, ox As Long, oy As Long
    rows = Array("10010 01110", _
                 "10100 10001", _
                 "11000 10001", _
                 "10100 10001", _
                 "10010 01110")
    ox = SCREEN_W \ 2 - 11: oy = 14
    For y = 0 To 4
        For x = 1 To 11
            ch = Mid$(rows(y), x, 1)
            If ch = "1" Then
                Renderer.FillRect ox + (x - 1) * 2 - 1, oy + y * 2 - 1, 4, 4, cBorder
            End If
        Next x
    Next y
    For y = 0 To 4
        For x = 1 To 11
            ch = Mid$(rows(y), x, 1)
            If ch = "1" Then Renderer.FillRect ox + (x - 1) * 2, oy + y * 2, 2, 2, cLow
        Next x
    Next y
End Sub

' Tiny input monitor in the bottom-left corner: one 2x2 square per button.
Private Sub DrawInputs()
    Dim b As Long, c As Long
    For b = BTN_LEFT To BTN_HK
        If KeyInput.IsDown(b) Then c = RGB(255, 255, 255) Else c = RGB(60, 52, 46)
        Renderer.FillRect 1 + b * 3, SCREEN_H - 3, 2, 2, c
    Next b
End Sub
