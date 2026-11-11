Attribute VB_Name = "Hud"
Option Explicit

' Health bars (with red damage trail), names, super gauges, combo counter,
' banners (COUNTER / SUPER / ESCAPE), KO banner.

Private Const BAR_W As Long = 50
Private Const BAR_Y As Long = 2
Private Const P1_X As Long = 4                       ' left bar outer x
Private Const P2_X As Long = 64                      ' right bar outer x (SCREEN_W - 4 - BAR_W - 2)
Private Const TRAIL_SPEED As Double = 6              ' hp per tick
Private Const METER_W As Long = 30
Private Const METER_Y As Long = 60

Private cBorder As Long, cEmpty As Long, cHp As Long, cHpDark As Long, cTrail As Long, cLow As Long
Private cMeter As Long, cMeterFull As Long
Private mTick As Long

Public Sub Init()
    cBorder = RGB(16, 14, 24)
    cEmpty = RGB(48, 40, 56)
    cHp = RGB(244, 200, 72)
    cHpDark = RGB(200, 140, 40)
    cTrail = RGB(200, 48, 48)
    cLow = RGB(240, 90, 60)
    cMeter = RGB(80, 200, 230)
    cMeterFull = RGB(255, 240, 140)
End Sub

Public Sub Update()
    Dim i As Long
    mTick = mTick + 1
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
    Dim t As String, owner As Long
    DrawBar 0, P1_X, False
    DrawBar 1, P2_X, True
    Font.DrawText UCase$(Fighter.F(0).prefix), P1_X + 1, 8, RGB(230, 230, 240)
    Font.DrawText UCase$(Fighter.F(1).prefix), P2_X + BAR_W + 1, 8, RGB(230, 230, 240), 1, 2
    DrawMeter 0, P1_X, False
    DrawMeter 1, SCREEN_W - P1_X - METER_W - 2, True
    If Combat.ComboCount(0) >= 2 Then Font.DrawText Combat.ComboCount(0) & " HITS", P1_X + 1, 15, RGB(255, 220, 90)
    If Combat.ComboCount(1) >= 2 Then Font.DrawText Combat.ComboCount(1) & " HITS", P2_X + BAR_W + 1, 15, RGB(255, 220, 90), 1, 2
    t = Combat.BannerText
    If Len(t) > 0 Then
        owner = Combat.BannerOwner
        If owner = 0 Then
            Font.DrawText t, P1_X + 1, 22, RGB(255, 140, 70)
        ElseIf owner = 1 Then
            Font.DrawText t, P2_X + BAR_W + 1, 22, RGB(255, 140, 70), 1, 2
        Else
            Font.DrawText t, SCREEN_W \ 2, 22, RGB(150, 220, 255), 1, 1
        End If
    End If
    If Combat.IsKO Then DrawKO
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

Private Sub DrawMeter(ByVal i As Long, ByVal ox As Long, ByVal fromRight As Boolean)
    Dim w As Long, full As Boolean, c As Long
    w = (METER_W * Fighter.F(i).meter) \ METER_MAX
    full = (Fighter.F(i).meter >= METER_MAX)
    Renderer.FillRect ox, METER_Y, METER_W + 2, 3, cBorder
    Renderer.FillRect ox + 1, METER_Y + 1, METER_W, 1, cEmpty
    If full Then
        If (mTick \ 3) Mod 2 = 0 Then c = cMeterFull Else c = cMeter
    Else
        c = cMeter
    End If
    If fromRight Then
        Renderer.FillRect ox + 1 + METER_W - w, METER_Y + 1, w, 1, c
        If full Then Font.DrawText "MAX", ox - 2, METER_Y - 2, cMeterFull, 1, 2
    Else
        Renderer.FillRect ox + 1, METER_Y + 1, w, 1, c
        If full Then Font.DrawText "MAX", ox + METER_W + 5, METER_Y - 2, cMeterFull
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
