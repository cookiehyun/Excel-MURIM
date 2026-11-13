Attribute VB_Name = "Hud"
Option Explicit

' Health bars (with red damage trail), names, super gauges, combo counter,
' banners (COUNTER / SUPER / ESCAPE), KO banner.

Private Const BAR_W As Long = 80
Private Const BAR_Y As Long = 2
Private Const P1_X As Long = 4                       ' left bar outer x
Private Const P2_X As Long = 106                     ' right bar outer x (SCREEN_W - 4 - BAR_W - 2)
Private Const TRAIL_SPEED As Double = 6              ' hp per tick
Private Const METER_W As Long = 48
Private Const METER_Y As Long = 96
Private Const NG_W As Long = 36
Private Const NG_Y As Long = 92

Private cBorder As Long, cEmpty As Long, cHp As Long, cHpDark As Long, cTrail As Long, cLow As Long
Private cMeter As Long, cMeterFull As Long, cNg As Long, cNgLow As Long, cRed As Long
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
    cNg = RGB(110, 220, 140)
    cNgLow = RGB(230, 80, 70)
    cRed = RGB(110, 30, 40)
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
    PixelFont.DrawText UCase$(Fighter.F(0).prefix), P1_X + 1, 8, RGB(230, 230, 240)
    PixelFont.DrawText UCase$(Fighter.F(1).prefix), P2_X + BAR_W + 1, 8, RGB(230, 230, 240), 1, 2
    If Tag.Active Then
        DrawReserve 0, P1_X, False
        DrawReserve 1, P2_X, True
    End If
    DrawMeter 0, P1_X, False
    DrawMeter 1, SCREEN_W - P1_X - METER_W - 2, True
    DrawNaegong 0, P1_X, False
    DrawNaegong 1, SCREEN_W - P1_X - NG_W - 2, True
    If Combat.ComboCount(0) >= 2 Then PixelFont.DrawText Combat.ComboCount(0) & " HITS", P1_X + 1, 15, RGB(255, 220, 90)
    If Combat.ComboCount(1) >= 2 Then PixelFont.DrawText Combat.ComboCount(1) & " HITS", P2_X + BAR_W + 1, 15, RGB(255, 220, 90), 1, 2
    t = Combat.BannerText
    If Len(t) > 0 Then
        owner = Combat.BannerOwner
        If owner = 0 Then
            PixelFont.DrawText t, P1_X + 1, 22, RGB(255, 140, 70)
        ElseIf owner = 1 Then
            PixelFont.DrawText t, P2_X + BAR_W + 1, 22, RGB(255, 140, 70), 1, 2
        Else
            PixelFont.DrawText t, SCREEN_W \ 2, 22, RGB(150, 220, 255), 1, 1
        End If
    End If
    If Combat.IsKO Then DrawKO
End Sub

' fromRight = True: bar is anchored on its right edge and drains toward the right (P2 mirror).
Private Sub DrawBar(ByVal i As Long, ByVal ox As Long, ByVal fromRight As Boolean)
    Dim wHp As Long, wTrail As Long, top As Long, low As Boolean, wRed As Long
    With Fighter.F(i)
        wHp = (BAR_W * .health + .maxHealth - 1) \ .maxHealth
        wTrail = (BAR_W * Int(.healthShown) + .maxHealth - 1) \ .maxHealth
        low = (.health * 4 < .maxHealth)
    End With
    Renderer.FillRect ox, BAR_Y, BAR_W + 2, 4, cBorder
    Renderer.FillRect ox + 1, BAR_Y + 1, BAR_W, 2, cEmpty
    top = IIf(low, cLow, cHp)
    If Tag.Active Then                                      ' recoverable (red) part, drawn under the trail
        wRed = (BAR_W * Fighter.F(i).redCap + Fighter.F(i).maxHealth - 1) \ Fighter.F(i).maxHealth
        If fromRight Then
            Renderer.FillRect ox + 1 + BAR_W - wRed, BAR_Y + 1, wRed, 2, cRed
        Else
            Renderer.FillRect ox + 1, BAR_Y + 1, wRed, 2, cRed
        End If
    End If
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
        If full Then PixelFont.DrawText "MAX", ox - 2, METER_Y - 2, cMeterFull, 1, 2
    Else
        Renderer.FillRect ox + 1, METER_Y + 1, w, 1, c
        If full Then PixelFont.DrawText "MAX", ox + METER_W + 5, METER_Y - 2, cMeterFull
    End If
End Sub

' Tag mode: the waiting partner's name and a thin health bar under the main bar.
Private Sub DrawReserve(ByVal side As Long, ByVal ox As Long, ByVal fromRight As Boolean)
    Dim w As Long, wr As Long, mx As Long, nm As String, c As Long
    Const RW As Long = 40
    mx = Tag.ReserveMax(side)
    w = (RW * Tag.ReserveHealth(side) + mx - 1) \ mx
    wr = (RW * Tag.ReserveRed(side) + mx - 1) \ mx
    If Tag.ReserveHealth(side) <= 0 Then c = RGB(90, 80, 100) Else c = RGB(170, 160, 190)
    nm = "+" & UCase$(Tag.ReservePrefix(side))
    If fromRight Then
        Renderer.FillRect ox + 1 + BAR_W - RW - 1, BAR_Y + 4, RW + 2, 2, cBorder
        Renderer.FillRect ox + 1 + BAR_W - wr, BAR_Y + 4, wr, 1, cRed
        Renderer.FillRect ox + 1 + BAR_W - w, BAR_Y + 4, w, 1, cHpDark
        PixelFont.DrawText nm, ox + BAR_W - 44, 8, c, 1, 2
    Else
        Renderer.FillRect ox, BAR_Y + 4, RW + 2, 2, cBorder
        Renderer.FillRect ox + 1, BAR_Y + 4, wr, 1, cRed
        Renderer.FillRect ox + 1, BAR_Y + 4, w, 1, cHpDark
        PixelFont.DrawText nm, ox + 46, 8, c
    End If
End Sub

' Inner energy (naegong) bar just above the super gauge; blinks red when a special failed for lack of it.
Private Sub DrawNaegong(ByVal i As Long, ByVal ox As Long, ByVal fromRight As Boolean)
    Dim w As Long, c As Long
    w = (NG_W * Fighter.F(i).naegong) \ NAEGONG_MAX
    Renderer.FillRect ox, NG_Y, NG_W + 2, 3, cBorder
    If Fighter.F(i).ngFlash > 0 And (mTick \ 2) Mod 2 = 0 Then
        Renderer.FillRect ox + 1, NG_Y + 1, NG_W, 1, cNgLow
    Else
        Renderer.FillRect ox + 1, NG_Y + 1, NG_W, 1, cEmpty
    End If
    c = cNg
    If fromRight Then
        Renderer.FillRect ox + 1 + NG_W - w, NG_Y + 1, w, 1, c
    Else
        Renderer.FillRect ox + 1, NG_Y + 1, w, 1, c
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
