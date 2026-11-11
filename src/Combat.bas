Attribute VB_Name = "Combat"
Option Explicit

' Hit detection and resolution: hitbox vs hurtbox, guard check by level,
' damage, hit stop, sparks, KO.

Public HitStop As Long              ' ticks the whole game freezes after a hit
Private mKoTimer As Long
Private Const KO_TICKS As Long = 80

Public Sub Init()
    HitStop = 0
    mKoTimer = 0
End Sub

Public Property Get IsKO() As Boolean
    IsKO = (mKoTimer > 0)
End Property

Public Sub Resolve()
    If mKoTimer > 0 Then
        mKoTimer = mKoTimer - 1
        If mKoTimer = 0 Then Fighter.ResetRound     ' round system comes in step 6
        Exit Sub
    End If
    CheckHit 0, 1
    CheckHit 1, 0
End Sub

Private Sub CheckHit(ByVal a As Long, ByVal d As Long)
    Dim hx As Long, hy As Long, hw As Long, hh As Long
    Dim ux As Long, uy As Long, uw As Long, uh As Long
    Dim ix0 As Long, iy0 As Long, ix1 As Long, iy1 As Long
    Dim mv As Long, heavy As Boolean, push As Double, ko As Boolean, cx As Long, cy As Long

    If Fighter.F(a).hasHit Then Exit Sub
    If Not Fighter.HitRect(a, hx, hy, hw, hh) Then Exit Sub
    If Not Fighter.HurtRect(d, ux, uy, uw, uh) Then Exit Sub

    ix0 = Max2(hx, ux): iy0 = Max2(hy, uy)
    ix1 = Min2(hx + hw, ux + uw): iy1 = Min2(hy + hh, uy + uh)
    If ix1 <= ix0 Or iy1 <= iy0 Then Exit Sub

    Fighter.F(a).hasHit = True
    cx = (ix0 + ix1) \ 2
    cy = (iy0 + iy1) \ 2
    mv = Fighter.F(a).mv
    heavy = (Data.MvDamage(mv) >= 80)
    If heavy Then push = 2.2 Else push = 1.3

    If Blocks(d, Data.MvLevel(mv)) Then
        Fighter.ApplyBlock d, Data.MvBlockstun(mv), push * 0.8
        FX.Spawn cx, cy, FX_BLOCK
        HitStop = 2
    Else
        DealDamage d, Data.MvDamage(mv)
        ko = (Fighter.F(d).health <= 0)
        Fighter.ApplyHit d, Data.MvHitstun(mv), push, Data.MvKD(mv) Or ko
        Fighter.F(d).flash = 2
        If heavy Then
            FX.Spawn cx, cy, FX_HEAVY
            HitStop = 4
        Else
            FX.Spawn cx, cy, FX_HIT
            HitStop = 3
        End If
        If ko Then HitStop = 10
    End If
End Sub

' Guarding fighter blocks unless the level beats its stance.
Private Function Blocks(ByVal d As Long, ByVal level As Long) As Boolean
    With Fighter.F(d)
        If .state <> ST_GUARD Then Exit Function
        Select Case level
        Case LV_LOW: Blocks = .crouch
        Case LV_OVERHEAD: Blocks = Not .crouch
        Case Else: Blocks = True
        End Select
    End With
End Function

Public Sub DealDamage(ByVal d As Long, ByVal dmg As Long)
    With Fighter.F(d)
        .health = .health - dmg
        If .health < 0 Then .health = 0
        .trailDelay = 12
    End With
    If Fighter.F(d).health = 0 And mKoTimer = 0 Then mKoTimer = KO_TICKS
End Sub

Private Function Max2(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then Max2 = a Else Max2 = b
End Function

Private Function Min2(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then Min2 = a Else Min2 = b
End Function
