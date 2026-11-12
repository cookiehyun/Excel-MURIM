Attribute VB_Name = "Combat"
Option Explicit

' Hit resolution shared by melee and projectiles:
'   guard check by level, chip damage (specials), counter hits, combo counter
'   with damage scaling, super meter, hit stop, corner pushback, KO.

Public HitStop As Long              ' ticks the whole game freezes after a hit
Public SuperFreeze As Long          ' ticks of the super "blackout" freeze
Public SuperOwner As Long

Private mKoTimer As Long
Public KoFinished As Boolean        ' set on the tick the KO pause ends
Private Const KO_TICKS As Long = 80

Private mCombo(0 To 1) As Long      ' hits in the current combo, by attacker
Private mComboShow(0 To 1) As Long  ' ticks to keep showing the counter
Private mComboShown(0 To 1) As Long
Private mText As String             ' short banner (COUNTER, ESCAPE)
Private mTextTicks As Long
Private mTextOwner As Long

Public Sub Init()
    HitStop = 0
    SuperFreeze = 0
    mKoTimer = 0
    KoFinished = False
    mCombo(0) = 0: mCombo(1) = 0
    mComboShow(0) = 0: mComboShow(1) = 0
    mTextTicks = 0
End Sub

Public Property Get IsKO() As Boolean
    IsKO = (mKoTimer > 0)
End Property

Public Property Get KoTicksLeft() As Long
    KoTicksLeft = mKoTimer
End Property

Public Function ComboCount(ByVal a As Long) As Long
    If mComboShow(a) > 0 Then ComboCount = mComboShown(a)
End Function

Public Function BannerText() As String
    If mTextTicks > 0 Then BannerText = mText
End Function

Public Property Get BannerOwner() As Long
    BannerOwner = mTextOwner
End Property

Public Sub ShowText(ByVal s As String, ByVal ticks As Long, Optional ByVal owner As Long = -1)
    mText = s
    mTextTicks = ticks
    mTextOwner = owner
End Sub

' A hit landed outside ApplyAttack (super grab sequence): count it in the combo.
Public Sub AddComboHit(ByVal a As Long)
    mCombo(a) = mCombo(a) + 1
    If mCombo(a) >= 2 Then
        mComboShown(a) = mCombo(a)
        mComboShow(a) = 40
    End If
End Sub

Public Sub StartSuper(ByVal i As Long)
    SuperFreeze = 12
    SuperOwner = i
    ShowText "SUPER", 24, i
End Sub

Public Sub AddMeter(ByVal i As Long, ByVal n As Long)
    With Fighter.F(i)
        .meter = .meter + n
        If .meter > METER_MAX Then .meter = METER_MAX
    End With
End Sub

' Called every tick (even during freezes) for timers.
Public Sub TickTimers()
    Dim a As Long
    For a = 0 To 1
        If mComboShow(a) > 0 Then mComboShow(a) = mComboShow(a) - 1
    Next a
    If mTextTicks > 0 Then mTextTicks = mTextTicks - 1
End Sub

Public Sub Resolve()
    Dim a As Long, d As Long, st As Long
    If mKoTimer > 0 Then
        mKoTimer = mKoTimer - 1
        If mKoTimer = 0 Then KoFinished = True
        Exit Sub
    End If
    ' combo ends when the defender is back in control
    For a = 0 To 1
        d = 1 - a
        st = Fighter.F(d).state
        If Not (st = ST_HIT Or st = ST_THROWN Or (st = ST_DOWN And Fighter.F(d).downPhase = 0)) Then mCombo(a) = 0
    Next a
    CheckMelee 0, 1
    CheckMelee 1, 0
End Sub

Private Sub CheckMelee(ByVal a As Long, ByVal d As Long)
    Dim hx As Long, hy As Long, hw As Long, hh As Long
    Dim ux As Long, uy As Long, uw As Long, uh As Long
    Dim ix0 As Long, iy0 As Long, ix1 As Long, iy1 As Long, cx As Long, cy As Long

    If Fighter.F(a).hasHit Then Exit Sub
    If Not Fighter.HitRect(a, hx, hy, hw, hh) Then Exit Sub
    If Not Fighter.HurtRect(d, ux, uy, uw, uh) Then Exit Sub

    ix0 = Max2(hx, ux): iy0 = Max2(hy, uy)
    ix1 = Min2(hx + hw, ux + uw): iy1 = Min2(hy + hh, uy + uh)
    If ix1 <= ix0 Or iy1 <= iy0 Then Exit Sub

    Fighter.F(a).hasHit = True
    cx = (ix0 + ix1) \ 2
    cy = (iy0 + iy1) \ 2
    ApplyAttack a, d, Fighter.F(a).mv, cx, cy, False, True
End Sub

' Resolve one hit of move mv from attacker a on defender d at (cx, cy).
Public Sub ApplyAttack(ByVal a As Long, ByVal d As Long, ByVal mv As Long, ByVal cx As Long, ByVal cy As Long, _
                       ByVal isProj As Boolean, ByVal lastHit As Boolean)
    Dim heavy As Boolean, push As Double, ko As Boolean, mt As Long, special As Boolean
    Dim dmg As Double, hits As Long, counter As Boolean, stun As Long, kd As Boolean

    If mKoTimer > 0 Then Exit Sub
    mt = Data.MvType(mv)
    special = (mt <> MT_NORMAL)
    hits = 1
    If isProj Then hits = Data.MvP2(mv)
    If hits < 1 Then hits = 1
    heavy = (Data.MvDamage(mv) >= 80)
    If heavy Then push = 2.2 Else push = 1.3
    If Not lastHit Then push = 0.4

    If mt = MT_SUPER_RUSH And Not Blocks(d, Data.MvLevel(mv)) Then
        FX.Spawn cx, cy, FX_SLASH                           ' super rush connected: slash sequence takes over
        Fighter.StartSuperSlash a, d
        Exit Sub
    End If

    If Blocks(d, Data.MvLevel(mv)) Then
        Fighter.ApplyBlock d, Data.MvBlockstun(mv), push * 0.8
        If special Then DealDamage d, Data.MvDamage(mv) \ (4 * hits)      ' chip damage
        AddMeter a, 15: AddMeter d, 10
        FX.Spawn cx, cy, FX_BLOCK
        HitStop = 2
    Else
        With Fighter.F(d)
            counter = (Not isProj) And (.state = ST_ATTACK) And (.stTick < Data.MvStartup(.mv))
        End With
        mCombo(a) = mCombo(a) + 1
        dmg = Data.MvDamage(mv) / hits
        If mCombo(a) >= 3 Then dmg = dmg * Max2D(0.5, 1 - 0.1 * (mCombo(a) - 2))   ' combo scaling
        If counter Then dmg = dmg * 1.25
        DealDamage d, Int(dmg + 0.5)
        ko = (Fighter.F(d).health <= 0)
        stun = Data.MvHitstun(mv)
        If counter Then stun = stun + 4
        kd = (Data.MvKD(mv) And lastHit) Or ko
        Fighter.ApplyHit d, stun, push, kd
        Fighter.F(d).flash = 2
        If isProj Then
            AddMeter a, 30
        Else
            AddMeter a, 50
        End If
        AddMeter d, 25
        If heavy Or counter Then
            FX.Spawn cx, cy, FX_HEAVY
            HitStop = 4
        Else
            FX.Spawn cx, cy, FX_HIT
            HitStop = 3
        End If
        If counter Then ShowText "COUNTER", 20, a
        If mCombo(a) >= 2 Then
            mComboShown(a) = mCombo(a)
            mComboShow(a) = 40
        End If
        If ko Then HitStop = 10
    End If

    ' a defender pinned to the wall cannot be pushed: push the attacker back instead
    If Not isProj And Fighter.AtWall(d) Then
        Fighter.F(a).slide = -Fighter.F(a).facing * push * 1.2
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

Private Function Max2D(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Max2D = a Else Max2D = b
End Function

Private Function Min2(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then Min2 = a Else Min2 = b
End Function
