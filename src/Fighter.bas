Attribute VB_Name = "Fighter"
Option Explicit

' Fighter state machine.
'   IDLE / WALK / CROUCH / JUMP / ATTACK / GUARD / HIT / DOWN / THROW / THROWN
' Attacks run on frame data from data\moves.txt (startup / active / recovery).
' Combos: a light normal that connected (hit or guard) can be chained into
' another normal (up to MAX_CHAIN). Inputs are buffered for BUF_TICKS ticks.
' Throw: LP+LK together at close range.
' Debug keys: Tab = switch controlled fighter, D = dummy mode (off/guard/attack),
'             H / J = hit / knock down the other fighter, F1 = show boxes.

Public Const ST_IDLE As Long = 0
Public Const ST_WALK As Long = 1
Public Const ST_CROUCH As Long = 2
Public Const ST_JUMP As Long = 3
Public Const ST_ATTACK As Long = 4
Public Const ST_GUARD As Long = 5
Public Const ST_HIT As Long = 6
Public Const ST_DOWN As Long = 7
Public Const ST_THROW As Long = 8
Public Const ST_THROWN As Long = 9

' move slots per fighter
Public Const MV_LP As Long = 0
Public Const MV_HP As Long = 1
Public Const MV_LK As Long = 2
Public Const MV_HK As Long = 3
Public Const MV_CLP As Long = 4
Public Const MV_CHP As Long = 5
Public Const MV_CLK As Long = 6
Public Const MV_CHK As Long = 7
Public Const MV_JP As Long = 8
Public Const MV_JK As Long = 9
Public Const MV_THROW As Long = 10

Private Const GRAVITY As Double = 0.5        ' px / tick^2
Private Const BACK_WALK As Double = 0.75     ' back-walk speed factor
Private Const GUARD_RANGE As Double = 60     ' hold back within this distance of an attack -> guard
Private Const DOWN_LIE_TICKS As Long = 16
Private Const GETUP_TICKS As Long = 4
Private Const DUMMY_INTERVAL As Long = 30
Private Const BUF_TICKS As Long = 5
Private Const MAX_CHAIN As Long = 3

Public Const DUMMY_OFF As Long = 0
Public Const DUMMY_GUARD As Long = 1
Public Const DUMMY_ATTACK As Long = 2

Public Type TFighter
    prefix As String
    state As Long
    stTick As Long           ' ticks since entering the state
    x As Double              ' anchor x (center under hip)
    footY As Double          ' GROUND_Y when on the ground
    vx As Double
    vy As Double
    onGround As Boolean
    crouch As Boolean        ' crouching variant of ATTACK / GUARD / HIT
    backward As Boolean      ' walking away from the opponent
    facing As Long           ' +1 right, -1 left
    mv As Long               ' current move id (ATTACK / THROW)
    hasHit As Boolean        ' current attack already connected
    chain As Long            ' chained normals so far
    airAttacked As Boolean
    stun As Long             ' remaining hit/block stun
    stunTotal As Long
    downPhase As Long        ' 0 airborne, 1 lying, 2 getting up
    flash As Long            ' ticks left of white hit flash
    bufAtk As Long           ' buffered attack slot (-1 = none)
    bufAge As Long
    health As Long
    maxHealth As Long
    healthShown As Double    ' red damage trail on the health bar
    trailDelay As Long
    ' tuning
    walkSpeed As Double
    jumpVY As Double
    jumpVX As Double
    halfW As Double          ' push-box half width
    hurtH As Long            ' hurtbox height standing
    hurtHC As Long           ' hurtbox height crouching
    startX As Double
    ' resources
    aIdle As Long
    aWalk As Long
    sCrouch As Long
    sJump As Long
    sGuard As Long
    sCGuard As Long
    sHit0 As Long
    sHit1 As Long
    sCHit As Long
    sDown0 As Long
    sDown1 As Long
    moves(0 To 10) As Long
End Type

Public F(0 To 1) As TFighter
Private mControl As Long
Private mDummy As Long
Private mDummyTick As Long
Private mDummyNext As Long
Private mShowBoxes As Boolean

Public Sub Init()
    '      i  prefix      x   walk  jumpVY jumpVX halfW  hp    hurtH hurtHC
    Setup 0, "danwoo", 40, 2#, -4.5, 1.8, 6, 1000, 30, 21
    Setup 1, "cheolsan", 80, 1.4, -4#, 1.4, 8, 1150, 34, 25
    mControl = 0
    mDummy = DUMMY_OFF
    mShowBoxes = False
End Sub

Public Property Get Controlled() As Long
    Controlled = mControl
End Property

Public Property Get StatusText() As String
    Dim d As String
    Select Case mDummy
        Case DUMMY_GUARD: d = "guard"
        Case DUMMY_ATTACK: d = "attack"
        Case Else: d = "off"
    End Select
    StatusText = "P1 " & F(mControl).prefix & " (Tab) | dummy " & d & " (D)" & IIf(mShowBoxes, " | boxes (F1)", "")
End Property

' Harness / debug helpers
Public Sub SetControl(ByVal i As Long)
    mControl = i
End Sub

Public Sub SetDummy(ByVal mode As Long)
    mDummy = mode
    mDummyTick = 0
End Sub

Private Sub Setup(ByVal i As Long, ByVal prefix As String, ByVal x As Double, ByVal walkSpeed As Double, _
                  ByVal jumpVY As Double, ByVal jumpVX As Double, ByVal halfW As Double, _
                  ByVal maxHealth As Long, ByVal hurtH As Long, ByVal hurtHC As Long)
    Dim names As Variant, k As Long
    names = Array("lp", "hp", "lk", "hk", "clp", "chp", "clk", "chk", "jp", "jk", "throw")
    With F(i)
        .prefix = prefix
        .startX = x
        .walkSpeed = walkSpeed
        .jumpVY = jumpVY
        .jumpVX = jumpVX
        .halfW = halfW
        .maxHealth = maxHealth
        .hurtH = hurtH
        .hurtHC = hurtHC
        .aIdle = Data.AnimId(prefix & ".idle")
        .aWalk = Data.AnimId(prefix & ".walk")
        .sCrouch = Data.SpriteId(prefix & ".crouch")
        .sJump = Data.SpriteId(prefix & ".jump")
        .sGuard = Data.SpriteId(prefix & ".guard")
        .sCGuard = Data.SpriteId(prefix & ".cguard")
        .sHit0 = Data.SpriteId(prefix & ".hit_0")
        .sHit1 = Data.SpriteId(prefix & ".hit_1")
        .sCHit = Data.SpriteId(prefix & ".chit")
        .sDown0 = Data.SpriteId(prefix & ".down_0")
        .sDown1 = Data.SpriteId(prefix & ".down_1")
        For k = 0 To 10
            .moves(k) = Data.MoveId(prefix & "." & names(k))
        Next k
    End With
    ResetOne i
End Sub

Public Sub ResetRound()
    ResetOne 0
    ResetOne 1
End Sub

Private Sub ResetOne(ByVal i As Long)
    With F(i)
        .x = .startX
        .footY = GROUND_Y
        .vx = 0: .vy = 0
        .onGround = True
        .crouch = False
        .facing = IIf(i = 0, 1, -1)
        .health = .maxHealth
        .healthShown = .maxHealth
        .trailDelay = 0
        .stun = 0
        .flash = 0
        .bufAtk = -1
        .chain = 0
        .state = -1
    End With
    SetState i, ST_IDLE
End Sub

' ======================================================================
' input
' ======================================================================

' Called every tick (also during hit stop) so presses are never lost.
Public Sub BufferInput()
    Dim a As Long
    If F(0).flash > 0 Then F(0).flash = F(0).flash - 1
    If F(1).flash > 0 Then F(1).flash = F(1).flash - 1
    a = -1
    ' throw = LP + LK together (one of them pressed this tick, the other held)
    If KeyInput.IsDown(BTN_LP) And KeyInput.IsDown(BTN_LK) And _
       (KeyInput.IsPressed(BTN_LP) Or KeyInput.IsPressed(BTN_LK)) Then
        a = MV_THROW
    Else
        ' heavier button wins if pressed on the same tick
        If KeyInput.IsPressed(BTN_LP) Then a = MV_LP
        If KeyInput.IsPressed(BTN_LK) Then a = MV_LK
        If KeyInput.IsPressed(BTN_HP) Then a = MV_HP
        If KeyInput.IsPressed(BTN_HK) Then a = MV_HK
    End If
    With F(mControl)
        If a >= 0 Then
            .bufAtk = a
            .bufAge = 0
        ElseIf .bufAtk >= 0 Then
            .bufAge = .bufAge + 1
            If .bufAge > BUF_TICKS Then .bufAtk = -1
        End If
    End With
End Sub

Private Function TakeAttack(ByVal i As Long) As Long
    TakeAttack = F(i).bufAtk
    F(i).bufAtk = -1
End Function

Private Sub ReadDirections(ByVal i As Long, l As Boolean, r As Boolean, u As Boolean, d As Boolean)
    If i = mControl Then
        l = KeyInput.IsDown(BTN_LEFT): r = KeyInput.IsDown(BTN_RIGHT)
        u = KeyInput.IsDown(BTN_UP): d = KeyInput.IsDown(BTN_DOWN)
    ElseIf mDummy = DUMMY_ATTACK Then
        ' training dummy: cycles through its ground normals
        mDummyTick = mDummyTick + 1
        If mDummyTick >= DUMMY_INTERVAL Then
            mDummyTick = 0
            F(i).bufAtk = mDummyNext Mod 4
            F(i).bufAge = 0
            d = (mDummyNext Mod 8) >= 4
            mDummyNext = mDummyNext + 1
        ElseIf F(i).state = ST_ATTACK Then
            d = F(i).crouch
        End If
    End If
End Sub

' ======================================================================
' update
' ======================================================================

Public Sub Update()
    Dim i As Long
    If KeyInput.IsPressed(BTN_SWAP) Then mControl = 1 - mControl: F(0).bufAtk = -1: F(1).bufAtk = -1
    If KeyInput.IsPressed(BTN_DBG_DUMMY) Then SetDummy ((mDummy + 1) Mod 3)
    If KeyInput.IsPressed(BTN_DBG_BOX) Then mShowBoxes = Not mShowBoxes
    If KeyInput.IsPressed(BTN_DBG_HIT) Then ApplyHit 1 - mControl, 10, 1.5, False
    If KeyInput.IsPressed(BTN_DBG_DOWN) Then ApplyHit 1 - mControl, 0, 1.5, True

    For i = 0 To 1
        UpdateFighter i
    Next i
    ResolvePush
    For i = 0 To 1
        UpdateFacing i
    Next i
End Sub

Private Sub UpdateFighter(ByVal i As Long)
    Dim l As Boolean, r As Boolean, u As Boolean, d As Boolean
    Dim fwd As Boolean, back As Boolean, wantGuard As Boolean, atk As Long

    ReadDirections i, l, r, u, d
    With F(i)
        If .facing > 0 Then
            fwd = r: back = l
        Else
            fwd = l: back = r
        End If
        wantGuard = back And Threatened(i)
        If i <> mControl And mDummy = DUMMY_GUARD Then wantGuard = Threatened(i)   ' guard only, never walk away
        .stTick = .stTick + 1

        Select Case .state
        Case ST_IDLE, ST_WALK, ST_CROUCH
            Actionable i, fwd, back, u, d, wantGuard

        Case ST_GUARD
            If .stun > 0 Then
                .stun = .stun - 1
                .x = .x + .vx
                .vx = .vx * 0.7
            Else
                Actionable i, fwd, back, u, d, wantGuard
            End If

        Case ST_JUMP
            atk = TakeAttack(i)
            If atk >= 0 And Not .airAttacked Then
                .airAttacked = True
                If atk = MV_LP Or atk = MV_HP Or atk = MV_THROW Then .mv = .moves(MV_JP) Else .mv = .moves(MV_JK)
                .hasHit = False
                SetState i, ST_ATTACK
            End If
            If AirStep(i) Then SetState i, ST_IDLE

        Case ST_ATTACK
            If Not .onGround Then
                If AirStep(i) Then SetState i, ST_IDLE      ' landing ends an air attack
            Else
                ' chain combo: light normal that connected -> next normal
                If .hasHit And .bufAtk >= 0 And .bufAtk <> MV_THROW And .chain < MAX_CHAIN And IsLight(.mv) Then
                    atk = TakeAttack(i)
                    .chain = .chain + 1
                    StartGroundAttack i, atk, d, True
                ElseIf .stTick >= MoveTotal(.mv) Then
                    If d Then
                        SetState i, ST_CROUCH
                    Else
                        SetState i, ST_IDLE
                    End If
                End If
            End If

        Case ST_HIT
            .x = .x + .vx
            .vx = .vx * 0.7
            .stun = .stun - 1
            If .stun <= 0 Then
                If .crouch Then
                    SetState i, ST_CROUCH
                Else
                    SetState i, ST_IDLE
                End If
            End If

        Case ST_DOWN
            Select Case .downPhase
            Case 0
                If AirStep(i) Then
                    .downPhase = 1
                    .stTick = 0
                End If
            Case 1
                If .stTick >= DOWN_LIE_TICKS And .health > 0 Then        ' KO'd fighters stay down
                    .downPhase = 2
                    .stTick = 0
                End If
            Case Else
                If .stTick >= GETUP_TICKS Then SetState i, ST_IDLE
            End Select

        Case ST_THROW
            If .stTick = Data.MvStartup(.mv) Then
                Combat.DealDamage 1 - i, Data.MvDamage(.mv)
                ApplyHit 1 - i, 0, 2.5, True
                FX.Spawn Int(F(1 - i).x + 0.5), Int(F(1 - i).footY) - 18, FX_HEAVY
                Combat.HitStop = 4
            End If
            If .stTick >= MoveTotal(.mv) Then SetState i, ST_IDLE

        Case ST_THROWN
            .vx = 0                ' held in place by the thrower until the throw lands
        End Select
    End With
    ClampX i
End Sub

' Neutral ground states: attack, jump, guard, crouch, walk.
Private Sub Actionable(ByVal i As Long, ByVal fwd As Boolean, ByVal back As Boolean, _
                       ByVal u As Boolean, ByVal d As Boolean, ByVal wantGuard As Boolean)
    Dim atk As Long
    atk = TakeAttack(i)
    With F(i)
        If atk = MV_THROW Then
            If TryThrow(i) Then Exit Sub
            atk = MV_LP                                   ' whiffed throw becomes a jab
        End If
        If atk >= 0 Then
            StartGroundAttack i, atk, d, False
        ElseIf u Then
            SetState i, ST_JUMP
            .onGround = False
            .airAttacked = False
            .vy = .jumpVY
            If fwd Then
                .vx = .facing * .jumpVX
            ElseIf back Then
                .vx = -.facing * .jumpVX
            Else
                .vx = 0
            End If
        ElseIf wantGuard Then
            .crouch = d
            SetState i, ST_GUARD
        ElseIf d Then
            SetState i, ST_CROUCH
        ElseIf fwd Or back Then
            .backward = back
            SetState i, ST_WALK
            If back Then .x = .x - .facing * .walkSpeed * BACK_WALK Else .x = .x + .facing * .walkSpeed
        Else
            SetState i, ST_IDLE
        End If
    End With
End Sub

Private Sub StartGroundAttack(ByVal i As Long, ByVal atk As Long, ByVal crouching As Boolean, ByVal chained As Boolean)
    With F(i)
        If atk = MV_THROW Then atk = MV_LP
        If crouching Then .mv = .moves(atk + 4) Else .mv = .moves(atk)
        .crouch = crouching
        .vx = 0
        .hasHit = False
        If Not chained Then .chain = 0
        SetState i, ST_ATTACK
    End With
End Sub

Private Function TryThrow(ByVal i As Long) As Boolean
    Dim o As Long, mv As Long, reach As Double
    o = 1 - i
    mv = F(i).moves(MV_THROW)
    reach = Data.MvHbX(mv) + Data.MvHbW(mv)
    If Not F(o).onGround Then Exit Function
    If Abs(F(o).x - F(i).x) > reach Then Exit Function
    Select Case F(o).state
    Case ST_IDLE, ST_WALK, ST_CROUCH
    Case ST_GUARD
        If F(o).stun > 0 Then Exit Function
    Case ST_ATTACK
        If F(o).stTick < Data.MvStartup(F(o).mv) + Data.MvActive(F(o).mv) Then Exit Function   ' only in recovery
    Case Else
        Exit Function
    End Select
    F(i).mv = mv
    F(i).vx = 0
    SetState i, ST_THROW
    SetState o, ST_THROWN
    F(o).crouch = False
    F(o).x = F(i).x + F(i).facing * (F(i).halfW + F(o).halfW - 2)
    TryThrow = True
End Function

Private Function IsLight(ByVal mv As Long) As Boolean
    Dim k As Long
    For k = 0 To 1
        If mv = F(k).moves(MV_LP) Or mv = F(k).moves(MV_LK) Or mv = F(k).moves(MV_CLP) Or mv = F(k).moves(MV_CLK) Then
            IsLight = True
            Exit Function
        End If
    Next k
End Function

Private Function MoveTotal(ByVal mv As Long) As Long
    MoveTotal = Data.MvStartup(mv) + Data.MvActive(mv) + Data.MvRecovery(mv)
End Function

' Is the opponent attacking close enough that holding back should guard?
Private Function Threatened(ByVal i As Long) As Boolean
    Dim o As Long
    o = 1 - i
    Threatened = (F(o).state = ST_ATTACK) And (Abs(F(o).x - F(i).x) < GUARD_RANGE)
End Function

' Ballistic step. Returns True on the tick it lands.
Private Function AirStep(ByVal i As Long) As Boolean
    With F(i)
        .vy = .vy + GRAVITY
        .footY = .footY + .vy
        .x = .x + .vx
        If .footY >= GROUND_Y Then
            .footY = GROUND_Y
            .vy = 0
            .vx = 0
            .onGround = True
            AirStep = True
        End If
    End With
End Function

Private Sub SetState(ByVal i As Long, ByVal st As Long)
    With F(i)
        If .state <> st Or st = ST_ATTACK Or st = ST_HIT Or st = ST_DOWN Or st = ST_THROW Then .stTick = 0
        .state = st
        If st = ST_CROUCH Then .crouch = True
        If st = ST_IDLE Or st = ST_WALK Or st = ST_JUMP Then .crouch = False
    End With
End Sub

' Called by combat and the debug keys.
Public Sub ApplyHit(ByVal i As Long, ByVal hitstun As Long, ByVal push As Double, ByVal knockdown As Boolean)
    With F(i)
        .bufAtk = -1
        If knockdown Or Not .onGround Then
            SetState i, ST_DOWN
            .downPhase = 0
            .onGround = False
            .vy = -3.5
            .vx = -.facing * push
        Else
            .crouch = (.state = ST_CROUCH) Or ((.state = ST_ATTACK Or .state = ST_GUARD Or .state = ST_HIT) And .crouch)
            SetState i, ST_HIT
            .stun = hitstun
            .stunTotal = hitstun
            .vx = -.facing * push
        End If
    End With
End Sub

Public Sub ApplyBlock(ByVal i As Long, ByVal blockstun As Long, ByVal push As Double)
    With F(i)
        SetState i, ST_GUARD
        .stun = blockstun
        .vx = -.facing * push
    End With
End Sub

Private Sub ClampX(ByVal i As Long)
    With F(i)
        If .x < .halfW Then .x = .halfW
        If .x > SCREEN_W - .halfW Then .x = SCREEN_W - .halfW
    End With
End Sub

' Keep the two bodies from overlapping (only when vertically close; not while lying down).
Private Sub ResolvePush()
    Dim dx As Double, minD As Double, ov As Double, s As Double
    If Abs(F(0).footY - F(1).footY) > 14 Then Exit Sub
    If F(0).state = ST_DOWN Or F(1).state = ST_DOWN Then Exit Sub
    If F(0).state = ST_THROWN Or F(1).state = ST_THROWN Then Exit Sub
    dx = F(1).x - F(0).x
    minD = F(0).halfW + F(1).halfW
    If Abs(dx) >= minD Then Exit Sub
    If dx > 0 Then
        s = 1
    ElseIf dx < 0 Then
        s = -1
    Else
        s = F(0).facing
    End If
    ov = minD - Abs(dx)
    F(0).x = F(0).x - s * ov / 2
    F(1).x = F(1).x + s * ov / 2
    ClampX 0: ClampX 1
    ' if one was pinned to the wall, push the other the rest of the way
    dx = F(1).x - F(0).x
    If Abs(dx) < minD Then
        If F(0).x <= F(0).halfW Or F(0).x >= SCREEN_W - F(0).halfW Then
            F(1).x = F(0).x + s * minD
            ClampX 1
        Else
            F(0).x = F(1).x - s * minD
            ClampX 0
        End If
    End If
End Sub

Private Sub UpdateFacing(ByVal i As Long)
    Dim o As Long
    o = 1 - i
    With F(i)
        Select Case .state
        Case ST_IDLE, ST_WALK, ST_CROUCH, ST_GUARD
            If F(o).x > .x Then
                .facing = 1
            ElseIf F(o).x < .x Then
                .facing = -1
            End If
        End Select
    End With
End Sub

' ======================================================================
' boxes (used by Combat)
' ======================================================================

' True while the current attack is in its active frames.
Public Function InActive(ByVal i As Long) As Boolean
    Dim t As Long, mv As Long
    If F(i).state <> ST_ATTACK Then Exit Function
    t = F(i).stTick: mv = F(i).mv
    InActive = (t >= Data.MvStartup(mv)) And (t < Data.MvStartup(mv) + Data.MvActive(mv))
End Function

Public Function HitRect(ByVal i As Long, x0 As Long, y0 As Long, w As Long, h As Long) As Boolean
    Dim mv As Long
    If Not InActive(i) Then Exit Function
    With F(i)
        mv = .mv
        w = Data.MvHbW(mv): h = Data.MvHbH(mv)
        If .facing > 0 Then
            x0 = Int(.x + 0.5) + Data.MvHbX(mv)
        Else
            x0 = Int(.x + 0.5) - Data.MvHbX(mv) - w + 1
        End If
        y0 = Int(.footY + 0.5) + Data.MvHbY(mv)
    End With
    HitRect = True
End Function

Public Function HurtRect(ByVal i As Long, x0 As Long, y0 As Long, w As Long, h As Long) As Boolean
    With F(i)
        If .state = ST_DOWN Or .state = ST_THROWN Then Exit Function     ' invulnerable
        If .crouch Then h = .hurtHC Else h = .hurtH
        w = Int(.halfW * 2 + 4)
        x0 = Int(.x + 0.5) - w \ 2
        y0 = Int(.footY + 0.5) - h
    End With
    HurtRect = True
End Function

' ======================================================================
' draw
' ======================================================================

Public Sub Draw()
    DrawOne 1 - mControl
    DrawOne mControl
    If mShowBoxes Then
        DrawBoxes 0
        DrawBoxes 1
    End If
End Sub

Private Function CurrentSprite(ByVal i As Long) As Long
    Dim t As Long
    With F(i)
        t = .stTick
        Select Case .state
        Case ST_IDLE
            CurrentSprite = Data.AnimFrame(.aIdle, t)
        Case ST_WALK
            CurrentSprite = Data.AnimFrame(.aWalk, t, .backward)
        Case ST_CROUCH
            CurrentSprite = .sCrouch
        Case ST_JUMP
            CurrentSprite = .sJump
        Case ST_ATTACK, ST_THROW
            If .state = ST_THROW Then
                If t >= Data.MvStartup(.mv) Then CurrentSprite = Data.MvSprA(.mv) Else CurrentSprite = Data.MvSprS(.mv)
            ElseIf InActive(i) Then
                CurrentSprite = Data.MvSprA(.mv)
            Else
                CurrentSprite = Data.MvSprS(.mv)
            End If
        Case ST_GUARD
            If .crouch Then CurrentSprite = .sCGuard Else CurrentSprite = .sGuard
        Case ST_HIT
            If .crouch Then
                CurrentSprite = .sCHit
            ElseIf t < (.stunTotal + 1) \ 2 Then
                CurrentSprite = .sHit0
            Else
                CurrentSprite = .sHit1
            End If
        Case ST_THROWN
            CurrentSprite = .sHit0
        Case ST_DOWN
            Select Case .downPhase
            Case 0: CurrentSprite = .sDown0
            Case 1: CurrentSprite = .sDown1
            Case Else: CurrentSprite = .sCrouch
            End Select
        Case Else
            CurrentSprite = Data.AnimFrame(.aIdle, t)
        End Select
    End With
End Function

Private Sub DrawOne(ByVal i As Long)
    Dim tint As Long
    tint = -1
    If F(i).flash > 0 Then tint = RGB(255, 255, 255)
    With F(i)
        Renderer.DrawSprite CurrentSprite(i), Int(.x + 0.5), Int(.footY + 0.5), (.facing < 0), tint
    End With
End Sub

' Debug: hurtbox (green) and, during active frames, the hitbox (red).
Private Sub DrawBoxes(ByVal i As Long)
    Dim x0 As Long, y0 As Long, w As Long, h As Long
    If HurtRect(i, x0, y0, w, h) Then Renderer.DrawRectOutline x0, y0, w, h, RGB(60, 255, 120)
    If HitRect(i, x0, y0, w, h) Then Renderer.DrawRectOutline x0, y0, w, h, RGB(255, 40, 40)
End Sub
