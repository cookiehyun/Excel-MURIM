Attribute VB_Name = "Fighter"
Option Explicit

' Fighter state machine.
'   IDLE / WALK / CROUCH / JUMP / ATTACK / GUARD / HIT / DOWN / THROW / THROWN / DASH / LAND
' Input: every tick each fighter receives a direction (numpad, relative to facing:
' 6 = forward, 4 = back, 2 = down, 8 = up) plus pressed/held button masks, from the
' keyboard (FeedHuman), the CPU (CpuAI -> FeedInput) or the training dummy.
' A 32-tick direction history drives command specials and dashes.
' Moves and their frame data come from data\moves.txt.

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
Public Const ST_DASH As Long = 10
Public Const ST_LAND As Long = 11

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
Public Const MV_SP1 As Long = 11
Public Const MV_SP2 As Long = 12
Public Const MV_SUPER As Long = 13

' button masks
Public Const B_LP As Long = 1
Public Const B_HP As Long = 2
Public Const B_LK As Long = 4
Public Const B_HK As Long = 8

Public Const METER_MAX As Long = 1000

Public Const DUMMY_CPU As Long = 0
Public Const DUMMY_IDLE As Long = 1
Public Const DUMMY_GUARD As Long = 2
Public Const DUMMY_ATTACK As Long = 3

Private Const GRAVITY As Double = 0.5        ' px / tick^2
Private Const BACK_WALK As Double = 0.75     ' back-walk speed factor
Private Const GUARD_RANGE As Double = 60     ' hold back within this distance of an attack -> guard
Private Const DOWN_LIE_TICKS As Long = 16
Private Const GETUP_TICKS As Long = 4
Private Const DUMMY_INTERVAL As Long = 30
Private Const BUF_TICKS As Long = 5
Private Const MAX_CHAIN As Long = 3
Private Const HIST As Long = 32
Private Const MOTION_WINDOW As Long = 12
Private Const SUPER_WINDOW As Long = 20
Private Const DASH_TICKS As Long = 6
Private Const TECH_WINDOW As Long = 3

Public Type TFighter
    prefix As String
    state As Long
    stTick As Long           ' ticks since entering the state
    x As Double              ' anchor x (center under hip)
    footY As Double          ' GROUND_Y when on the ground
    vx As Double
    vy As Double
    slide As Double          ' extra horizontal push (corner pushback), decays
    onGround As Boolean
    crouch As Boolean        ' crouching variant of ATTACK / GUARD / HIT
    backward As Boolean      ' walking / dashing away from the opponent
    facing As Long           ' +1 right, -1 left
    mv As Long               ' current move id (ATTACK / THROW)
    hasHit As Boolean        ' current attack already connected
    chain As Long            ' chained normals so far
    airAttacked As Boolean
    stun As Long             ' remaining hit/block stun, or landing lag
    stunTotal As Long
    downPhase As Long        ' 0 airborne, 1 lying, 2 getting up
    flash As Long            ' ticks left of white hit flash
    bufAtk As Long           ' buffered attack slot (-1 = none)
    bufAge As Long
    bufDash As Long          ' +1 forward dash, -1 back dash, 0 none
    health As Long
    maxHealth As Long
    healthShown As Double    ' red damage trail on the health bar
    trailDelay As Long
    meter As Long            ' super gauge 0..METER_MAX
    ' input this tick
    inDir As Long
    inPress As Long
    inHeld As Long
    hist(0 To 31) As Long
    histPos As Long
    ' tuning
    walkSpeed As Double
    dashSpeed As Double
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
    sDash As Long
    sBDash As Long
    sPunchA As Long          ' super grab sequence poses
    sPunchB As Long
    sLift As Long
    sSlam As Long
    sSlashA As Long          ' super slash sequence poses
    sSlashB As Long
    sSlashC As Long
    seqFx As Long            ' effect used by the super sequence hits
    superName As String
    hitCount As Long         ' hits landed by the current multi-hit rush
    moves(0 To 13) As Long
    spOrder(0 To 2) As Long  ' special slots sorted by command length (longest first)
End Type

Public F(0 To 1) As TFighter
Private mControl As Long
Private mDummy As Long
Private mDummyTick As Long
Private mDummyNext As Long
Private mShowBoxes As Boolean

' c0 / c1 = character names ("danwoo", "cheolsan"); the same character may appear twice.
Public Sub Init(ByVal c0 As String, ByVal c1 As String, ByVal humanIndex As Long)
    SetupCharacter 0, c0, 72
    SetupCharacter 1, c1, 120
    mControl = humanIndex
    mDummy = DUMMY_CPU
    mShowBoxes = False
End Sub

Public Property Get Controlled() As Long
    Controlled = mControl
End Property

Public Property Get DummyMode() As Long
    DummyMode = mDummy
End Property

Public Property Get StatusText() As String
    Dim d As String
    Select Case mDummy
        Case DUMMY_CPU: d = "CPU"
        Case DUMMY_IDLE: d = "dummy idle"
        Case DUMMY_GUARD: d = "dummy guard"
        Case Else: d = "dummy attack"
    End Select
    StatusText = "P1 " & F(mControl).prefix & " (Tab) | opponent " & d & " (D)" & IIf(mShowBoxes, " | boxes (F1)", "")
End Property

' Harness / debug helpers
Public Sub SetControl(ByVal i As Long)
    mControl = i
End Sub

Public Sub SetDummy(ByVal mode As Long)
    mDummy = mode
    mDummyTick = 0
End Sub

' Per-character tuning (movement, size, health).
Private Sub SetupCharacter(ByVal i As Long, ByVal prefix As String, ByVal x As Double)
    Select Case prefix
    Case "cheolsan"
        '         prefix      x  walk dash  jumpVY jumpVX halfW  hp    hurtH hurtHC
        Setup i, prefix, x, 1.4, 2.6, -4#, 1.4, 8, 1060, 34, 25
    Case "muyeong"
        Setup i, prefix, x, 2#, 3#, -4.4, 1.7, 6, 980, 30, 21
    Case "seolhwa"
        Setup i, prefix, x, 2.1, 3.3, -4.8, 2#, 5, 850, 29, 20
    Case "heukrim"
        Setup i, prefix, x, 1.8, 2.8, -4.2, 1.5, 6, 1000, 31, 22
    Case "hongryeon"
        Setup i, prefix, x, 2#, 3.2, -4.6, 1.8, 5, 980, 29, 20
    Case Else
        Setup i, "danwoo", x, 2#, 3.2, -4.5, 1.8, 6, 1000, 30, 21
    End Select
    With F(i)
        .seqFx = FX_SLASH
        Select Case .prefix
        Case "cheolsan": .superName = "NAHAN BOKMA"
        Case "muyeong": .superName = "MANGEOM GWIJONG"
        Case "seolhwa"
            .superName = "NAKHWA YUSU"
            .seqFx = FX_HIT                              ' kicks instead of cuts
            .sSlashA = Data.SpriteId("seolhwa.hk_1")
            .sSlashB = Data.SpriteId("seolhwa.lk_1")
            .sSlashC = Data.SpriteId("seolhwa.rush_1")
        Case "heukrim"
            .superName = "CHEONMA GANGRIM"
            .seqFx = FX_DARK
        Case "hongryeon": .superName = "HONGRYEON EOPHWA"
        Case Else: .superName = "CHEONPUNG MANJANG"
        End Select
    End With
End Sub

Private Sub Setup(ByVal i As Long, ByVal prefix As String, ByVal x As Double, ByVal walkSpeed As Double, _
                  ByVal dashSpeed As Double, ByVal jumpVY As Double, ByVal jumpVX As Double, ByVal halfW As Double, _
                  ByVal maxHealth As Long, ByVal hurtH As Long, ByVal hurtHC As Long)
    Dim names As Variant, k As Long, a As Long, b As Long, t As Long
    names = Array("lp", "hp", "lk", "hk", "clp", "chp", "clk", "chk", "jp", "jk", "throw", "sp1", "sp2", "super")
    With F(i)
        .prefix = prefix
        .startX = x
        .walkSpeed = walkSpeed
        .dashSpeed = dashSpeed
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
        .sDash = Data.SpriteId(prefix & ".dash")
        .sBDash = Data.SpriteId(prefix & ".bdash")
        .sPunchA = Data.SpriteId(prefix & ".lp_1")
        .sPunchB = Data.SpriteId(prefix & ".hp_1")
        .sLift = Data.SpriteId(prefix & ".throw_1")
        .sSlam = Data.SpriteId(prefix & ".sp_1")
        .sSlashA = Data.SpriteId(prefix & ".hp_1")
        .sSlashB = Data.SpriteId(prefix & ".chp_1")
        .sSlashC = Data.SpriteId(prefix & ".lp_1")
        For k = 0 To 13
            .moves(k) = Data.MoveId(prefix & "." & names(k))
        Next k
        ' specials ordered by command length so e.g. 623 is tested before 26
        .spOrder(0) = MV_SUPER: .spOrder(1) = MV_SP1: .spOrder(2) = MV_SP2
        For a = 0 To 1
            For b = a + 1 To 2
                If Len(Data.MvCmd(.moves(.spOrder(b)))) > Len(Data.MvCmd(.moves(.spOrder(a)))) Then
                    t = .spOrder(a): .spOrder(a) = .spOrder(b): .spOrder(b) = t
                End If
            Next b
        Next a
    End With
    ResetOne i
End Sub

Public Sub ResetRound(Optional ByVal keepMeter As Boolean = False)
    Dim m0 As Long, m1 As Long
    m0 = F(0).meter: m1 = F(1).meter
    ResetOne 0
    ResetOne 1
    If keepMeter Then F(0).meter = m0: F(1).meter = m1
End Sub

Private Sub ResetOne(ByVal i As Long)
    Dim k As Long
    With F(i)
        .x = .startX
        .footY = GROUND_Y
        .vx = 0: .vy = 0: .slide = 0
        .onGround = True
        .crouch = False
        .facing = IIf(i = 0, 1, -1)
        .health = .maxHealth
        .healthShown = .maxHealth
        .trailDelay = 0
        .stun = 0
        .flash = 0
        .bufAtk = -1
        .bufDash = 0
        .chain = 0
        .meter = 0
        .inDir = 5: .inPress = 0: .inHeld = 0
        For k = 0 To HIST - 1
            .hist(k) = 5
        Next k
        .state = -1
    End With
    SetState i, ST_IDLE
End Sub

' ======================================================================
' input
' ======================================================================

Public Function DirFrom(ByVal fwd As Boolean, ByVal back As Boolean, ByVal up As Boolean, ByVal down As Boolean) As Long
    Dim d As Long
    d = 5
    If fwd Then d = d + 1
    If back And Not fwd Then d = d - 1
    If up Then d = d + 3
    If down And Not up Then d = d - 3
    DirFrom = d
End Function

' Keyboard -> relative input for fighter i.
Public Sub FeedHuman(ByVal i As Long)
    Dim l As Boolean, r As Boolean, press As Long, held As Long, d As Long
    l = KeyInput.IsDown(BTN_LEFT): r = KeyInput.IsDown(BTN_RIGHT)
    If F(i).facing > 0 Then
        d = DirFrom(r, l, KeyInput.IsDown(BTN_UP), KeyInput.IsDown(BTN_DOWN))
    Else
        d = DirFrom(l, r, KeyInput.IsDown(BTN_UP), KeyInput.IsDown(BTN_DOWN))
    End If
    If KeyInput.IsPressed(BTN_LP) Then press = press Or B_LP
    If KeyInput.IsPressed(BTN_HP) Then press = press Or B_HP
    If KeyInput.IsPressed(BTN_LK) Then press = press Or B_LK
    If KeyInput.IsPressed(BTN_HK) Then press = press Or B_HK
    If KeyInput.IsDown(BTN_LP) Then held = held Or B_LP
    If KeyInput.IsDown(BTN_HP) Then held = held Or B_HP
    If KeyInput.IsDown(BTN_LK) Then held = held Or B_LK
    If KeyInput.IsDown(BTN_HK) Then held = held Or B_HK
    FeedInput i, d, press, held
End Sub

' Training dummy (when the CPU is switched off with D).
Public Sub FeedDummy(ByVal i As Long)
    Dim d As Long, press As Long
    d = 5
    If mDummy = DUMMY_GUARD Then
        If Threatened(i) Then d = 4
    ElseIf mDummy = DUMMY_ATTACK Then
        mDummyTick = mDummyTick + 1
        If mDummyTick >= DUMMY_INTERVAL Then
            mDummyTick = 0
            press = 2 ^ (mDummyNext Mod 4)
            If (mDummyNext Mod 8) >= 4 Then d = 2
            mDummyNext = mDummyNext + 1
        ElseIf F(i).state = ST_ATTACK And F(i).crouch Then
            d = 2
        End If
    End If
    FeedInput i, d, press, press
End Sub

' Common entry: record history, decode buffered attacks and dashes.
Public Sub FeedInput(ByVal i As Long, ByVal d As Long, ByVal press As Long, ByVal held As Long)
    Dim a As Long
    With F(i)
        .inDir = d: .inPress = press: .inHeld = held
        .histPos = (.histPos + 1) Mod HIST
        .hist(.histPos) = d
        If .flash > 0 Then .flash = .flash - 1
        a = DecodeAttack(i)
        If a >= 0 Then
            .bufAtk = a
            .bufAge = 0
        ElseIf .bufAtk >= 0 Then
            .bufAge = .bufAge + 1
            If .bufAge > BUF_TICKS Then .bufAtk = -1
        End If
        If Tapped(i, 6) Then .bufDash = 1
        If Tapped(i, 4) Then .bufDash = -1
        If d <> 6 And d <> 4 Then .bufDash = 0
    End With
End Sub

' Direction t ticks ago (0 = this tick).
Private Function HistAt(ByVal i As Long, ByVal t As Long) As Long
    HistAt = F(i).hist((F(i).histPos - t + HIST * 4) Mod HIST)
End Function

' Double tap: d now (new press), released, and d tapped again within DASH window.
Private Function Tapped(ByVal i As Long, ByVal d As Long) As Boolean
    Dim t As Long
    If HistAt(i, 0) <> d Or HistAt(i, 1) = d Then Exit Function
    For t = 2 To 6
        If HistAt(i, t) = d Then
            Tapped = (HistAt(i, t + 1) <> d)
            Exit Function
        End If
    Next t
End Function

Private Function DirMatches(ByVal want As Long, ByVal got As Long) As Boolean
    If want = 2 Then
        DirMatches = (got = 1 Or got = 2 Or got = 3)       ' "down" is lenient
    Else
        DirMatches = (want = got)
    End If
End Function

' Motion like "623" found in order in the recent history; last step within 4 ticks.
Private Function MatchMotion(ByVal i As Long, ByVal dirs As String, ByVal window As Long) As Boolean
    Dim j As Long, t As Long, firstHit As Long
    j = Len(dirs)
    firstHit = -1
    For t = 0 To window
        If DirMatches(CLng(Mid$(dirs, j, 1)), HistAt(i, t)) Then
            If firstHit < 0 Then firstHit = t
            j = j - 1
            If j = 0 Then
                MatchMotion = (firstHit <= 4)
                Exit Function
            End If
        End If
    Next t
End Function

Private Function DecodeAttack(ByVal i As Long) As Long
    Dim press As Long, isP As Boolean, isK As Boolean, k As Long, slot As Long, cmd As String, btn As String
    DecodeAttack = -1
    With F(i)
        press = .inPress
        If press = 0 Then Exit Function
        ' throw = LP + LK together
        If (.inHeld And (B_LP Or B_LK)) = (B_LP Or B_LK) And (press And (B_LP Or B_LK)) <> 0 Then
            DecodeAttack = MV_THROW
            Exit Function
        End If
        isP = (press And (B_LP Or B_HP)) <> 0
        isK = (press And (B_LK Or B_HK)) <> 0
        For k = 0 To 2
            slot = .spOrder(k)
            cmd = Data.MvCmd(.moves(slot))
            If Len(cmd) >= 2 Then
                btn = Right$(cmd, 1)
                If (btn = "P" And isP) Or (btn = "K" And isK) Then
                    If slot = MV_SUPER Then
                        If .meter >= METER_MAX Then
                            If MatchMotion(i, Left$(cmd, Len(cmd) - 1), SUPER_WINDOW) Then
                                DecodeAttack = slot
                                Exit Function
                            End If
                        End If
                    ElseIf MatchMotion(i, Left$(cmd, Len(cmd) - 1), MOTION_WINDOW) Then
                        DecodeAttack = slot
                        Exit Function
                    End If
                End If
            End If
        Next k
        ' plain normals: heavier button wins if pressed on the same tick
        If (press And B_LP) <> 0 Then DecodeAttack = MV_LP
        If (press And B_LK) <> 0 Then DecodeAttack = MV_LK
        If (press And B_HP) <> 0 Then DecodeAttack = MV_HP
        If (press And B_HK) <> 0 Then DecodeAttack = MV_HK
    End With
End Function

Private Function TakeAttack(ByVal i As Long) As Long
    TakeAttack = F(i).bufAtk
    F(i).bufAtk = -1
End Function

' ======================================================================
' update
' ======================================================================

Public Sub Update()
    Dim i As Long
    If KeyInput.IsPressed(BTN_SWAP) Then mControl = 1 - mControl: F(0).bufAtk = -1: F(1).bufAtk = -1
    If KeyInput.IsPressed(BTN_DBG_DUMMY) Then SetDummy ((mDummy + 1) Mod 4)
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
    Dim fwd As Boolean, back As Boolean, u As Boolean, d As Boolean, wantGuard As Boolean, atk As Long, mt As Long

    With F(i)
        fwd = (.inDir = 3 Or .inDir = 6 Or .inDir = 9)
        back = (.inDir = 1 Or .inDir = 4 Or .inDir = 7)
        u = (.inDir >= 7)
        d = (.inDir <= 3)
        wantGuard = back And Threatened(i)
        .stTick = .stTick + 1
        If .slide <> 0 Then
            .x = .x + .slide
            .slide = .slide * 0.6
            If Abs(.slide) < 0.1 Then .slide = 0
        End If

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

        Case ST_DASH
            .x = .x + .vx
            If .backward Then
                .footY = GROUND_Y - 3 * Sin(3.14159 * .stTick / DASH_TICKS)     ' small hop
            End If
            If .stTick >= DASH_TICKS Then
                .footY = GROUND_Y
                SetState i, ST_IDLE
            End If

        Case ST_JUMP
            atk = TakeAttack(i)
            If atk >= 0 And Not .airAttacked Then
                .airAttacked = True
                If atk = MV_LK Or atk = MV_HK Then .mv = .moves(MV_JK) Else .mv = .moves(MV_JP)
                .hasHit = False
                SetState i, ST_ATTACK
            End If
            If AirStep(i) Then SetState i, ST_IDLE

        Case ST_ATTACK
            UpdateAttack i, d

        Case ST_LAND
            .stun = .stun - 1
            If .stun <= 0 Then SetState i, ST_IDLE

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
            If Data.MvType(.mv) = MT_SUPER_GRAB Then
                SuperGrabStep i
            ElseIf Data.MvType(.mv) = MT_SUPER_RUSH Or Data.MvType(.mv) = MT_SUPER_TELE Then
                SuperSlashStep i
            ElseIf .stTick = Data.MvStartup(.mv) Then
                Combat.DealDamage 1 - i, Data.MvDamage(.mv)
                Combat.AddMeter i, 80
                ApplyHit 1 - i, 0, 2.5, True
                FX.Spawn Int(F(1 - i).x + 0.5), Int(F(1 - i).footY) - 18, FX_HEAVY
                Combat.HitStop = 4
            End If
            If Data.MvType(.mv) = MT_THROW Or Data.MvType(.mv) = MT_GRAB Then
                If .stTick >= MoveTotal(.mv) Then SetState i, ST_IDLE
            End If

        Case ST_THROWN
            .vx = 0
            ' throw escape: press throw right after being grabbed (normal throws only)
            If .stTick <= TECH_WINDOW And .bufAtk = MV_THROW Then
                If Data.MvType(F(1 - i).mv) = MT_THROW Then TechThrow i
            End If
        End Select
    End With
    ClampX i
End Sub

Private Sub UpdateAttack(ByVal i As Long, ByVal d As Boolean)
    Dim mt As Long, atk As Long, t As Long
    With F(i)
        mt = Data.MvType(.mv)
        t = .stTick

        ' per-type behaviour
        Select Case mt
        Case MT_PROJ, MT_SUPER_PROJ, MT_PROJ_BOOM, MT_SUPER_FIRE
            If t = Data.MvStartup(.mv) Then
                Projectile.Spawn i, .mv
                If mt = MT_SUPER_FIRE Then Combat.ShowText .superName, 40, i
            End If
        Case MT_RISE
            If t = Data.MvStartup(.mv) Then
                .onGround = False
                .vy = Data.MvP1(.mv)
                .vx = .facing * 0.8
            End If
        Case MT_RUSH, MT_SUPER_RUSH
            If t >= Data.MvStartup(.mv) And t < Data.MvStartup(.mv) + Data.MvActive(.mv) Then
                .x = .x + .facing * Data.MvP1(.mv)
                ' multi-hit rush (p2 = hits): allow a new hit every active / hits ticks
                If mt = MT_RUSH And Data.MvP2(.mv) > 1 And .hasHit And .hitCount < Data.MvP2(.mv) Then
                    If (t - Data.MvStartup(.mv)) Mod (Data.MvActive(.mv) \ Data.MvP2(.mv)) = 0 Then .hasHit = False
                End If
            End If
        Case MT_GRAB, MT_SUPER_GRAB
            If t = Data.MvStartup(.mv) Then
                If TryGrab(i, .mv) Then Exit Sub
            End If
        Case MT_TELEPORT, MT_SUPER_TELE
            If t = 1 Then FX.Spawn Int(.x + 0.5), Int(.footY) - 16, FX_DARK
            If t = Data.MvStartup(.mv) - 1 Then
                Teleport i
                FX.Spawn Int(.x + 0.5), Int(.footY) - 16, FX_DARK
            End If
        End Select

        If Not .onGround Then
            If AirStep(i) Then
                If mt = MT_RISE Then
                    SetState i, ST_LAND
                    .stun = Data.MvRecovery(.mv)
                Else
                    SetState i, ST_IDLE                     ' landing ends an air attack
                End If
            End If
            Exit Sub
        End If

        ' cancels after the attack connected (hit or guard)
        If .hasHit And .bufAtk >= 0 And .bufAtk <> MV_THROW Then
            If .bufAtk >= MV_SP1 And CanCancelIntoSpecial(i, .bufAtk) Then
                atk = TakeAttack(i)
                StartAttack i, atk, d, True
                Exit Sub
            ElseIf .bufAtk < MV_SP1 And .chain < MAX_CHAIN And IsLight(.mv) Then
                atk = TakeAttack(i)
                .chain = .chain + 1
                StartAttack i, atk, d, True
                Exit Sub
            End If
        End If

        If t >= MoveTotal(.mv) Then
            If d Then
                SetState i, ST_CROUCH
            Else
                SetState i, ST_IDLE
            End If
        End If
    End With
End Sub

' normal -> special, special -> super
Private Function CanCancelIntoSpecial(ByVal i As Long, ByVal slot As Long) As Boolean
    Dim cur As Long
    cur = Data.MvType(F(i).mv)
    If slot = MV_SUPER Then
        CanCancelIntoSpecial = (cur <> MT_SUPER_PROJ And cur <> MT_SUPER_GRAB)
    Else
        CanCancelIntoSpecial = (cur = MT_NORMAL)
    End If
End Function

' Neutral ground states: attack, jump, dash, guard, crouch, walk.
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
            If StartAttack(i, atk, d, False) Then Exit Sub
        End If
        If .bufDash <> 0 Then
            .backward = (.bufDash < 0)
            .vx = .facing * .dashSpeed * .bufDash
            .bufDash = 0
            SetState i, ST_DASH
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

' Returns False if the move cannot start (e.g. projectile already on screen).
Private Function StartAttack(ByVal i As Long, ByVal atk As Long, ByVal crouching As Boolean, ByVal chained As Boolean) As Boolean
    Dim mv As Long, mt As Long
    With F(i)
        If atk = MV_THROW Then atk = MV_LP
        If atk >= MV_SP1 Then
            mv = .moves(atk)
            mt = Data.MvType(mv)
            If (mt = MT_PROJ Or mt = MT_SUPER_PROJ Or mt = MT_PROJ_BOOM Or mt = MT_SUPER_FIRE) And Projectile.Alive(i) Then Exit Function
            If atk = MV_SUPER Then
                If .meter < METER_MAX Then Exit Function
                .meter = 0
                Combat.StartSuper i
            End If
            crouching = False
        ElseIf crouching Then
            mv = .moves(atk + 4)
        Else
            mv = .moves(atk)
        End If
        .mv = mv
        .crouch = crouching
        .vx = 0
        .hasHit = False
        .hitCount = 0
        If Not chained Then .chain = 0
        SetState i, ST_ATTACK
    End With
    StartAttack = True
End Function

Private Function CanBeGrabbed(ByVal o As Long) As Boolean
    If Not F(o).onGround Then Exit Function
    Select Case F(o).state
    Case ST_IDLE, ST_WALK, ST_CROUCH, ST_DASH, ST_LAND
        CanBeGrabbed = True
    Case ST_GUARD
        CanBeGrabbed = (F(o).stun = 0)
    Case ST_ATTACK
        ' only during recovery (attacks beat throws)
        CanBeGrabbed = (F(o).stTick >= Data.MvStartup(F(o).mv) + Data.MvActive(F(o).mv))
    End Select
End Function

Private Function TryThrow(ByVal i As Long) As Boolean
    TryThrow = TryGrab(i, F(i).moves(MV_THROW))
End Function

Private Function TryGrab(ByVal i As Long, ByVal mv As Long) As Boolean
    Dim o As Long, reach As Double
    o = 1 - i
    reach = Data.MvHbX(mv) + Data.MvHbW(mv)
    If Abs(F(o).x - F(i).x) > reach Then Exit Function
    If Not CanBeGrabbed(o) Then Exit Function
    F(i).mv = mv
    F(i).vx = 0
    SetState i, ST_THROW
    SetState o, ST_THROWN
    F(o).crouch = False
    F(o).bufAtk = -1
    If Data.MvType(mv) <> MT_SUPER_GRAB Then F(o).x = F(i).x + F(i).facing * (F(i).halfW + F(o).halfW - 2)
    If Data.MvType(mv) = MT_SUPER_GRAB Then FX.Spawn Int(F(o).x + 0.5), Int(F(o).footY) - 20, FX_HEAVY
    TryGrab = True
End Function

' Super grab: lunge in, four punches, lift overhead, slam into the ground.
Private Const SG_END As Long = 46
Private Sub SuperGrabStep(ByVal i As Long)
    Dim o As Long, t As Long, tx As Double, hx As Long, hy As Long, slamDmg As Long
    o = 1 - i
    With F(i)
        t = .stTick
        If t <= 4 Then                                         ' close the distance
            tx = F(o).x - .facing * (.halfW + F(o).halfW - 2)
            .x = .x + (tx - .x) * 0.5
        End If
        hx = Int(F(o).x + 0.5) - .facing * 2
        hy = Int(F(o).footY) - 22
        Select Case t
        Case 6, 10, 14, 18                                     ' rapid punches
            Combat.DealDamage o, 40
            Combat.AddComboHit i
            F(o).flash = 1
            If t = 18 Then
                FX.Spawn hx, hy, FX_HEAVY
            Else
                FX.Spawn hx, hy, FX_HIT
            End If
            Combat.HitStop = 2
        Case 34                                                 ' slam it down in front of him
            F(o).footY = GROUND_Y
            F(o).x = .x + .facing * (.halfW + F(o).halfW + 6)
            slamDmg = Data.MvDamage(.mv) - 160
            If slamDmg < 0 Then slamDmg = 0
            Combat.DealDamage o, slamDmg
            Combat.AddComboHit i
            ApplyHit o, 0, 1.2, True
            F(o).vy = -2
            F(o).flash = 2
            FX.Spawn Int(F(o).x + 0.5), GROUND_Y - 3, FX_SLAM
            FX.Spawn Int(F(o).x + 0.5) - 8, GROUND_Y - 6, FX_HEAVY
            FX.Spawn Int(F(o).x + 0.5) + 8, GROUND_Y - 6, FX_HEAVY
            Combat.HitStop = 8
            Combat.ShowText .superName, 40, i
        End Select
        If t >= 22 And t < 34 And F(o).state = ST_THROWN Then   ' lifted overhead, then swung forward
            F(o).x = .x + .facing * IIf(t < 30, 3, 3 + (t - 29) * 4)
            F(o).footY = GROUND_Y - IIf(t < 28, (t - 21) * 5, IIf(t < 30, 30, 30 - (t - 29) * 6))
        End If
        If t >= SG_END Then SetState i, ST_IDLE
    End With
End Sub

' Vanish and reappear just behind the opponent (in front of them if there is no room).
Private Sub Teleport(ByVal i As Long)
    Dim o As Long, sd As Double, nx As Double, gap As Double
    o = 1 - i
    With F(i)
        If F(o).x >= .x Then sd = 1 Else sd = -1
        gap = .halfW + F(o).halfW + 3
        nx = F(o).x + sd * gap
        If nx < .halfW Or nx > SCREEN_W - .halfW Then nx = F(o).x - sd * gap
        .x = nx
        If F(o).x > .x Then .facing = 1 Else .facing = -1
    End With
End Sub

' True while a teleporting fighter is gone (not drawn, cannot be hit).
Public Function Vanished(ByVal i As Long) As Boolean
    Dim mt As Long
    If F(i).state <> ST_ATTACK Then Exit Function
    mt = Data.MvType(F(i).mv)
    If mt = MT_TELEPORT Or mt = MT_SUPER_TELE Then Vanished = (F(i).stTick >= 1 And F(i).stTick < Data.MvStartup(F(i).mv))
End Function

' Super rush connected: hold the opponent and run the slash sequence.
Public Sub StartSuperSlash(ByVal a As Long, ByVal d As Long)
    F(a).vx = 0
    SetState a, ST_THROW
    SetState d, ST_THROWN
    F(d).crouch = False
    F(d).bufAtk = -1
    F(d).footY = GROUND_Y
    F(d).onGround = True
End Sub

' Super slash: six quick cuts around the opponent, then a finishing slash that knocks down.
Private Const SS_END As Long = 40
Private Sub SuperSlashStep(ByVal i As Long)
    Dim o As Long, t As Long, cx As Long, cy As Long, fin As Long
    o = 1 - i
    With F(i)
        t = .stTick
        cx = Int(F(o).x + 0.5)
        cy = Int(F(o).footY) - 18
        Select Case t
        Case 2, 5, 8, 11, 14, 17
            Combat.DealDamage o, 30
            Combat.AddComboHit i
            F(o).flash = 1
            FX.Spawn cx + (t Mod 3) - 1, cy + ((t \ 3) Mod 3) * 3 - 3, .seqFx
            Combat.HitStop = 1
            .x = .x + .facing * IIf((t \ 3) Mod 2 = 0, 1, -1)     ' small footwork between cuts
        Case 26
            fin = Data.MvDamage(.mv) - 180
            If fin < 0 Then fin = 0
            Combat.DealDamage o, fin
            Combat.AddComboHit i
            ApplyHit o, 0, 3, True
            F(o).flash = 2
            FX.Spawn cx, cy, .seqFx
            FX.Spawn cx, cy - 4, FX_HEAVY
            FX.Spawn cx, cy + 4, FX_HEAVY
            Combat.HitStop = 8
            Combat.ShowText .superName, 40, i
        End Select
        If t >= SS_END Then SetState i, ST_IDLE
    End With
End Sub

Private Sub TechThrow(ByVal i As Long)
    Dim o As Long
    o = 1 - i
    SetState i, ST_GUARD: F(i).stun = 6: F(i).vx = -F(i).facing * 2.5
    SetState o, ST_GUARD: F(o).stun = 6: F(o).vx = -F(o).facing * 2.5
    F(i).bufAtk = -1
    FX.Spawn Int((F(0).x + F(1).x) / 2), Int(F(i).footY) - 22, FX_BLOCK
    Combat.ShowText "ESCAPE", 20
End Sub

Private Function IsLight(ByVal mv As Long) As Boolean
    Dim k As Long
    For k = 0 To 1
        If mv = F(k).moves(MV_LP) Or mv = F(k).moves(MV_LK) Or mv = F(k).moves(MV_CLP) Or mv = F(k).moves(MV_CLK) Then
            IsLight = True
            Exit Function
        End If
    Next k
End Function

Public Function MoveTotal(ByVal mv As Long) As Long
    MoveTotal = Data.MvStartup(mv) + Data.MvActive(mv) + Data.MvRecovery(mv)
End Function

' Is the opponent attacking (or its projectile coming) close enough that holding back should guard?
Public Function Threatened(ByVal i As Long) As Boolean
    Dim o As Long
    o = 1 - i
    If F(o).state = ST_ATTACK And Abs(F(o).x - F(i).x) < GUARD_RANGE Then Threatened = True
    If Projectile.IncomingDistance(i) < 40 Then Threatened = True
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
        If .state <> st Or st = ST_ATTACK Or st = ST_HIT Or st = ST_DOWN Or st = ST_THROW Or st = ST_DASH Then .stTick = 0
        .state = st
        If st = ST_CROUCH Then .crouch = True
        If st = ST_IDLE Or st = ST_WALK Or st = ST_JUMP Or st = ST_DASH Then .crouch = False
    End With
End Sub

' Called by combat and the debug keys.
Public Sub ApplyHit(ByVal i As Long, ByVal hitstun As Long, ByVal push As Double, ByVal knockdown As Boolean)
    With F(i)
        .bufAtk = -1
        .footY = IIf(.onGround, GROUND_Y, .footY)
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
    ClampX i            ' a knockdown pose is wider: keep it on screen already during hit stop
End Sub

Public Sub ApplyBlock(ByVal i As Long, ByVal blockstun As Long, ByVal push As Double)
    With F(i)
        SetState i, ST_GUARD
        .stun = blockstun
        .vx = -.facing * push
    End With
End Sub

Public Function AtWall(ByVal i As Long) As Boolean
    With F(i)
        AtWall = (.x <= .halfW + 0.5) Or (.x >= SCREEN_W - .halfW - 0.5)
    End With
End Function

' Keep the push box inside the stage. While knocked down the body is drawn lying
' sideways, so the push box becomes the sprite's real width (like the wider
' knockdown push boxes in fighting games) and the whole body stays on screen.
Private Sub ClampX(ByVal i As Long)
    Dim id As Long, l As Long, r As Long
    With F(i)
        If .x < .halfW Then .x = .halfW
        If .x > SCREEN_W - .halfW Then .x = SCREEN_W - .halfW
        If .state = ST_DOWN Then
            id = CurrentSprite(i)
            If .facing < 0 Then
                l = Int(.x + 0.5) + Data.SprAX(id) - (Data.SprW(id) - 1)
                r = Int(.x + 0.5) + Data.SprAX(id)
            Else
                l = Int(.x + 0.5) - Data.SprAX(id)
                r = l + Data.SprW(id) - 1
            End If
            If l < 0 Then .x = .x - l
            If r > SCREEN_W - 1 Then .x = .x - (r - (SCREEN_W - 1))
        End If
    End With
End Sub

' Keep the two bodies from overlapping (only when vertically close; not while lying down).
Private Sub ResolvePush()
    Dim dx As Double, minD As Double, ov As Double, s As Double
    If Abs(F(0).footY - F(1).footY) > 14 Then Exit Sub
    If F(0).state = ST_DOWN Or F(1).state = ST_DOWN Then Exit Sub
    If F(0).state = ST_THROWN Or F(1).state = ST_THROWN Then Exit Sub
    If Vanished(0) Or Vanished(1) Then Exit Sub
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
' boxes (used by Combat / CpuAI)
' ======================================================================

' True while the current attack is in its active frames.
Public Function InActive(ByVal i As Long) As Boolean
    Dim t As Long, mv As Long
    If F(i).state <> ST_ATTACK Then Exit Function
    t = F(i).stTick: mv = F(i).mv
    InActive = (t >= Data.MvStartup(mv)) And (t < Data.MvStartup(mv) + Data.MvActive(mv))
End Function

Public Function InRecovery(ByVal i As Long) As Boolean
    If F(i).state = ST_LAND Then
        InRecovery = True
    ElseIf F(i).state = ST_ATTACK And F(i).onGround Then
        InRecovery = (F(i).stTick >= Data.MvStartup(F(i).mv) + Data.MvActive(F(i).mv))
    End If
End Function

Public Function HitRect(ByVal i As Long, x0 As Long, y0 As Long, w As Long, h As Long) As Boolean
    Dim mv As Long, mt As Long
    If Not InActive(i) Then Exit Function
    With F(i)
        mv = .mv
        mt = Data.MvType(mv)
        If mt = MT_PROJ Or mt = MT_SUPER_PROJ Or mt = MT_GRAB Or mt = MT_SUPER_GRAB Or mt = MT_THROW Then Exit Function
        If mt = MT_PROJ_BOOM Or mt = MT_SUPER_FIRE Or mt = MT_TELEPORT Then Exit Function
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
        If Vanished(i) Then Exit Function
        If .state = ST_ATTACK Then
            If Data.MvType(.mv) = MT_RISE And .stTick < Data.MvP2(.mv) Then Exit Function   ' invincible uppercut
        End If
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
        Case ST_CROUCH, ST_LAND
            CurrentSprite = .sCrouch
        Case ST_JUMP
            CurrentSprite = .sJump
        Case ST_DASH
            If .backward Then CurrentSprite = .sBDash Else CurrentSprite = .sDash
        Case ST_ATTACK, ST_THROW
            If .state = ST_THROW And Data.MvType(.mv) = MT_SUPER_GRAB Then
                If t < 6 Then
                    CurrentSprite = Data.MvSprS(.mv)
                ElseIf t < 21 Then
                    If (t \ 4) Mod 2 = 0 Then CurrentSprite = .sPunchB Else CurrentSprite = .sPunchA
                ElseIf t < 34 Then
                    CurrentSprite = .sLift
                Else
                    CurrentSprite = .sSlam
                End If
            ElseIf .state = ST_THROW And (Data.MvType(.mv) = MT_SUPER_RUSH Or Data.MvType(.mv) = MT_SUPER_TELE) Then
                If t < 20 Then
                    Select Case (t \ 3) Mod 3
                    Case 0: CurrentSprite = .sSlashA
                    Case 1: CurrentSprite = .sSlashB
                    Case Else: CurrentSprite = .sSlashC
                    End Select
                ElseIf t < 26 Then
                    CurrentSprite = Data.MvSprS(.mv)
                Else
                    CurrentSprite = Data.MvSprA(.mv)
                End If
            ElseIf .state = ST_THROW Then
                If t >= Data.MvStartup(.mv) Then CurrentSprite = Data.MvSprA(.mv) Else CurrentSprite = Data.MvSprS(.mv)
            ElseIf Data.MvType(.mv) = MT_RISE And Not .onGround Then
                CurrentSprite = Data.MvSprA(.mv)
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
            If F(1 - i).state = ST_THROW Then
                If Data.MvType(F(1 - i).mv) = MT_SUPER_GRAB Then
                    If F(1 - i).stTick >= 22 Then
                        CurrentSprite = .sDown0                 ' held overhead
                    ElseIf (F(1 - i).stTick \ 4) Mod 2 = 1 Then
                        CurrentSprite = .sHit1
                    End If
                ElseIf Data.MvType(F(1 - i).mv) = MT_SUPER_RUSH Or Data.MvType(F(1 - i).mv) = MT_SUPER_TELE Then
                    If (F(1 - i).stTick \ 3) Mod 2 = 1 Then CurrentSprite = .sHit1
                End If
            End If
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
    If Vanished(i) Then Exit Sub
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
