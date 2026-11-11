Attribute VB_Name = "Fighter"
Option Explicit

' Step 3: fighter state machine.
'   IDLE / WALK / CROUCH / JUMP / ATTACK / GUARD / HIT / DOWN
' Attacks run on frame data from data\moves.txt (startup / active / recovery).
' Hit detection is step 4; until then debug keys trigger hit reactions:
'   Tab = switch controlled fighter, H = hit the other one, J = knock it down,
'   D = dummy mode (other fighter attacks on its own), F1 = show hitboxes.

Public Const ST_IDLE As Long = 0
Public Const ST_WALK As Long = 1
Public Const ST_CROUCH As Long = 2
Public Const ST_JUMP As Long = 3
Public Const ST_ATTACK As Long = 4
Public Const ST_GUARD As Long = 5
Public Const ST_HIT As Long = 6
Public Const ST_DOWN As Long = 7

' move slots per fighter
Private Const MV_LP As Long = 0
Private Const MV_HP As Long = 1
Private Const MV_LK As Long = 2
Private Const MV_HK As Long = 3
Private Const MV_CLP As Long = 4
Private Const MV_CHP As Long = 5
Private Const MV_CLK As Long = 6
Private Const MV_CHK As Long = 7
Private Const MV_JP As Long = 8
Private Const MV_JK As Long = 9

Private Const GRAVITY As Double = 0.5        ' px / tick^2
Private Const BACK_WALK As Double = 0.75     ' back-walk speed factor
Private Const GUARD_RANGE As Double = 60     ' hold back within this distance of an attack -> guard
Private Const DOWN_LIE_TICKS As Long = 16
Private Const GETUP_TICKS As Long = 4
Private Const DUMMY_INTERVAL As Long = 30

Private Type TFighter
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
    mv As Long               ' current move id (ATTACK)
    airAttacked As Boolean
    stun As Long             ' remaining hit/block stun
    stunTotal As Long
    downPhase As Long        ' 0 airborne, 1 lying, 2 getting up
    ' tuning
    walkSpeed As Double
    jumpVY As Double
    jumpVX As Double
    halfW As Double          ' push-box half width
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
    moves(0 To 9) As Long
End Type

Private F(0 To 1) As TFighter
Private mControl As Long
Private mDummy As Boolean
Private mDummyTick As Long
Private mDummyNext As Long
Private mShowBoxes As Boolean

Public Sub Init()
    Setup 0, "danwoo", 40, 2#, -4.5, 1.8, 6
    Setup 1, "cheolsan", 80, 1.4, -4#, 1.4, 8
    mControl = 0
    mDummy = False
    mShowBoxes = False
End Sub

Public Property Get StatusText() As String
    StatusText = "P1 " & F(mControl).prefix & " (Tab)" & _
                 IIf(mDummy, " | dummy ON (D)", " | dummy off (D)") & _
                 IIf(mShowBoxes, " | boxes ON (F1)", "")
End Property

Private Sub Setup(ByVal i As Long, ByVal prefix As String, ByVal x As Double, ByVal walkSpeed As Double, _
                  ByVal jumpVY As Double, ByVal jumpVX As Double, ByVal halfW As Double)
    Dim names As Variant, k As Long
    names = Array("lp", "hp", "lk", "hk", "clp", "chp", "clk", "chk", "jp", "jk")
    With F(i)
        .prefix = prefix
        .x = x
        .footY = GROUND_Y
        .vx = 0: .vy = 0
        .onGround = True
        .facing = 1
        .walkSpeed = walkSpeed
        .jumpVY = jumpVY
        .jumpVX = jumpVX
        .halfW = halfW
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
        For k = 0 To 9
            .moves(k) = Data.MoveId(prefix & "." & names(k))
        Next k
    End With
    SetState i, ST_IDLE
End Sub

' ======================================================================
' update
' ======================================================================

Public Sub Update()
    Dim i As Long
    If KeyInput.IsPressed(BTN_SWAP) Then mControl = 1 - mControl
    If KeyInput.IsPressed(BTN_DBG_DUMMY) Then mDummy = Not mDummy
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

' Called by combat (step 4) and the debug keys.
Public Sub ApplyHit(ByVal i As Long, ByVal hitstun As Long, ByVal push As Double, ByVal knockdown As Boolean)
    With F(i)
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

Private Sub UpdateFighter(ByVal i As Long)
    Dim l As Boolean, r As Boolean, u As Boolean, d As Boolean
    Dim fwd As Boolean, back As Boolean, atk As Long, wantGuard As Boolean

    ReadControls i, l, r, u, d, atk
    With F(i)
        If .facing > 0 Then
            fwd = r: back = l
        Else
            fwd = l: back = r
        End If
        wantGuard = back And Threatened(i)
        .stTick = .stTick + 1

        Select Case .state
        Case ST_IDLE, ST_WALK, ST_CROUCH, ST_GUARD
            If atk >= 0 Then
                StartGroundAttack i, atk, d
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

        Case ST_JUMP
            If atk >= 0 And Not .airAttacked Then
                .airAttacked = True
                If atk = MV_LP Or atk = MV_HP Then .mv = .moves(MV_JP) Else .mv = .moves(MV_JK)
                SetState i, ST_ATTACK
            End If
            If AirStep(i) Then SetState i, ST_IDLE

        Case ST_ATTACK
            If Not .onGround Then
                If AirStep(i) Then SetState i, ST_IDLE      ' landing ends an air attack
            ElseIf .stTick >= MoveTotal(.mv) Then
                If d Then SetState i, ST_CROUCH Else SetState i, ST_IDLE
            End If

        Case ST_HIT
            .x = .x + .vx
            .vx = .vx * 0.7
            .stun = .stun - 1
            If .stun <= 0 Then
                If .crouch Then SetState i, ST_CROUCH Else SetState i, ST_IDLE
            End If

        Case ST_DOWN
            Select Case .downPhase
            Case 0
                If AirStep(i) Then
                    .downPhase = 1
                    .stTick = 0
                End If
            Case 1
                If .stTick >= DOWN_LIE_TICKS Then
                    .downPhase = 2
                    .stTick = 0
                End If
            Case Else
                If .stTick >= GETUP_TICKS Then SetState i, ST_IDLE
            End Select
        End Select
    End With
    ClampX i
End Sub

Private Sub ReadControls(ByVal i As Long, l As Boolean, r As Boolean, u As Boolean, d As Boolean, atk As Long)
    atk = -1
    If i = mControl Then
        l = KeyInput.IsDown(BTN_LEFT): r = KeyInput.IsDown(BTN_RIGHT)
        u = KeyInput.IsDown(BTN_UP): d = KeyInput.IsDown(BTN_DOWN)
        ' heavier button wins if pressed on the same tick
        If KeyInput.IsPressed(BTN_LP) Then atk = MV_LP
        If KeyInput.IsPressed(BTN_LK) Then atk = MV_LK
        If KeyInput.IsPressed(BTN_HP) Then atk = MV_HP
        If KeyInput.IsPressed(BTN_HK) Then atk = MV_HK
    ElseIf mDummy Then
        ' training dummy: cycles through its ground normals
        mDummyTick = mDummyTick + 1
        If mDummyTick >= DUMMY_INTERVAL Then
            mDummyTick = 0
            atk = mDummyNext Mod 4
            d = (mDummyNext Mod 8) >= 4
            mDummyNext = mDummyNext + 1
        ElseIf F(i).state = ST_ATTACK Then
            d = F(i).crouch
        End If
    End If
End Sub

Private Sub StartGroundAttack(ByVal i As Long, ByVal atk As Long, ByVal crouching As Boolean)
    With F(i)
        If crouching Then .mv = .moves(atk + 4) Else .mv = .moves(atk)
        .crouch = crouching
        .vx = 0
        SetState i, ST_ATTACK
    End With
End Sub

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
        If .state <> st Or st = ST_ATTACK Or st = ST_HIT Or st = ST_DOWN Then .stTick = 0
        .state = st
        If st = ST_CROUCH Then .crouch = True
        If st = ST_IDLE Or st = ST_WALK Or st = ST_JUMP Then .crouch = False
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
        Case ST_ATTACK
            If InActive(i) Then CurrentSprite = Data.MvSprA(.mv) Else CurrentSprite = Data.MvSprS(.mv)
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
        Case ST_DOWN
            Select Case .downPhase
            Case 0: CurrentSprite = .sDown0
            Case 1: CurrentSprite = .sDown1
            Case Else: CurrentSprite = .sCrouch
            End Select
        End Select
    End With
End Function

' True while the current attack is in its active frames.
Public Function InActive(ByVal i As Long) As Boolean
    Dim t As Long, mv As Long
    If F(i).state <> ST_ATTACK Then Exit Function
    t = F(i).stTick: mv = F(i).mv
    InActive = (t >= Data.MvStartup(mv)) And (t < Data.MvStartup(mv) + Data.MvActive(mv))
End Function

Private Sub DrawOne(ByVal i As Long)
    With F(i)
        Renderer.DrawSprite CurrentSprite(i), Int(.x + 0.5), Int(.footY + 0.5), (.facing < 0)
    End With
End Sub

' Debug: push box (white) and, during active frames, the hitbox (red).
Private Sub DrawBoxes(ByVal i As Long)
    Dim mv As Long, x0 As Long, y0 As Long, w As Long, h As Long
    With F(i)
        Renderer.DrawRectOutline Int(.x - .halfW + 0.5), Int(.footY) - 2, Int(.halfW * 2), 2, RGB(255, 255, 255)
        If InActive(i) Then
            mv = .mv
            w = Data.MvHbW(mv): h = Data.MvHbH(mv)
            If .facing > 0 Then
                x0 = Int(.x + 0.5) + Data.MvHbX(mv)
            Else
                x0 = Int(.x + 0.5) - Data.MvHbX(mv) - w + 1
            End If
            y0 = Int(.footY + 0.5) + Data.MvHbY(mv)
            Renderer.DrawRectOutline x0, y0, w, h, RGB(255, 40, 40)
        End If
    End With
End Sub
