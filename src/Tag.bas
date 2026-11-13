Attribute VB_Name = "Tag"
Option Explicit

' Tag mode (2 vs 2, one fighter of each team on screen).
'   C           : tag out (standing / walking / crouching, short cooldown); the partner drops in from above
'   waiting     : the partner slowly gets back its recoverable (red) health and naegong
'   super gauge : shared by the team (handed over on every tag)
'   KO          : one knocked-out fighter loses the round (Tekken Tag rule)

Private Const TAG_COOLDOWN As Long = 40        ' ticks between tags
Private Const DROP_HEIGHT As Double = 46       ' incoming partner starts this high
Private Const RED_REGEN As Long = 2            ' hp per tick while waiting
Private Const CPU_TAG_LOW As Double = 0.5      ' CPU tags out below this health ratio...

Private mOn As Boolean
Private mCool(0 To 1) As Long
Private mTags(0 To 1) As Long                  ' tags this round (stats / tests)
' the waiting partner of side i lives in Fighter.F(2 + i); Fighter.F(4) is the swap buffer

Public Property Get Active() As Boolean
    Active = mOn
End Property

' Build both teams: partners first (parked in slots 2 / 3), then the point fighters in 0 / 1.
Public Sub Begin(ByVal p1a As String, ByVal p2a As String, ByVal p1b As String, ByVal p2b As String, ByVal humanIndex As Long)
    Fighter.Init p1b, p2b, humanIndex
    Fighter.CopyFighter Fighter.F(2), Fighter.F(0)
    Fighter.CopyFighter Fighter.F(3), Fighter.F(1)
    Fighter.Init p1a, p2a, humanIndex
    mOn = True
    mCool(0) = 0: mCool(1) = 0
End Sub

Public Sub Off()
    mOn = False
End Sub

Public Function ReservePrefix(ByVal side As Long) As String
    ReservePrefix = Fighter.F(2 + side).prefix
End Function

Public Function ReserveHealth(ByVal side As Long) As Long
    ReserveHealth = Fighter.F(2 + side).health
End Function

Public Function ReserveRed(ByVal side As Long) As Long
    ReserveRed = Fighter.F(2 + side).redCap
End Function

Public Function ReserveMax(ByVal side As Long) As Long
    ReserveMax = Fighter.F(2 + side).maxHealth
End Function

Public Function TagCount(ByVal side As Long) As Long
    TagCount = mTags(side)
End Function

' Health ratio of the whole team (time-up decision).
Public Function TeamRatio(ByVal side As Long) As Double
    TeamRatio = (Fighter.F(side).health + Fighter.F(2 + side).health) / (Fighter.F(side).maxHealth + Fighter.F(2 + side).maxHealth)
End Function

' New round: both team members back to full (point fighter is reset by Fighter.ResetRound).
Public Sub ResetRound()
    Dim i As Long
    If Not mOn Then Exit Sub
    For i = 0 To 1
        With Fighter.F(2 + i)
            .health = .maxHealth
            .healthShown = .maxHealth
            .redCap = .maxHealth
            .naegong = NAEGONG_MAX
        End With
        mCool(i) = 0
        mTags(i) = 0
    Next i
End Sub

' Once per simulation tick (not during hit stop / super freeze).
Public Sub Tick()
    Dim i As Long
    If Not mOn Then Exit Sub
    For i = 0 To 1
        If mCool(i) > 0 Then mCool(i) = mCool(i) - 1
        With Fighter.F(2 + i)                                   ' resting partner recovers
            If .health < .redCap Then
                .health = .health + RED_REGEN
                If .health > .redCap Then .health = .redCap
            End If
            .healthShown = .health
            If .naegong < NAEGONG_MAX Then .naegong = .naegong + 10
            If .naegong > NAEGONG_MAX Then .naegong = NAEGONG_MAX
        End With
    Next i
End Sub

' Human: C pressed. CPU: decides here as well.
Public Sub Inputs(ByVal allowed As Boolean)
    Dim i As Long
    If Not mOn Or Not allowed Then Exit Sub
    For i = 0 To 1
        If Game.IsCpu(i) Then
            If CpuWantsTag(i) Then TryTag i
        ElseIf i = Fighter.Controlled Then
            If KeyInput.IsPressed(BTN_TAG) Then TryTag i
        End If
    Next i
End Sub

Private Function CanTag(ByVal i As Long) As Boolean
    If mCool(i) > 0 Then Exit Function
    If Fighter.F(2 + i).health <= 0 Then Exit Function
    With Fighter.F(i)
        If Not .onGround Then Exit Function
        Select Case .state
        Case ST_IDLE, ST_WALK, ST_CROUCH
            CanTag = True
        End Select
    End With
End Function

Private Function CpuWantsTag(ByVal i As Long) As Boolean
    Dim r As Double, rr As Double
    If Not CanTag(i) Then Exit Function
    r = Fighter.F(i).health / Fighter.F(i).maxHealth
    rr = Fighter.F(2 + i).health / Fighter.F(2 + i).maxHealth
    If r < CPU_TAG_LOW And rr > r + 0.15 And Abs(Fighter.F(0).x - Fighter.F(1).x) > 30 Then
        CpuWantsTag = (Rnd < 0.15)
    End If
End Function

Public Sub TryTag(ByVal i As Long)
    Dim x As Double, fc As Long, meter As Long
    If Not CanTag(i) Then Exit Sub
    x = Fighter.F(i).x: fc = Fighter.F(i).facing: meter = Fighter.F(i).meter
    FX.Spawn Int(x + 0.5), Int(Fighter.F(i).footY) - 16, FX_DARK                  ' leaves in a puff
    Fighter.CopyFighter Fighter.F(4), Fighter.F(i)
    Fighter.CopyFighter Fighter.F(i), Fighter.F(2 + i)
    Fighter.CopyFighter Fighter.F(2 + i), Fighter.F(4)
    Fighter.F(2 + i).state = ST_IDLE
    Fighter.F(2 + i).stTick = 0
    Fighter.F(2 + i).bufAtk = -1
    With Fighter.F(i)
        .x = x
        .facing = fc
        .meter = meter                                                              ' shared team gauge
        .footY = GROUND_Y - DROP_HEIGHT
        .vx = 0: .vy = 0
        .onGround = False
        .bufAtk = -1: .bufDash = 0
        .flash = 0: .stun = 0
        .chain = 0
        .state = ST_JUMP
        .stTick = 0
        .airAttacked = False
    End With
    Fighter.ClearGhosts i
    mCool(i) = TAG_COOLDOWN
    mTags(i) = mTags(i) + 1
    Combat.ShowText "TAG", 16, i
End Sub
