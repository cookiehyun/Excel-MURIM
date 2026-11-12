Attribute VB_Name = "Match"
Option Explicit

' Rounds, timer and best-of-3 for VS CPU; free play with refills for training.
'   VS CPU   : ROUND n -> FIGHT! -> (KO | TIME) -> round winner -> ... first to 2 wins
'   Training : no timer / rounds; health refills after a combo, gauge stays full
'   Harness  : neither (plain sandbox)

Public Const MP_INTRO As Long = 1
Public Const MP_FIGHT As Long = 2
Public Const MP_KO As Long = 3
Public Const MP_TIMEUP As Long = 4
Public Const MP_ROUNDEND As Long = 5
Public Const MP_MATCHEND As Long = 6

Public RoundTicks As Long               ' round length in ticks (60 s = 1200); harness may shorten it
Private Const WIN_ROUNDS As Long = 2
Private Const MAX_ROUNDS As Long = 5
Private Const INTRO_TICKS As Long = 50  ' ROUND n (30) + FIGHT! (20)
Private Const TIMEUP_TICKS As Long = 40
Private Const ROUNDEND_TICKS As Long = 60
Private Const REFILL_DELAY As Long = 30

Private mRounds As Boolean              ' rounds + timer active (VS CPU)
Private mTraining As Boolean
Private mPhase As Long
Private mT As Long
Private mRound As Long
Private mWins(0 To 1) As Long
Private mTime As Long
Private mRoundWinner As Long            ' 0 / 1, -1 = draw
Private mMatchWinner As Long
Private mPerfect As Boolean
Private mCalm(0 To 1) As Long           ' training: ticks since the fighter was last in trouble

Public Sub BeginMatch(ByVal rounds As Boolean, ByVal training As Boolean)
    If RoundTicks <= 0 Then RoundTicks = 1200
    mRounds = rounds
    mTraining = training
    mRound = 1
    mWins(0) = 0: mWins(1) = 0
    StartRound False
End Sub

Private Sub StartRound(ByVal keepMeter As Boolean)
    Fighter.ResetRound keepMeter
    Projectile.Clear
    Combat.Init
    FX.Init
    mTime = RoundTicks
    mT = 0
    mCalm(0) = 0: mCalm(1) = 0
    If mRounds Then mPhase = MP_INTRO Else mPhase = MP_FIGHT
    If mTraining Then
        Fighter.F(0).meter = METER_MAX
        Fighter.F(1).meter = METER_MAX
    End If
End Sub

Public Property Get Phase() As Long
    Phase = mPhase
End Property

Public Property Get Wins(ByVal i As Long) As Long
    Wins = mWins(i)
End Property

Public Property Get RoundsOn() As Boolean
    RoundsOn = mRounds
End Property

Public Function InputAllowed() As Boolean
    InputAllowed = (mPhase = MP_FIGHT)
End Function

' Called once per game tick, after the simulation step.
Public Sub Tick(ByVal frozen As Boolean)
    Dim i As Long
    mT = mT + 1
    Select Case mPhase
    Case MP_INTRO
        If mT >= INTRO_TICKS Then mPhase = MP_FIGHT: mT = 0

    Case MP_FIGHT
        If Combat.IsKO Then
            mPhase = MP_KO: mT = 0
        ElseIf mRounds And Not frozen Then
            mTime = mTime - 1
            If mTime <= 0 Then
                mTime = 0
                mPhase = MP_TIMEUP: mT = 0
            End If
        End If
        If mTraining Then TrainingRefill

    Case MP_KO
        If Combat.KoFinished Then
            If mRounds Then
                RoundOver
            ElseIf Game.AutoRestart Then
                Game.StartMatch
            Else
                StartRound True
            End If
        End If

    Case MP_TIMEUP
        If mT >= TIMEUP_TICKS Then RoundOver

    Case MP_ROUNDEND
        If mT >= ROUNDEND_TICKS Then
            If mWins(0) >= WIN_ROUNDS Or mWins(1) >= WIN_ROUNDS Or mRound >= MAX_ROUNDS Then
                If mWins(0) > mWins(1) Then
                    mMatchWinner = 0
                ElseIf mWins(1) > mWins(0) Then
                    mMatchWinner = 1
                Else
                    mMatchWinner = -1
                End If
                mPhase = MP_MATCHEND: mT = 0
            Else
                mRound = mRound + 1
                StartRound True
            End If
        End If

    Case MP_MATCHEND
        If Game.AutoRestart Then
            If mT >= 30 Then Game.StartMatch
        ElseIf KeyInput.IsPressed(BTN_LP) Then
            Game.StartMatch                              ' rematch, same settings
        ElseIf KeyInput.IsPressed(BTN_HP) Then
            Game.GameState = GS_MENU
        End If
    End Select
End Sub

' Decide the round: KO = whoever still stands; time = higher health ratio.
Private Sub RoundOver()
    Dim r0 As Double, r1 As Double
    r0 = Fighter.F(0).health / Fighter.F(0).maxHealth
    r1 = Fighter.F(1).health / Fighter.F(1).maxHealth
    If r0 > r1 Then
        mRoundWinner = 0
    ElseIf r1 > r0 Then
        mRoundWinner = 1
    Else
        mRoundWinner = -1
    End If
    mPerfect = False
    If mRoundWinner >= 0 Then
        mWins(mRoundWinner) = mWins(mRoundWinner) + 1
        mPerfect = (Fighter.F(mRoundWinner).health = Fighter.F(mRoundWinner).maxHealth)
    End If
    mPhase = MP_ROUNDEND: mT = 0
End Sub

' Training: refill health once a fighter has been out of trouble for a moment; keep gauges full.
Private Sub TrainingRefill()
    Dim i As Long, st As Long
    For i = 0 To 1
        With Fighter.F(i)
            st = .state
            If st = ST_HIT Or st = ST_DOWN Or st = ST_THROWN Or st = ST_GUARD Then
                mCalm(i) = 0
            Else
                mCalm(i) = mCalm(i) + 1
                If mCalm(i) = REFILL_DELAY Then .health = .maxHealth
            End If
            If .meter < METER_MAX And st <> ST_ATTACK Then .meter = METER_MAX
        End With
    Next i
End Sub

' ---------------------------------------------------------------- drawing

Public Sub Draw()
    Dim cx As Long, secs As Long, c As Long, k As Long, t As String
    cx = SCREEN_W \ 2
    If mRounds Then
        ' timer between the health bars
        secs = (mTime + 19) \ 20
        If secs <= 10 And (mT \ 5) Mod 2 = 0 And mPhase = MP_FIGHT Then c = RGB(255, 90, 70) Else c = RGB(255, 236, 190)
        PixelFont.DrawBig Right$("0" & secs, 2), cx, 1, c, 1
        ' round win markers next to the timer
        For k = 0 To WIN_ROUNDS - 1
            If mWins(0) > k Then c = RGB(255, 200, 80) Else c = RGB(50, 42, 60)
            Renderer.FillRect cx - 16 - k * 5, 10, 3, 3, c
            If mWins(1) > k Then c = RGB(255, 200, 80) Else c = RGB(50, 42, 60)
            Renderer.FillRect cx + 14 + k * 5, 10, 3, 3, c
        Next k
    Else
        If mTraining Then PixelFont.DrawText "TRAINING", cx, 2, RGB(150, 220, 255), 1, 1
    End If

    Select Case mPhase
    Case MP_INTRO
        If mT < 30 Then
            If mWins(0) = WIN_ROUNDS - 1 And mWins(1) = WIN_ROUNDS - 1 Then t = "FINAL ROUND" Else t = "ROUND " & mRound
            Banner t, 30, RGB(255, 236, 190), 2
        Else
            Banner "FIGHT!", 28, RGB(255, 120, 60), 3
        End If
    Case MP_TIMEUP
        Banner "TIME", 28, RGB(255, 236, 190), 3
    Case MP_ROUNDEND
        If mRoundWinner < 0 Then
            Banner "DRAW", 30, RGB(255, 236, 190), 2
        Else
            Banner UCase$(Fighter.F(mRoundWinner).prefix) & " WINS", 30, RGB(255, 236, 190), 2
            If mPerfect Then PixelFont.DrawBig "PERFECT", cx, 47, RGB(255, 200, 80), 1
        End If
    Case MP_MATCHEND
        If mMatchWinner < 0 Then
            t = "DRAW GAME"
        ElseIf Game.WatchMode Then
            t = UCase$(Fighter.F(mMatchWinner).prefix) & " WINS"
        ElseIf mMatchWinner = Fighter.Controlled Then
            t = "YOU WIN"
        Else
            t = "YOU LOSE"
        End If
        Banner t, 28, RGB(255, 204, 92), 3
        If Not Game.AutoRestart Then PixelFont.DrawText "A = REMATCH   S = MENU", cx, 54, RGB(220, 210, 230), 1, 1
    End Select
End Sub

' Big centred text on a dark strip so it reads on any stage.
Private Sub Banner(ByVal t As String, ByVal y As Long, ByVal c As Long, ByVal px As Long)
    Renderer.FillRect 0, y - 3, SCREEN_W, 7 * px + 7, RGB(14, 10, 20)
    Renderer.FillRect 0, y - 4, SCREEN_W, 1, RGB(120, 90, 40)
    Renderer.FillRect 0, y + 7 * px + 4, SCREEN_W, 1, RGB(120, 90, 40)
    PixelFont.DrawBig t, SCREEN_W \ 2, y, c, px
End Sub
