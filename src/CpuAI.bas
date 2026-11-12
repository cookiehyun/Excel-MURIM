Attribute VB_Name = "CpuAI"
Option Explicit

' CPU opponent. Produces the same inputs a player would (direction + buttons)
' and feeds them through Fighter.FeedInput, so specials need real motions.
' Difficulty = reaction delay (it sees the opponent N ticks late) + how often it
' blocks, anti-airs, punishes, combos, escapes throws + mistake rate.

Public Const LV_EASY As Long = 0
Public Const LV_NORMAL As Long = 1
Public Const LV_HARD As Long = 2

Private Const SNAPS As Long = 16
Private Const QMAX As Long = 8

' difficulty parameters
Private pReact As Long, pMistake As Double, pBlock As Double, pAntiAir As Double
Private pCombo As Double, pTech As Double, pAggr As Double, pSpecial As Double, pPunish As Double
Private mLevel As Long

' per-CPU state (index = fighter controlled by the CPU)
Private sAtk(0 To 1, 0 To SNAPS - 1) As Boolean     ' opponent threatening (attack / projectile)
Private sLow(0 To 1, 0 To SNAPS - 1) As Boolean
Private sOver(0 To 1, 0 To SNAPS - 1) As Boolean
Private sAir(0 To 1, 0 To SNAPS - 1) As Boolean     ' opponent jumping in
Private sRecov(0 To 1, 0 To SNAPS - 1) As Boolean   ' opponent in recovery (punishable)
Private sPos(0 To 1) As Long

Private qDir(0 To 1, 0 To QMAX - 1) As Long
Private qPress(0 To 1, 0 To QMAX - 1) As Long
Private qHeld(0 To 1, 0 To QMAX - 1) As Long
Private qLen(0 To 1) As Long
Private qPos(0 To 1) As Long

Private mPlanDir(0 To 1) As Long
Private mPlanTicks(0 To 1) As Long
Private mGuardDecided(0 To 1) As Boolean
Private mGuardOn(0 To 1) As Boolean
Private mGuardLow(0 To 1) As Boolean
Private mAADecided(0 To 1) As Boolean
Private mPunishDecided(0 To 1) As Boolean
Private mComboDecided(0 To 1) As Boolean
Private mTechDecided(0 To 1) As Boolean
Private mAirDecided(0 To 1) As Boolean
Private mIdleLock(0 To 1) As Long

Public Sub Init(ByVal level As Long, Optional ByVal seed As Long = 0)
    Dim i As Long, k As Long
    mLevel = level
    Select Case level
    Case LV_EASY
        pReact = 7: pMistake = 0.35: pBlock = 0.25: pAntiAir = 0.1: pCombo = 0.15
        pTech = 0: pAggr = 0.35: pSpecial = 0.25: pPunish = 0.2
    Case LV_HARD
        pReact = 2: pMistake = 0.04: pBlock = 0.9: pAntiAir = 0.85: pCombo = 0.9
        pTech = 0.7: pAggr = 0.75: pSpecial = 0.8: pPunish = 0.9
    Case Else
        pReact = 4: pMistake = 0.15: pBlock = 0.6: pAntiAir = 0.45: pCombo = 0.55
        pTech = 0.3: pAggr = 0.55: pSpecial = 0.5: pPunish = 0.55
    End Select
    If seed <> 0 Then
        Rnd -1
        Randomize seed
    Else
        Randomize
    End If
    For i = 0 To 1
        qLen(i) = 0: qPos(i) = 0: mPlanTicks(i) = 0: mIdleLock(i) = 0
        For k = 0 To SNAPS - 1
            sAtk(i, k) = False: sLow(i, k) = False: sOver(i, k) = False: sAir(i, k) = False: sRecov(i, k) = False
        Next k
    Next i
End Sub

Public Function LevelName(ByVal level As Long) As String
    Select Case level
    Case LV_EASY: LevelName = "EASY"
    Case LV_HARD: LevelName = "HARD"
    Case Else: LevelName = "NORMAL"
    End Select
End Function

' ---------------------------------------------------------------- perception

' Record what the opponent of CPU fighter i is doing this tick.
Public Sub Observe(ByVal i As Long)
    Dim o As Long, p As Long, mv As Long
    o = 1 - i
    sPos(i) = (sPos(i) + 1) Mod SNAPS
    p = sPos(i)
    With Fighter.F(o)
        sAtk(i, p) = Fighter.Threatened(i)
        sLow(i, p) = False: sOver(i, p) = False
        If .state = ST_ATTACK Then
            mv = .mv
            sLow(i, p) = (Data.MvLevel(mv) = LV_LOW)
            sOver(i, p) = (Data.MvLevel(mv) = LV_OVERHEAD)
        End If
        sAir(i, p) = (Not .onGround) And (.state = ST_JUMP Or .state = ST_ATTACK) And _
                     ((Fighter.F(i).x - .x) * .vx >= 0)
        sRecov(i, p) = Fighter.InRecovery(o)
    End With
End Sub

' snapshot index "react" ticks ago
Private Function ViewIdx(ByVal i As Long) As Long
    ViewIdx = (sPos(i) - pReact + SNAPS * 4) Mod SNAPS
End Function

' ---------------------------------------------------------------- input queue

Private Sub QClear(ByVal i As Long)
    qLen(i) = 0: qPos(i) = 0
End Sub

Private Sub QAdd(ByVal i As Long, ByVal d As Long, Optional ByVal press As Long = 0, Optional ByVal held As Long = -1)
    If qLen(i) >= QMAX Then Exit Sub
    qDir(i, qLen(i)) = d
    qPress(i, qLen(i)) = press
    If held < 0 Then held = press
    qHeld(i, qLen(i)) = held
    qLen(i) = qLen(i) + 1
End Sub

' Queue a command move: motion string like "623" + button mask on the last step.
Private Sub QMotion(ByVal i As Long, ByVal motion As String, ByVal btn As Long)
    Dim k As Long
    QClear i
    For k = 1 To Len(motion)
        If k = Len(motion) Then
            QAdd i, CLng(Mid$(motion, k, 1)), btn
        Else
            QAdd i, CLng(Mid$(motion, k, 1))
        End If
    Next k
End Sub

' Queue the special in slot (MV_SP1 / MV_SP2 / MV_SUPER) using the motion from moves.txt.
Private Function QSpecial(ByVal i As Long, ByVal slot As Long, ByVal heavy As Boolean) As Boolean
    Dim cmd As String, btn As Long
    cmd = Data.MvCmd(Fighter.F(i).moves(slot))
    If Len(cmd) < 2 Then Exit Function
    If Right$(cmd, 1) = "K" Then
        If heavy Then btn = B_HK Else btn = B_LK
    Else
        If heavy Then btn = B_HP Else btn = B_LP
    End If
    QMotion i, Left$(cmd, Len(cmd) - 1), btn
    QSpecial = True
End Function

Private Function SpecialType(ByVal i As Long, ByVal slot As Long) As Long
    SpecialType = Data.MvType(Fighter.F(i).moves(slot))
End Function

' Find this fighter's special of a given type (or -1).
Private Function FindSpecial(ByVal i As Long, ByVal mt As Long) As Long
    Dim s As Long, st As Long
    FindSpecial = -1
    For s = MV_SP1 To MV_SP2
        st = SpecialType(i, s)
        If st = MT_PROJ_BOOM Then st = MT_PROJ                ' a returning fan is used like any projectile
        If st = mt Then
            FindSpecial = s
            Exit Function
        End If
    Next s
End Function

' ---------------------------------------------------------------- decision

Public Sub Think(ByVal i As Long)
    Dim o As Long, dist As Double, v As Long, d As Long, press As Long, held As Long, s As Long

    ' replay queued inputs first (motions, throws)
    If qPos(i) < qLen(i) Then
        d = qDir(i, qPos(i)): press = qPress(i, qPos(i)): held = qHeld(i, qPos(i))
        qPos(i) = qPos(i) + 1
        If qPos(i) >= qLen(i) Then QClear i
        Fighter.FeedInput i, d, press, held
        Exit Sub
    End If

    o = 1 - i
    v = ViewIdx(i)
    dist = Abs(Fighter.F(o).x - Fighter.F(i).x)
    d = 5

    With Fighter.F(i)
        If .state <> ST_THROWN Then mTechDecided(i) = False
        If .state <> ST_ATTACK Then mComboDecided(i) = False
        If .state <> ST_JUMP Then mAirDecided(i) = False

        Select Case .state
        Case ST_THROWN
            If Not mTechDecided(i) Then
                mTechDecided(i) = True
                If Rnd < pTech Then
                    Fighter.FeedInput i, 5, B_LP Or B_LK, B_LP Or B_LK
                    Exit Sub
                End If
            End If

        Case ST_ATTACK
            If .hasHit And Not mComboDecided(i) Then
                mComboDecided(i) = True
                If Rnd < pCombo Then
                    ' combo: chain a light, or cancel into a special / super
                    If .meter >= METER_MAX And Rnd < 0.7 Then
                        QSpecial i, MV_SUPER, True
                    ElseIf Data.MvType(.mv) = MT_NORMAL And Rnd < pSpecial Then
                        s = FindSpecial(i, MT_PROJ)
                        If s < 0 Then s = FindSpecial(i, MT_RUSH)
                        If s >= 0 Then
                            QSpecial i, s, True
                        Else
                            QAdd i, 5, B_HP
                        End If
                    ElseIf .chain < 2 Then
                        If .crouch Then
                            QAdd i, 2, B_LK
                        Else
                            QAdd i, 5, B_LP
                        End If
                    Else
                        QAdd i, 5, B_HP
                    End If
                End If
            End If
            If .crouch Then d = 2

        Case ST_JUMP
            ' jump-in attack when falling near the opponent
            If .vy > 0 And dist < 26 And Not mAirDecided(i) Then
                mAirDecided(i) = True
                If Rnd > pMistake Then
                    Fighter.FeedInput i, 5, B_HK, B_HK
                    Exit Sub
                End If
            End If

        Case ST_DOWN, ST_HIT, ST_LAND
            ' hold guard when getting back up next to the opponent
            If dist < 40 And Rnd < pBlock Then d = 4

        Case ST_IDLE, ST_WALK, ST_CROUCH, ST_GUARD
            d = Neutral(i, o, v, dist)
            If qLen(i) > 0 Then
                Think i                       ' a plan queued inputs: start them now
                Exit Sub
            End If
        End Select
    End With
    Fighter.FeedInput i, d, 0, 0
End Sub

' Decisions while free to act. Returns the direction to hold (or queues inputs).
Private Function Neutral(ByVal i As Long, ByVal o As Long, ByVal v As Long, ByVal dist As Double) As Long
    Dim s As Long, r As Double, meFull As Boolean
    Neutral = 5

    ' 1. defend against an attack / projectile it has "seen"
    If sAtk(i, v) Then
        If Not mGuardDecided(i) Then
            mGuardDecided(i) = True
            mGuardOn(i) = (Rnd < pBlock)
            If mLevel = LV_HARD Then
                mGuardLow(i) = sLow(i, v)
            Else
                mGuardLow(i) = sLow(i, v) Or (Not sOver(i, v) And Rnd < 0.3)
            End If
        End If
        If mGuardOn(i) Then
            If mGuardLow(i) Then Neutral = 1 Else Neutral = 4
            mPlanTicks(i) = 0
            Exit Function
        End If
    Else
        mGuardDecided(i) = False
    End If

    ' 2. anti-air
    If sAir(i, v) And dist < 42 Then
        If Not mAADecided(i) Then
            mAADecided(i) = True
            If Rnd < pAntiAir Then
                s = FindSpecial(i, MT_RISE)
                If s >= 0 Then
                    QSpecial i, s, False
                Else
                    QAdd i, 2: QAdd i, 2, B_HP           ' crouching uppercut
                End If
                Exit Function
            End If
        End If
    Else
        mAADecided(i) = False
    End If

    ' 3. punish a whiffed / blocked move
    If sRecov(i, v) And dist < 34 Then
        If Not mPunishDecided(i) Then
            mPunishDecided(i) = True
            If Rnd < pPunish Then
                s = FindSpecial(i, MT_GRAB)
                If s >= 0 And dist < 26 Then
                    QSpecial i, s, False
                ElseIf dist < 22 Then
                    QAdd i, 5, B_HP
                Else
                    QAdd i, 2, B_HK                       ' sweep
                End If
                Exit Function
            End If
        End If
    Else
        mPunishDecided(i) = False
    End If

    ' 4. super when the gauge is full
    meFull = (Fighter.F(i).meter >= METER_MAX)
    If meFull And Rnd < 0.04 * (1 + pAggr) Then
        If SpecialType(i, MV_SUPER) = MT_SUPER_TELE Or SpecialType(i, MV_SUPER) = MT_SUPER_FIRE Then
            QSpecial i, MV_SUPER, False                        ' teleport strike works from anywhere
            Exit Function
        ElseIf SpecialType(i, MV_SUPER) = MT_SUPER_RUSH Then
            If dist < 52 Then
                QSpecial i, MV_SUPER, True
                Exit Function
            End If
        ElseIf SpecialType(i, MV_SUPER) = MT_SUPER_PROJ Then
            If dist > 30 Then
                QSpecial i, MV_SUPER, True
                Exit Function
            End If
        ElseIf dist < 38 Then
            QSpecial i, MV_SUPER, True
            Exit Function
        End If
    End If

    ' 5. keep following the current plan
    If mPlanTicks(i) > 0 Then
        mPlanTicks(i) = mPlanTicks(i) - 1
        Neutral = mPlanDir(i)
        Exit Function
    End If

    ' 6. new plan
    r = Rnd
    If Rnd < pMistake Then
        ' mistake: wander or stand around
        Plan i, Choose(1 + Int(Rnd * 4), 5, 4, 6, 2), 4 + Int(Rnd * 8)
    ElseIf Zoner(i) > 0 And dist > 30 And r < 0.4 * Zoner(i) And FindSpecial(i, MT_PROJ) >= 0 And Not Projectile.Alive(i) Then
        QSpecial i, FindSpecial(i, MT_PROJ), Rnd < 0.5      ' zoners: keep throwing from range
    ElseIf Zoner(i) > 0 And dist < 42 And r < 0.3 * Zoner(i) Then
        Plan i, 4, 4 + Int(Rnd * 5)                         ' zoners: back off to their range
    ElseIf dist > 55 Then
        s = FindSpecial(i, MT_PROJ)
        If s >= 0 And r < 0.3 * pSpecial + 0.1 And Not Projectile.Alive(i) Then
            QSpecial i, s, Rnd < 0.5
        ElseIf r < 0.5 Then
            Plan i, 6, 6 + Int(Rnd * 8)
        ElseIf r < 0.65 Then
            QAdd i, 6: QAdd i, 5: QAdd i, 6                ' dash in
        ElseIf r < 0.65 + 0.15 * pAggr Then
            QAdd i, 9                                     ' jump in
        Else
            Plan i, 5, 4 + Int(Rnd * 6)
        End If
    ElseIf dist > 26 Then
        s = FindSpecial(i, MT_RUSH)
        If s < 0 Then s = FindSpecial(i, MT_PROJ)
        If s >= 0 And r < 0.2 * pSpecial And Not Projectile.Alive(i) Then
            QSpecial i, s, True
        ElseIf r < 0.35 + 0.2 * pAggr Then
            Plan i, 6, 3 + Int(Rnd * 5)
        ElseIf r < 0.55 + 0.2 * pAggr Then
            If Rnd < 0.5 Then                             ' poke
                QAdd i, 5, B_LK
            Else
                QAdd i, 2, B_LK
            End If
        ElseIf r < 0.65 + 0.2 * pAggr Then
            QAdd i, 9
        ElseIf r < 0.85 Then
            Plan i, 4, 3 + Int(Rnd * 5)
        Else
            Plan i, 1, 4 + Int(Rnd * 6)                   ' crouch and wait
        End If
    Else
        s = FindSpecial(i, MT_TELEPORT)
        If s >= 0 And Rnd < 0.04 + 0.04 * pSpecial Then
            QSpecial i, s, False                              ' escape / cross-up
            Exit Function
        End If
        s = FindSpecial(i, MT_GRAB)
        If r < 0.3 * pAggr + 0.1 Then
            QAdd i, 5, B_LP                               ' jab (chains if it hits)
        ElseIf r < 0.45 * pAggr + 0.2 Then
            If s >= 0 And Rnd < pSpecial Then
                QSpecial i, s, False
            Else
                QAdd i, 5, B_LP Or B_LK                   ' throw
            End If
        ElseIf r < 0.55 + 0.1 * pAggr Then
            QAdd i, 2, B_LK
        ElseIf r < 0.75 Then
            Plan i, 4, 3 + Int(Rnd * 4)
        ElseIf r < 0.85 Then
            QAdd i, 4: QAdd i, 5: QAdd i, 4               ' back dash
        Else
            Plan i, 1, 3 + Int(Rnd * 4)
        End If
    End If
    Neutral = mPlanDir(i)
    If qLen(i) > 0 Then Neutral = 5
End Function

' How much this fighter's CPU prefers to keep distance and throw projectiles (0 = rush down).
Private Function Zoner(ByVal i As Long) As Double
    Select Case Fighter.F(i).prefix
    Case "heukrim": Zoner = 1#
    Case "hongryeon": Zoner = 0.7
    Case "danwoo": Zoner = 0.3
    Case Else: Zoner = 0
    End Select
End Function

Private Sub Plan(ByVal i As Long, ByVal d As Long, ByVal ticks As Long)
    mPlanDir(i) = d
    mPlanTicks(i) = ticks
End Sub
