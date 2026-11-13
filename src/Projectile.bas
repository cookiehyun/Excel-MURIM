Attribute VB_Name = "Projectile"
Option Explicit

' Projectiles (Danwoo's wind palm and its super version).
' One slot per owner; a fighter cannot fire while its projectile is alive.
' Two projectiles that touch cancel hit-for-hit (the super one usually wins).

Private mOn(0 To 1) As Boolean
Private mX(0 To 1) As Double
Private mY(0 To 1) As Double
Private mVX(0 To 1) As Double
Private mMv(0 To 1) As Long
Private mHits(0 To 1) As Long
Private mCD(0 To 1) As Long          ' ticks until it can hit again (multi-hit)
Private mAge(0 To 1) As Long
Private mStall(0 To 1) As Long       ' multi-hit: hold position while hitting
Private mBack(0 To 1) As Boolean     ' boomerang on its way back
Private mType(0 To 1) As Long
Private mAnimSmall As Long
Private mAnimBig As Long
Private Const BOOM_OUT As Long = 14   ' ticks before a thrown fan turns back
Private Const FIRE_DELAY As Long = 4  ' fire pillar: ticks of warning before it hurts
Private mAnimOwn(0 To 1) As Long      ' per-character projectile art (fx.proj_<name>) or -1

Private Function OwnAnim(ByVal owner As Long) As Long
    OwnAnim = -1
    If Data.HasAnim("fx.proj_" & Fighter.F(owner).prefix) Then OwnAnim = Data.AnimId("fx.proj_" & Fighter.F(owner).prefix)
End Function

Public Sub Init()
    mOn(0) = False: mOn(1) = False
    mAnimSmall = Data.AnimId("fx.wind")
    mAnimBig = Data.AnimId("fx.bigwind")
    mAnimOwn(0) = OwnAnim(0)
    mAnimOwn(1) = OwnAnim(1)
End Sub

Public Function Alive(ByVal owner As Long) As Boolean
    Alive = mOn(owner)
End Function

Public Sub Spawn(ByVal owner As Long, ByVal mv As Long)
    mType(owner) = Data.MvType(mv)
    mBack(owner) = False
    If mType(owner) = MT_SUPER_FIRE Then
        ' fire pillar under the opponent, where they stand right now
        mX(owner) = Fighter.F(1 - owner).x
        mY(owner) = GROUND_Y - Data.MvHbH(mv) \ 2
        mVX(owner) = 0
    Else
        With Fighter.F(owner)
            mX(owner) = .x + .facing * 12
            mY(owner) = .footY - 21
            mVX(owner) = .facing * Data.MvP1(mv)
        End With
    End If
    mMv(owner) = mv
    mAnimOwn(owner) = OwnAnim(owner)          ' art of whoever throws it now (tag mode swaps fighters)
    mHits(owner) = Data.MvP2(mv)
    If mHits(owner) < 1 Then mHits(owner) = 1
    mCD(owner) = 0
    mStall(owner) = 0
    mAge(owner) = 0
    mOn(owner) = True
End Sub

' Distance (px) from projectile of owner to fighter target, or 999 if none / moving away.
Public Function IncomingDistance(ByVal target As Long) As Double
    Dim o As Long
    o = 1 - target
    IncomingDistance = 999
    If Not mOn(o) Then Exit Function
    If (Fighter.F(target).x - mX(o)) * mVX(o) <= 0 Then Exit Function
    IncomingDistance = Abs(Fighter.F(target).x - mX(o))
End Function

Public Sub Update()
    Dim k As Long
    For k = 0 To 1
        If mOn(k) Then
            If mStall(k) > 0 Then
                mStall(k) = mStall(k) - 1
            Else
                mX(k) = mX(k) + mVX(k)
            End If
            mAge(k) = mAge(k) + 1
            If mCD(k) > 0 Then mCD(k) = mCD(k) - 1
            If mType(k) = MT_PROJ_BOOM Then
                If Not mBack(k) And mAge(k) >= BOOM_OUT Then          ' turn around
                    mBack(k) = True
                    mVX(k) = -mVX(k) * 1.15
                    mCD(k) = 0
                ElseIf mBack(k) And Abs(mX(k) - Fighter.F(k).x) < 6 Then
                    mOn(k) = False                                   ' caught by the thrower
                End If
            ElseIf mType(k) = MT_SUPER_FIRE Then
                If mAge(k) > 40 Then mOn(k) = False
            End If
            If mX(k) < -20 Or mX(k) > SCREEN_W + 20 Then mOn(k) = False
        End If
    Next k
    ' clash
    If mOn(0) And mOn(1) Then
        If Abs(mX(0) - mX(1)) < (HalfW(0) + HalfW(1)) And Abs(mY(0) - mY(1)) < 10 Then
            FX.Spawn Int((mX(0) + mX(1)) / 2), Int(mY(0)), FX_HEAVY
            mHits(0) = mHits(0) - 1: mHits(1) = mHits(1) - 1
            If mHits(0) <= 0 Then mOn(0) = False
            If mHits(1) <= 0 Then mOn(1) = False
        End If
    End If
    For k = 0 To 1
        If mOn(k) Then CheckHit k
    Next k
End Sub

Private Function HalfW(ByVal k As Long) As Double
    HalfW = Data.MvHbW(mMv(k)) / 2
End Function

Private Sub CheckHit(ByVal k As Long)
    Dim d As Long, ux As Long, uy As Long, uw As Long, uh As Long
    Dim px0 As Long, py0 As Long, pw As Long, ph As Long, lastHit As Boolean
    If mCD(k) > 0 Then Exit Sub
    If mType(k) = MT_SUPER_FIRE And mAge(k) < FIRE_DELAY Then Exit Sub
    d = 1 - k
    If Not Fighter.HurtRect(d, ux, uy, uw, uh) Then Exit Sub
    pw = Data.MvHbW(mMv(k)): ph = Data.MvHbH(mMv(k))
    px0 = Int(mX(k)) - pw \ 2: py0 = Int(mY(k)) - ph \ 2
    If px0 + pw <= ux Or ux + uw <= px0 Or py0 + ph <= uy Or uy + uh <= py0 Then Exit Sub
    mHits(k) = mHits(k) - 1
    lastHit = (mHits(k) <= 0)
    Combat.ApplyAttack k, d, mMv(k), Int(mX(k)), Int(mY(k)), True, lastHit
    If lastHit Then
        mOn(k) = False
    ElseIf mType(k) = MT_PROJ_BOOM Then
        mCD(k) = 999                                  ' one hit per direction
    Else
        mCD(k) = 3
        If mType(k) <> MT_SUPER_FIRE Then mStall(k) = 3
    End If
End Sub

Public Sub Draw()
    Dim k As Long, a As Long
    For k = 0 To 1
        If mOn(k) Then
            If mType(k) = MT_SUPER_FIRE Then
                FirePillar k
                a = -1
            ElseIf Data.MvType(mMv(k)) = MT_SUPER_PROJ Then
                a = mAnimBig
            ElseIf mAnimOwn(k) >= 0 Then
                a = mAnimOwn(k)
            Else
                a = mAnimSmall
            End If
            If a >= 0 Then Renderer.DrawSprite Data.AnimFrame(a, mAge(k)), Int(mX(k)), Int(mY(k)), (mVX(k) < 0)
        End If
    Next k
End Sub

' Flickering fire column: embers first (warning), then a tall pillar of flame.
Private Sub FirePillar(ByVal k As Long)
    Dim x0 As Long, y As Long, h As Long, w As Long, j As Long, c As Long, wob As Long
    x0 = Int(mX(k))
    If mAge(k) < FIRE_DELAY Then
        For j = -3 To 3 Step 2
            Renderer.PutPixel x0 + j * 2, GROUND_Y - 1 - ((mAge(k) + j) Mod 2), RGB(255, 160, 60)
        Next j
        Exit Sub
    End If
    h = Data.MvHbH(mMv(k)) * IIf(mAge(k) < FIRE_DELAY + 3, (mAge(k) - FIRE_DELAY + 1) / 3, 1)
    If mAge(k) > 34 Then h = h * (40 - mAge(k)) / 6
    For y = 0 To h - 1
        wob = Int(2 * Sin((y + mAge(k) * 3) * 0.5))
        w = 9 - y \ 6
        If w < 2 Then w = 2
        If y < h \ 3 Then
            c = RGB(255, 240, 160)
        ElseIf y < (h * 2) \ 3 Then
            c = RGB(255, 160, 50)
        Else
            c = RGB(220, 60, 40)
        End If
        Renderer.FillRect x0 - w + wob, GROUND_Y - 1 - y, w * 2, 1, c
        If w > 3 Then Renderer.FillRect x0 - w \ 3 + wob, GROUND_Y - 1 - y, (w * 2) \ 3, 1, RGB(255, 250, 220)
    Next y
End Sub

Public Sub Clear()
    mOn(0) = False: mOn(1) = False
End Sub
