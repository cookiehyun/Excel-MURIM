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
Private mAnimSmall As Long
Private mAnimBig As Long

Public Sub Init()
    mOn(0) = False: mOn(1) = False
    mAnimSmall = Data.AnimId("fx.wind")
    mAnimBig = Data.AnimId("fx.bigwind")
End Sub

Public Function Alive(ByVal owner As Long) As Boolean
    Alive = mOn(owner)
End Function

Public Sub Spawn(ByVal owner As Long, ByVal mv As Long)
    With Fighter.F(owner)
        mX(owner) = .x + .facing * 12
        mY(owner) = .footY - 21
        mVX(owner) = .facing * Data.MvP1(mv)
    End With
    mMv(owner) = mv
    mHits(owner) = Data.MvP2(mv)
    If mHits(owner) < 1 Then mHits(owner) = 1
    mCD(owner) = 0
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
            mX(k) = mX(k) + mVX(k)
            mAge(k) = mAge(k) + 1
            If mCD(k) > 0 Then mCD(k) = mCD(k) - 1
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
    d = 1 - k
    If Not Fighter.HurtRect(d, ux, uy, uw, uh) Then Exit Sub
    pw = Data.MvHbW(mMv(k)): ph = Data.MvHbH(mMv(k))
    px0 = Int(mX(k)) - pw \ 2: py0 = Int(mY(k)) - ph \ 2
    If px0 + pw <= ux Or ux + uw <= px0 Or py0 + ph <= uy Or uy + uh <= py0 Then Exit Sub
    mHits(k) = mHits(k) - 1
    lastHit = (mHits(k) <= 0)
    Combat.ApplyAttack k, d, mMv(k), Int(mX(k)), Int(mY(k)), True, lastHit
    If lastHit Then mOn(k) = False Else mCD(k) = 3
End Sub

Public Sub Draw()
    Dim k As Long, a As Long
    For k = 0 To 1
        If mOn(k) Then
            If Data.MvType(mMv(k)) = MT_SUPER_PROJ Then a = mAnimBig Else a = mAnimSmall
            Renderer.DrawSprite Data.AnimFrame(a, mAge(k)), Int(mX(k)), Int(mY(k)), (mVX(k) < 0)
        End If
    Next k
End Sub

Public Sub Clear()
    mOn(0) = False: mOn(1) = False
End Sub
