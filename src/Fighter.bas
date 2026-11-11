Attribute VB_Name = "Fighter"
Option Explicit

' Step 2: two fighters with sprite animation (idle / walk / crouch / jump),
' auto-facing and body push. Tab switches which fighter the keyboard controls;
' the other one stands still. Attacks come with the state machine (step 3).

Private Const GRAVITY As Double = 0.5       ' px / tick^2

Private Type TFighter
    prefix As String
    x As Double             ' anchor x (center under hip)
    footY As Double         ' GROUND_Y when standing
    vy As Double
    onGround As Boolean
    crouch As Boolean
    moving As Boolean
    facing As Long          ' +1 right, -1 left
    walkSpeed As Double
    jumpVY As Double
    halfW As Double         ' push-box half width
    anim As Long
    animTick As Long
    aIdle As Long
    aWalk As Long
    aCrouch As Long
    aJump As Long
End Type

Private F(0 To 1) As TFighter
Private mControl As Long

Public Sub Init()
    Setup 0, "danwoo", 40, 2#, -4.5, 6
    Setup 1, "cheolsan", 80, 1.4, -4#, 8
    mControl = 0
End Sub

Public Property Get ControlledName() As String
    ControlledName = F(mControl).prefix
End Property

Private Sub Setup(ByVal i As Long, ByVal prefix As String, ByVal x As Double, _
                  ByVal walkSpeed As Double, ByVal jumpVY As Double, ByVal halfW As Double)
    With F(i)
        .prefix = prefix
        .x = x
        .footY = GROUND_Y
        .vy = 0
        .onGround = True
        .crouch = False
        .moving = False
        .facing = 1
        .walkSpeed = walkSpeed
        .jumpVY = jumpVY
        .halfW = halfW
        .aIdle = Data.AnimId(prefix & ".idle")
        .aWalk = Data.AnimId(prefix & ".walk")
        .aCrouch = Data.AnimId(prefix & ".crouch")
        .aJump = Data.AnimId(prefix & ".jump")
        .anim = .aIdle
        .animTick = 0
    End With
End Sub

Public Sub Update()
    Dim i As Long
    If KeyInput.IsPressed(BTN_SWAP) Then mControl = 1 - mControl
    For i = 0 To 1
        MoveFighter i, (i = mControl)
    Next i
    ResolvePush
    For i = 0 To 1
        UpdateFacing i
        UpdateAnim i
    Next i
End Sub

Private Sub MoveFighter(ByVal i As Long, ByVal human As Boolean)
    Dim vx As Double, l As Boolean, r As Boolean, u As Boolean, d As Boolean
    If human Then
        l = KeyInput.IsDown(BTN_LEFT): r = KeyInput.IsDown(BTN_RIGHT)
        u = KeyInput.IsPressed(BTN_UP): d = KeyInput.IsDown(BTN_DOWN)
    End If
    With F(i)
        If .onGround Then
            If l Then vx = vx - .walkSpeed
            If r Then vx = vx + .walkSpeed
            .crouch = d
            If .crouch Then vx = 0
            If u And Not .crouch Then
                .vy = .jumpVY
                .onGround = False
            End If
        Else
            .crouch = False
            If l Then vx = vx - .walkSpeed
            If r Then vx = vx + .walkSpeed
        End If

        If Not .onGround Then
            .vy = .vy + GRAVITY
            .footY = .footY + .vy
            If .footY >= GROUND_Y Then
                .footY = GROUND_Y
                .vy = 0
                .onGround = True
            End If
        End If

        .moving = (vx <> 0)
        .x = .x + vx
    End With
    ClampX i
End Sub

Private Sub ClampX(ByVal i As Long)
    With F(i)
        If .x < .halfW Then .x = .halfW
        If .x > SCREEN_W - .halfW Then .x = SCREEN_W - .halfW
    End With
End Sub

' Keep the two bodies from overlapping (only when vertically close).
Private Sub ResolvePush()
    Dim dx As Double, minD As Double, ov As Double, s As Double
    If Abs(F(0).footY - F(1).footY) > 14 Then Exit Sub
    dx = F(1).x - F(0).x
    minD = F(0).halfW + F(1).halfW
    If Abs(dx) >= minD Then Exit Sub
    If dx >= 0 Then s = 1 Else s = -1
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
        If .onGround Then
            If F(o).x > .x Then
                .facing = 1
            ElseIf F(o).x < .x Then
                .facing = -1
            End If
        End If
    End With
End Sub

Private Sub UpdateAnim(ByVal i As Long)
    Dim target As Long
    With F(i)
        If Not .onGround Then
            target = .aJump
        ElseIf .crouch Then
            target = .aCrouch
        ElseIf .moving Then
            target = .aWalk
        Else
            target = .aIdle
        End If
        If target <> .anim Then
            .anim = target
            .animTick = 0
        Else
            .animTick = .animTick + 1
        End If
    End With
End Sub

Public Sub Draw()
    ' draw the non-controlled fighter first so the player's one is on top
    DrawOne 1 - mControl
    DrawOne mControl
End Sub

Private Sub DrawOne(ByVal i As Long)
    With F(i)
        Renderer.DrawSprite Data.AnimFrame(.anim, .animTick), _
            Int(.x + 0.5), Int(.footY + 0.5), (.facing < 0)
    End With
End Sub
