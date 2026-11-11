Attribute VB_Name = "Fighter"
Option Explicit

' Step 1 placeholder: a single box that walks, jumps and crouches.
' Units: pixels and logic ticks. Replaced by the real state machine in step 3.

Private Const BOX_W As Long = 8
Private Const BOX_H As Long = 16
Private Const CROUCH_H As Long = 10
Private Const WALK_SPEED As Double = 2      ' px / tick
Private Const JUMP_VY As Double = -4.5      ' px / tick
Private Const GRAVITY As Double = 0.5       ' px / tick^2

Private mX As Double        ' left edge
Private mFootY As Double    ' bottom edge (exclusive): standing => GROUND_Y
Private mVY As Double
Private mOnGround As Boolean
Private mCrouch As Boolean

Public Sub Init()
    mX = 20
    mFootY = GROUND_Y
    mVY = 0
    mOnGround = True
    mCrouch = False
End Sub

Public Sub Update()
    Dim vx As Double
    If KeyInput.IsDown(BTN_LEFT) Then vx = vx - WALK_SPEED
    If KeyInput.IsDown(BTN_RIGHT) Then vx = vx + WALK_SPEED

    mCrouch = mOnGround And KeyInput.IsDown(BTN_DOWN)
    If mCrouch Then vx = 0

    If mOnGround And KeyInput.IsPressed(BTN_UP) Then
        mVY = JUMP_VY
        mOnGround = False
    End If

    If Not mOnGround Then
        mVY = mVY + GRAVITY
        mFootY = mFootY + mVY
        If mFootY >= GROUND_Y Then
            mFootY = GROUND_Y
            mVY = 0
            mOnGround = True
        End If
    End If

    mX = mX + vx
    If mX < 0 Then mX = 0
    If mX > SCREEN_W - BOX_W Then mX = SCREEN_W - BOX_W
End Sub

Public Sub Draw()
    Dim h As Long, x As Long, top As Long, body As Long
    If mCrouch Then h = CROUCH_H Else h = BOX_H
    x = Int(mX + 0.5)
    top = Int(mFootY + 0.5) - h

    ' flash body color while any attack button is held (input check)
    body = RGB(220, 80, 60)
    If KeyInput.IsDown(BTN_LP) Or KeyInput.IsDown(BTN_HP) Or KeyInput.IsDown(BTN_LK) Or KeyInput.IsDown(BTN_HK) Then
        body = RGB(255, 220, 80)
    End If

    Renderer.FillRect x, top, BOX_W, h, body
    Renderer.FillRect x + 2, top + 2, BOX_W - 4, 3, RGB(250, 210, 180)   ' "face"
End Sub
