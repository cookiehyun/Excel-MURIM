Attribute VB_Name = "Screens"
Option Explicit

' Front-end screens: title, character select (then level / stage select).

Private Const ROSTER As Long = 4
Private rPrefix(0 To ROSTER - 1) As String
Private rName(0 To ROSTER - 1) As String
Private rStyle(0 To ROSTER - 1) As String
Private rStat(0 To ROSTER - 1, 0 To 2) As Long     ' power, speed, range (1-5)
Private rHeadX(0 To ROSTER - 1) As Long           ' head centre in idle_0, relative to the anchor
Private rHeadY(0 To ROSTER - 1) As Long           ' (from docs/<name>_points.txt) - centres the portrait

Private mCursor As Long          ' title menu
Private mMenuTick As Long
Private mPhase As Long           ' select: 0 = player, 1 = opponent
Private mSel(0 To 1) As Long     ' cursor per phase
Private mPick(0 To 1) As Long    ' chosen roster slot per phase
Private mLevelSel As Long

Private cGold As Long, cDim As Long, cP1 As Long, cP2 As Long

Public Sub Init()
    rPrefix(0) = "danwoo": rName(0) = "DANWOO": rStyle(0) = "BALANCED"
    rStat(0, 0) = 3: rStat(0, 1) = 4: rStat(0, 2) = 4: rHeadX(0) = 2: rHeadY(0) = -30
    rPrefix(1) = "cheolsan": rName(1) = "CHEOLSAN": rStyle(1) = "GRAPPLER"
    rStat(1, 0) = 5: rStat(1, 1) = 2: rStat(1, 2) = 3: rHeadX(1) = 4: rHeadY(1) = -35
    rPrefix(2) = "": rName(2) = "???": rStyle(2) = "COMING SOON"
    rPrefix(3) = "": rName(3) = "???": rStyle(3) = "COMING SOON"
    cGold = RGB(240, 190, 80)
    cDim = RGB(130, 120, 150)
    cP1 = RGB(90, 170, 255)
    cP2 = RGB(255, 90, 80)
End Sub

Private Function Locked(ByVal k As Long) As Boolean
    Locked = (Len(rPrefix(k)) = 0)
End Function

' ---------------------------------------------------------------- title screen

Public Sub TitleTick()
    mMenuTick = mMenuTick + 1
    If KeyInput.IsPressed(BTN_UP) Or KeyInput.IsPressed(BTN_DOWN) Then mCursor = 1 - mCursor
    If KeyInput.IsPressed(BTN_LP) Then
        Game.GameMode = mCursor
        Game.WatchMode = False
        SelectBegin
    End If
End Sub

Public Sub TitleRender()
    Dim k As Long, y As Long, c As Long, lbl As String, cx As Long
    cx = SCREEN_W \ 2
    Renderer.BeginFrame True
    ' fighters facing each other on both sides
    Renderer.DrawSprite Data.AnimFrame(Fighter.F(0).aIdle, mMenuTick), 30, GROUND_Y, False
    Renderer.DrawSprite Data.AnimFrame(Fighter.F(1).aIdle, mMenuTick), SCREEN_W - 30, GROUND_Y, True
    ' title
    PixelFont.DrawBig "EXCEL", cx, 5, RGB(230, 220, 200), 2
    PixelFont.DrawBig "MUHYEOP", cx, 21, RGB(255, 204, 92), 3
    Renderer.FillRect cx - 50, 46, 100, 1, RGB(200, 150, 60)
    Renderer.FillRect cx - 38, 48, 76, 1, RGB(140, 100, 40)
    Renderer.FillRect cx - 52, 45, 3, 3, RGB(200, 150, 60)
    Renderer.FillRect cx + 49, 45, 3, 3, RGB(200, 150, 60)
    ' menu
    For k = 0 To 1
        y = 57 + k * 13
        If k = 0 Then lbl = "TRAINING" Else lbl = "VS CPU"
        If k = mCursor Then
            Renderer.FillRect cx - 30, y - 3, 60, 11, RGB(40, 28, 20)
            Renderer.DrawRectOutline cx - 30, y - 3, 60, 11, RGB(230, 180, 80)
            c = RGB(255, 236, 180)
            If (mMenuTick \ 6) Mod 2 = 0 Then
                PixelFont.DrawText ">", cx - 26, y, RGB(255, 200, 90)
                PixelFont.DrawText "<", cx + 24, y, RGB(255, 200, 90)
            End If
        Else
            Renderer.FillRect cx - 30, y - 3, 60, 11, RGB(20, 16, 26)
            c = RGB(130, 120, 150)
        End If
        PixelFont.DrawText lbl, cx, y, c, 1, 1
    Next k
    PixelFont.DrawText "UP DOWN = SELECT    A = OK    ESC = QUIT", cx, 92, RGB(150, 140, 170), 1, 1
End Sub

' ---------------------------------------------------------------- character select

Public Sub SelectBegin()
    mPhase = 0
    mSel(0) = 0: mSel(1) = 1
    Game.GameState = GS_SELECT
End Sub

Public Sub SelectTick()
    Dim k As Long
    mMenuTick = mMenuTick + 1
    k = mSel(mPhase)
    If KeyInput.IsPressed(BTN_LEFT) Or KeyInput.IsPressed(BTN_RIGHT) Then k = k Xor 1      ' 2 x 2 grid
    If KeyInput.IsPressed(BTN_UP) Or KeyInput.IsPressed(BTN_DOWN) Then k = k Xor 2
    mSel(mPhase) = k
    If KeyInput.IsPressed(BTN_LP) And Not Locked(k) Then
        mPick(mPhase) = k
        If mPhase = 0 Then
            mPhase = 1
        Else
            SelectDone
        End If
    ElseIf KeyInput.IsPressed(BTN_HP) Then                                                     ' S = back
        If mPhase = 1 Then mPhase = 0 Else Game.GameState = GS_MENU
    End If
End Sub

' VS CPU -> level select; training -> start (stage select comes next).
Private Sub SelectDone()
    Game.P1Char = rPrefix(mPick(0))
    Game.P2Char = rPrefix(mPick(1))
    Game.HumanIndex = 0
    If Game.GameMode = MODE_VSCPU Then
        mLevelSel = Game.CpuLevel
        Game.GameState = GS_LEVEL
    Else
        LaunchMatch
    End If
End Sub

Private Sub LaunchMatch()
    Game.StartMatch
    If Game.GameMode = MODE_TRAINING Then Fighter.SetDummy DUMMY_IDLE
End Sub

Private Function BoxX(ByVal k As Long) As Long
    BoxX = 72 + (k Mod 2) * 25
End Function

Private Function BoxY(ByVal k As Long) As Long
    BoxY = 47 + (k \ 2) * 24
End Function

Public Sub SelectRender()
    Dim k As Long, hover As Long, cx As Long, opp As String
    cx = SCREEN_W \ 2
    Renderer.BeginFrame True
    ' side panels for the two big views
    Renderer.FillRect 0, 10, 64, 90, RGB(14, 20, 40)
    Renderer.FillRect 128, 10, 64, 90, RGB(40, 14, 18)
    Renderer.FillRect 63, 10, 1, 90, cP1
    Renderer.FillRect 128, 10, 1, 90, cP2
    If mPhase = 0 Then
        PixelFont.DrawBig "SELECT FIGHTER", cx, 2, cGold, 1
    Else
        PixelFont.DrawBig "SELECT OPPONENT", cx, 2, cGold, 1
    End If
    If Game.GameMode = MODE_TRAINING Then opp = "DUMMY" Else opp = "CPU"
    PixelFont.DrawText "1P", 3, 13, cP1
    PixelFont.DrawText opp, SCREEN_W - 3, 13, cP2, 1, 2

    ' big views: player on the left, opponent on the right
    If mPhase = 0 Then
        BigView mSel(0), 32, False, 0, 63
        PixelFont.DrawBig "?", 160, 48, RGB(70, 40, 50), 3
    Else
        BigView mPick(0), 32, False, 0, 63
        BigView mSel(1), 160, True, 129, SCREEN_W - 1
    End If

    ' info for the slot under the cursor
    hover = mSel(mPhase)
    PixelFont.DrawBig rName(hover), cx, 13, RGB(255, 240, 210), 1
    PixelFont.DrawText rStyle(hover), cx, 23, cDim, 1, 1
    StatRow "POW", rStat(hover, 0), 28
    StatRow "SPD", rStat(hover, 1), 34
    StatRow "RNG", rStat(hover, 2), 40

    ' portrait grid
    For k = 0 To ROSTER - 1
        Portrait k
    Next k
    PixelFont.DrawText "A = OK  S = BACK", cx, 94, RGB(110, 100, 130), 1, 1
End Sub

Private Sub StatRow(ByVal lbl As String, ByVal v As Long, ByVal y As Long)
    Dim k As Long, c As Long
    PixelFont.DrawText lbl, 70, y, cDim
    For k = 0 To 4
        If k < v Then c = cGold Else c = RGB(50, 44, 60)
        Renderer.FillRect 84 + k * 7, y + 1, 6, 3, c
    Next k
End Sub

Private Sub BigView(ByVal k As Long, ByVal x As Long, ByVal flip As Boolean, ByVal clipL As Long, ByVal clipR As Long)
    Dim id As Long, ref As Long
    If Locked(k) Then
        PixelFont.DrawBig "?", x, 48, RGB(80, 70, 100), 3
        Exit Sub
    End If
    ' centre the drawn body (not the anchor) in the panel; use idle_0's width so breathing does not shift it
    ref = Data.SpriteId(rPrefix(k) & ".idle_0")
    If flip Then
        x = x - (Data.SprAX(ref) * 2 - Data.SprW(ref))
    Else
        x = x - (Data.SprW(ref) - Data.SprAX(ref) * 2)
    End If
    id = Data.AnimFrame(Data.AnimId(rPrefix(k) & ".idle"), mMenuTick)
    Renderer.DrawSpriteScaled id, x, 96, flip, 2, clipL, 10, clipR, 99
End Sub

Private Sub Portrait(ByVal k As Long)
    Dim x As Long, y As Long, id As Long, frameC As Long, blink As Boolean
    x = BoxX(k): y = BoxY(k)
    Renderer.FillRect x, y, 22, 22, RGB(26, 22, 36)
    If Locked(k) Then
        PixelFont.DrawBig "?", x + 11, y + 4, RGB(70, 64, 90), 2
    Else
        id = Data.SpriteId(rPrefix(k) & ".idle_0")
        ' head and shoulders, enlarged x2, head centred in the box
        Renderer.DrawSpriteScaled id, x + 11 - rHeadX(k) * 2, y + 10 - rHeadY(k) * 2, False, 2, x + 1, y + 1, x + 20, y + 20
    End If
    ' frames: chosen player slot (blue), cursor (gold for player, red for opponent)
    frameC = RGB(70, 60, 90)
    If mPhase = 1 And k = mPick(0) Then frameC = cP1
    blink = ((mMenuTick \ 5) Mod 2 = 0)
    If k = mSel(mPhase) And blink Then
        If mPhase = 0 Then frameC = cGold Else frameC = cP2
    End If
    Renderer.DrawRectOutline x, y, 22, 22, frameC
    If k = mSel(mPhase) Then Renderer.DrawRectOutline x - 1, y - 1, 24, 24, frameC
End Sub

' ---------------------------------------------------------------- level select (VS CPU)

Public Sub LevelTick()
    mMenuTick = mMenuTick + 1
    If KeyInput.IsPressed(BTN_UP) Then mLevelSel = (mLevelSel + 2) Mod 3
    If KeyInput.IsPressed(BTN_DOWN) Then mLevelSel = (mLevelSel + 1) Mod 3
    If KeyInput.IsPressed(BTN_LP) Then
        Game.CpuLevel = mLevelSel
        LaunchMatch
    ElseIf KeyInput.IsPressed(BTN_HP) Then
        mPhase = 1
        Game.GameState = GS_SELECT
    End If
End Sub

Public Sub LevelRender()
    Dim k As Long, y As Long, c As Long, cx As Long, desc As String, s As Long
    cx = SCREEN_W \ 2
    Renderer.BeginFrame True
    Renderer.FillRect 0, 10, 64, 90, RGB(14, 20, 40)
    Renderer.FillRect 128, 10, 64, 90, RGB(40, 14, 18)
    Renderer.FillRect 63, 10, 1, 90, cP1
    Renderer.FillRect 128, 10, 1, 90, cP2
    PixelFont.DrawBig "SELECT LEVEL", cx, 2, cGold, 1
    PixelFont.DrawText "1P", 3, 13, cP1
    PixelFont.DrawText "CPU", SCREEN_W - 3, 13, cP2, 1, 2
    BigView mPick(0), 32, False, 0, 63
    BigView mPick(1), 160, True, 129, SCREEN_W - 1
    PixelFont.DrawBig "VS", cx, 13, RGB(255, 240, 210), 2

    For k = 0 To 2
        y = 32 + k * 19
        Select Case k
        Case 0: desc = "SLOW REACTIONS"
        Case 1: desc = "SOLID DEFENSE"
        Case Else: desc = "MASTER LEVEL"
        End Select
        If k = mLevelSel Then
            Renderer.FillRect cx - 30, y - 2, 60, 17, RGB(40, 28, 20)
            Renderer.DrawRectOutline cx - 30, y - 2, 60, 17, cGold
            If (mMenuTick \ 6) Mod 2 = 0 Then Renderer.DrawRectOutline cx - 31, y - 3, 62, 19, cGold
            c = RGB(255, 236, 180)
        Else
            Renderer.FillRect cx - 30, y - 2, 60, 17, RGB(20, 16, 26)
            Renderer.DrawRectOutline cx - 30, y - 2, 60, 17, RGB(60, 52, 76)
            c = cDim
        End If
        PixelFont.DrawText CpuAI.LevelName(k), cx - 26, y + 1, c
        For s = 0 To k                                  ' difficulty stars
            Star cx + 22 - (2 - s) * 7, y + 3, IIf(k = mLevelSel, cGold, RGB(90, 80, 60))
        Next s
        PixelFont.DrawText desc, cx, y + 9, IIf(k = mLevelSel, RGB(200, 190, 170), RGB(90, 84, 104)), 1, 1
    Next k
    PixelFont.DrawText "A = START  S = BACK", cx, 94, RGB(110, 100, 130), 1, 1
End Sub

' 5x5 pixel star centred on (x, y)
Private Sub Star(ByVal x As Long, ByVal y As Long, ByVal c As Long)
    Renderer.FillRect x, y - 2, 1, 5, c
    Renderer.FillRect x - 2, y, 5, 1, c
    Renderer.FillRect x - 1, y - 1, 3, 3, c
End Sub
