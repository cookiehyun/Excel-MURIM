Attribute VB_Name = "Data"
Option Explicit

' Loads text data from <workbook folder>\data\ :
'   <fighter>.txt       PAL / SPRITE / ANIM records (see tools/gen_sprites.py)
'   stage_<name>.txt    PAL / STAGE records (see tools/gen_stage.py)
' Sprites are stored in one flat pixel pool; color -1 = transparent.
' Names are "<file>.<sprite>", e.g. "danwoo.idle_0".

Private Const DATA_DIR As String = "data"
Private Const ERR_BASE As Long = -2147221504   ' = vbObjectError (not defined in every VBA host)
Public BaseDir As String          ' optional override of the data folder (test harness)

Public SprCount As Long
Public SprW() As Long, SprH() As Long, SprAX() As Long, SprAY() As Long, SprOff() As Long
Public SprName() As String
Public SprPix() As Long
Private mPixUsed As Long

Public AnimCount As Long
Public AnimTPF() As Long, AnimStart() As Long, AnimLen() As Long
Public AnimFrames() As Long
Private mFrameNames() As String     ' resolved to sprite ids after all files are loaded
Private mFramesUsed As Long

Public StageW As Long, StageH As Long
Public StageCount As Long
Public StagePix() As Long          ' (stage, x, y)
Public StageName() As String       ' display names

' Moves (data\moves.txt)
Public Const MK_STAND As Long = 0
Public Const MK_CROUCH As Long = 1
Public Const MK_AIR As Long = 2
Public Const LV_HIGH As Long = 0
Public Const LV_LOW As Long = 1
Public Const LV_OVERHEAD As Long = 2

Public MvCount As Long
Public MvKind() As Long, MvStartup() As Long, MvActive() As Long, MvRecovery() As Long
Public MvDamage() As Long, MvHitstun() As Long, MvBlockstun() As Long, MvLevel() As Long, MvKD() As Boolean
Public MvHbX() As Long, MvHbY() As Long, MvHbW() As Long, MvHbH() As Long
Public MvSprS() As Long, MvSprA() As Long
Public MvType() As Long, MvP1() As Double, MvP2() As Double, MvCmd() As String
Public MvSeqPre() As Long, MvSeqAct() As Long, MvSeqRec() As Long   ' optional per-move anims (-1 = none)
Private mMvKey() As String

Public Const MT_NORMAL As Long = 0
Public Const MT_THROW As Long = 1
Public Const MT_PROJ As Long = 2
Public Const MT_RISE As Long = 3
Public Const MT_RUSH As Long = 4
Public Const MT_GRAB As Long = 5
Public Const MT_SUPER_PROJ As Long = 6
Public Const MT_SUPER_GRAB As Long = 7
Public Const MT_SUPER_RUSH As Long = 8      ' dash; on hit -> multi-slash sequence
Public Const MT_TELEPORT As Long = 9        ' vanish (invincible) and reappear behind the opponent
Public Const MT_SUPER_TELE As Long = 10     ' teleport, then a strike; on hit -> sequence
Public Const MT_PROJ_BOOM As Long = 11      ' projectile that flies out and returns (one hit each way)
Public Const MT_SUPER_FIRE As Long = 12     ' fire pillar erupting under the opponent (p2 = hits)

Private mSprIdx As Collection
Private mAnimIdx As Collection
Private mMoveIdx As Collection

Public Sub LoadAll()
    SprCount = 0: mPixUsed = 0: AnimCount = 0: mFramesUsed = 0
    StageW = 0: StageH = 0: StageCount = 0
    ReDim SprW(0 To 63): ReDim SprH(0 To 63): ReDim SprAX(0 To 63): ReDim SprAY(0 To 63)
    ReDim SprOff(0 To 63): ReDim SprName(0 To 63)
    ReDim SprPix(0 To 65535)
    ReDim AnimTPF(0 To 31): ReDim AnimStart(0 To 31): ReDim AnimLen(0 To 31)
    ReDim AnimFrames(0 To 255)
    ReDim mFrameNames(0 To 255)
    Set mSprIdx = New Collection
    Set mAnimIdx = New Collection

    LoadSpriteFile "danwoo"
    LoadSpriteFile "cheolsan"
    LoadSpriteFile "muyeong"
    LoadSpriteFile "seolhwa"
    LoadSpriteFile "heukrim"
    LoadSpriteFile "hongryeon"
    LoadSpriteFile "cheonma"                       ' arcade final boss (not on the select screen)
    LoadSpriteFile "fx"
    ' stages: file name, display name (order = stage select order)
    Dim st As Variant, k As Long
    st = Array("sanmun", "MOUNTAIN GATE", "bamboo", "BAMBOO FOREST", "inn", "NIGHT INN", _
               "snow", "SNOW PEAK", "waterfall", "WATERFALL")
    StageCount = (UBound(st) + 1) \ 2
    ReDim StageName(0 To StageCount - 1)
    For k = 0 To StageCount - 1
        StageName(k) = st(k * 2 + 1)
        LoadStageFile "stage_" & st(k * 2), k
    Next k
    LoadMoves
    ResolveAnimFrames
    ResolveMoveAnims
End Sub

' Optional multi-frame motion per move, from ANIM records in the character file:
'   <move>_pre (startup), <move>_seq (active and airborne part), <move>_rec (recovery).
Private Sub ResolveMoveAnims()
    Dim k As Long
    ReDim MvSeqPre(0 To MvCount): ReDim MvSeqAct(0 To MvCount): ReDim MvSeqRec(0 To MvCount)
    For k = 0 To MvCount - 1
        MvSeqPre(k) = OptAnim(mMvKey(k) & "_pre")
        MvSeqAct(k) = OptAnim(mMvKey(k) & "_seq")
        MvSeqRec(k) = OptAnim(mMvKey(k) & "_rec")
    Next k
End Sub

Private Function OptAnim(ByVal nm As String) As Long
    If HasAnim(nm) Then OptAnim = AnimId(nm) Else OptAnim = -1
End Function

Private Sub ResolveAnimFrames()
    Dim k As Long
    For k = 0 To mFramesUsed - 1
        AnimFrames(k) = SpriteId(mFrameNames(k))
    Next k
End Sub

Public Function HasAnim(ByVal nm As String) As Boolean
    Dim v As Long
    On Error GoTo Nope
    v = mAnimIdx.Item(nm)
    HasAnim = True
Nope:
End Function

Public Function MoveId(ByVal nm As String) As Long
    On Error GoTo Missing
    MoveId = mMoveIdx.Item(nm)
    Exit Function
Missing:
    Err.Raise ERR_BASE + 516, "Data", "Move not found: " & nm & " (data\moves.txt)"
End Function

Public Function SpriteId(ByVal nm As String) As Long
    On Error GoTo Missing
    SpriteId = mSprIdx.Item(nm)
    Exit Function
Missing:
    Err.Raise ERR_BASE + 513, "Data", "Sprite not found: " & nm
End Function

Public Function AnimId(ByVal nm As String) As Long
    On Error GoTo Missing
    AnimId = mAnimIdx.Item(nm)
    Exit Function
Missing:
    Err.Raise ERR_BASE + 514, "Data", "Anim not found: " & nm
End Function

' Sprite id at tick t, playing once and then holding the last frame.
Public Function AnimFrameHold(ByVal a As Long, ByVal t As Long) As Long
    Dim k As Long
    k = t \ AnimTPF(a)
    If k > AnimLen(a) - 1 Then k = AnimLen(a) - 1
    If k < 0 Then k = 0
    AnimFrameHold = AnimFrames(AnimStart(a) + k)
End Function

' Sprite id of the frame shown at tick t of an animation (loops). reverse = play backwards.
Public Function AnimFrame(ByVal a As Long, ByVal t As Long, Optional ByVal reverse As Boolean = False) As Long
    Dim k As Long
    k = (t \ AnimTPF(a)) Mod AnimLen(a)
    If reverse Then k = AnimLen(a) - 1 - k
    AnimFrame = AnimFrames(AnimStart(a) + k)
End Function

' ---------- loaders ----------

Private Sub LoadSpriteFile(ByVal fileName As String)
    Dim lines() As String, parts() As String, ln As String
    Dim pal(0 To 255) As Long, b() As Byte
    Dim i As Long, k As Long, x As Long, y As Long
    Dim id As Long, w As Long, h As Long, off As Long

    For k = 0 To 255: pal(k) = -1: Next k
    lines = ReadLines(DataPath(fileName & ".txt"))

    i = 0
    Do While i <= UBound(lines)
        ln = lines(i)
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            parts = Split(ln, " ")
            Select Case parts(0)
            Case "PAL"
                pal(Asc(parts(1))) = HexColor(parts(2))
            Case "SPRITE"
                w = CLng(parts(2)): h = CLng(parts(3))
                id = AddSprite(fileName & "." & parts(1), w, h, CLng(parts(4)), CLng(parts(5)))
                off = SprOff(id)
                For y = 0 To h - 1
                    i = i + 1
                    b = StrConv(lines(i), vbFromUnicode)
                    For x = 0 To w - 1
                        SprPix(off + y * w + x) = pal(b(x))
                    Next x
                Next y
            Case "ANIM"
                AddAnim fileName, parts
            End Select
        End If
        i = i + 1
    Loop
End Sub

Private Sub LoadStageFile(ByVal fileName As String, ByVal idx As Long)
    Dim lines() As String, parts() As String, ln As String
    Dim pal(0 To 255) As Long, b() As Byte
    Dim i As Long, k As Long, x As Long, y As Long

    For k = 0 To 255: pal(k) = 0: Next k
    lines = ReadLines(DataPath(fileName & ".txt"))

    i = 0
    Do While i <= UBound(lines)
        ln = lines(i)
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            parts = Split(ln, " ")
            Select Case parts(0)
            Case "PAL"
                pal(Asc(parts(1))) = HexColor(parts(2))
            Case "STAGE"
                If idx = 0 Then
                    StageW = CLng(parts(2)): StageH = CLng(parts(3))
                    ReDim StagePix(0 To StageCount - 1, 0 To StageW - 1, 0 To StageH - 1)
                ElseIf CLng(parts(2)) <> StageW Or CLng(parts(3)) <> StageH Then
                    Err.Raise ERR_BASE + 518, "Data", fileName & ": stage size differs from the first stage"
                End If
                For y = 0 To StageH - 1
                    i = i + 1
                    b = StrConv(lines(i), vbFromUnicode)
                    For x = 0 To StageW - 1
                        StagePix(idx, x, y) = pal(b(x))
                    Next x
                Next y
            End Select
        End If
        i = i + 1
    Loop
End Sub

Private Function AddSprite(ByVal nm As String, ByVal w As Long, ByVal h As Long, ByVal ax As Long, ByVal ay As Long) As Long
    Dim id As Long, n As Long
    id = SprCount
    If id > UBound(SprW) Then
        n = UBound(SprW) * 2 + 1
        ReDim Preserve SprW(0 To n): ReDim Preserve SprH(0 To n)
        ReDim Preserve SprAX(0 To n): ReDim Preserve SprAY(0 To n)
        ReDim Preserve SprOff(0 To n): ReDim Preserve SprName(0 To n)
    End If
    Do While mPixUsed + w * h - 1 > UBound(SprPix)
        ReDim Preserve SprPix(0 To UBound(SprPix) * 2 + 1)
    Loop
    SprW(id) = w: SprH(id) = h: SprAX(id) = ax: SprAY(id) = ay
    SprOff(id) = mPixUsed: SprName(id) = nm
    mPixUsed = mPixUsed + w * h
    mSprIdx.Add id, nm
    SprCount = SprCount + 1
    AddSprite = id
End Function

' parts: ANIM <name> <ticksPerFrame> <frame1> <frame2> ...
Private Sub AddAnim(ByVal fileName As String, parts() As String)
    Dim a As Long, k As Long, n As Long
    a = AnimCount
    If a > UBound(AnimTPF) Then
        n = UBound(AnimTPF) * 2 + 1
        ReDim Preserve AnimTPF(0 To n): ReDim Preserve AnimStart(0 To n): ReDim Preserve AnimLen(0 To n)
    End If
    AnimTPF(a) = CLng(parts(2))
    AnimStart(a) = mFramesUsed
    AnimLen(a) = UBound(parts) - 2
    For k = 3 To UBound(parts)
        If mFramesUsed > UBound(AnimFrames) Then
            ReDim Preserve AnimFrames(0 To UBound(AnimFrames) * 2 + 1)
            ReDim Preserve mFrameNames(0 To UBound(AnimFrames))
        End If
        mFrameNames(mFramesUsed) = fileName & "." & parts(k)
        mFramesUsed = mFramesUsed + 1
    Next k
    mAnimIdx.Add a, fileName & "." & parts(1)
    AnimCount = AnimCount + 1
End Sub

Private Sub LoadMoves()
    Dim lines() As String, p() As String, ln As String, i As Long, n As Long, id As Long

    lines = ReadLines(DataPath("moves.txt"))
    n = UBound(lines) + 1
    ReDim MvKind(0 To n): ReDim MvStartup(0 To n): ReDim MvActive(0 To n): ReDim MvRecovery(0 To n)
    ReDim MvDamage(0 To n): ReDim MvHitstun(0 To n): ReDim MvBlockstun(0 To n): ReDim MvLevel(0 To n)
    ReDim MvKD(0 To n)
    ReDim MvHbX(0 To n): ReDim MvHbY(0 To n): ReDim MvHbW(0 To n): ReDim MvHbH(0 To n)
    ReDim MvSprS(0 To n): ReDim MvSprA(0 To n)
    ReDim MvType(0 To n): ReDim MvP1(0 To n): ReDim MvP2(0 To n): ReDim MvCmd(0 To n)
    ReDim mMvKey(0 To n)
    Set mMoveIdx = New Collection
    MvCount = 0

    For i = 0 To UBound(lines)
        ln = Trim$(Replace(lines(i), vbTab, " "))
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            Do While InStr(ln, "  ") > 0
                ln = Replace(ln, "  ", " ")
            Loop
            p = Split(ln, " ")
            If UBound(p) < 20 Then Err.Raise ERR_BASE + 517, "Data", "moves.txt line " & (i + 1) & ": expected 21 columns"
            id = MvCount
            Select Case p(2)
                Case "S": MvKind(id) = MK_STAND
                Case "C": MvKind(id) = MK_CROUCH
                Case Else: MvKind(id) = MK_AIR
            End Select
            MvStartup(id) = CLng(p(3)): MvActive(id) = CLng(p(4)): MvRecovery(id) = CLng(p(5))
            MvDamage(id) = CLng(p(6)): MvHitstun(id) = CLng(p(7)): MvBlockstun(id) = CLng(p(8))
            Select Case p(9)
                Case "L": MvLevel(id) = LV_LOW
                Case "O": MvLevel(id) = LV_OVERHEAD
                Case Else: MvLevel(id) = LV_HIGH
            End Select
            MvKD(id) = (p(10) = "1")
            MvHbX(id) = CLng(p(11)): MvHbY(id) = CLng(p(12)): MvHbW(id) = CLng(p(13)): MvHbH(id) = CLng(p(14))
            MvSprS(id) = SpriteId(p(0) & "." & p(15))
            MvSprA(id) = SpriteId(p(0) & "." & p(16))
            Select Case p(17)
                Case "T": MvType(id) = MT_THROW
                Case "P": MvType(id) = MT_PROJ
                Case "U": MvType(id) = MT_RISE
                Case "R": MvType(id) = MT_RUSH
                Case "G": MvType(id) = MT_GRAB
                Case "SP": MvType(id) = MT_SUPER_PROJ
                Case "SG": MvType(id) = MT_SUPER_GRAB
                Case "SR": MvType(id) = MT_SUPER_RUSH
                Case "TP": MvType(id) = MT_TELEPORT
                Case "ST": MvType(id) = MT_SUPER_TELE
                Case "PB": MvType(id) = MT_PROJ_BOOM
                Case "SF": MvType(id) = MT_SUPER_FIRE
                Case Else: MvType(id) = MT_NORMAL
            End Select
            MvP1(id) = Val(p(18))
            MvP2(id) = Val(p(19))
            If p(20) = "-" Then MvCmd(id) = "" Else MvCmd(id) = p(20)
            mMoveIdx.Add id, p(0) & "." & p(1)
            mMvKey(id) = p(0) & "." & p(1)
            MvCount = MvCount + 1
        End If
    Next i
End Sub

' ---------- helpers ----------

Private Function DataPath(ByVal fileName As String) As String
    If Len(BaseDir) > 0 Then
        DataPath = BaseDir & fileName
    Else
        DataPath = ThisWorkbook.Path & "\" & DATA_DIR & "\" & fileName
    End If
End Function

Private Function ReadLines(ByVal fullPath As String) As Variant      ' String array
    Dim f As Integer, s As String, b() As Byte
    If Len(Dir(fullPath)) = 0 Then Err.Raise ERR_BASE + 515, "Data", "File not found: " & fullPath
    f = FreeFile
    Open fullPath For Binary Access Read As #f
    If LOF(f) > 0 Then
        ReDim b(0 To LOF(f) - 1)
        Get #f, , b
        s = StrConv(b, vbUnicode)
    End If
    Close #f
    s = Replace(s, vbCr, "")
    ReadLines = Split(s, vbLf)
End Function

Private Function HexColor(ByVal hx As String) As Long
    HexColor = RGB(CLng("&H" & Mid$(hx, 1, 2)), CLng("&H" & Mid$(hx, 3, 2)), CLng("&H" & Mid$(hx, 5, 2)))
End Function
