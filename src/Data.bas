Attribute VB_Name = "Data"
Option Explicit

' Loads text data from <workbook folder>\data\ :
'   <fighter>.txt       PAL / SPRITE / ANIM records (see tools/gen_sprites.py)
'   stage_<name>.txt    PAL / STAGE records (see tools/gen_stage.py)
' Sprites are stored in one flat pixel pool; color -1 = transparent.
' Names are "<file>.<sprite>", e.g. "danwoo.idle_0".

Private Const DATA_DIR As String = "data"

Public SprCount As Long
Public SprW() As Long, SprH() As Long, SprAX() As Long, SprAY() As Long, SprOff() As Long
Public SprName() As String
Public SprPix() As Long
Private mPixUsed As Long

Public AnimCount As Long
Public AnimTPF() As Long, AnimStart() As Long, AnimLen() As Long
Public AnimFrames() As Long
Private mFramesUsed As Long

Public StageW As Long, StageH As Long
Public StagePix() As Long          ' (x, y)

Private mSprIdx As Collection
Private mAnimIdx As Collection

Public Sub LoadAll()
    SprCount = 0: mPixUsed = 0: AnimCount = 0: mFramesUsed = 0
    StageW = 0: StageH = 0
    ReDim SprW(0 To 63): ReDim SprH(0 To 63): ReDim SprAX(0 To 63): ReDim SprAY(0 To 63)
    ReDim SprOff(0 To 63): ReDim SprName(0 To 63)
    ReDim SprPix(0 To 65535)
    ReDim AnimTPF(0 To 31): ReDim AnimStart(0 To 31): ReDim AnimLen(0 To 31)
    ReDim AnimFrames(0 To 255)
    Set mSprIdx = New Collection
    Set mAnimIdx = New Collection

    LoadSpriteFile "danwoo"
    LoadSpriteFile "cheolsan"
    LoadStageFile "stage_sanmun"
End Sub

Public Function SpriteId(ByVal nm As String) As Long
    On Error GoTo Missing
    SpriteId = mSprIdx(nm)
    Exit Function
Missing:
    Err.Raise vbObjectError + 513, "Data", "Sprite not found: " & nm
End Function

Public Function AnimId(ByVal nm As String) As Long
    On Error GoTo Missing
    AnimId = mAnimIdx(nm)
    Exit Function
Missing:
    Err.Raise vbObjectError + 514, "Data", "Anim not found: " & nm
End Function

' Sprite id of the frame shown at tick t of an animation (loops).
Public Function AnimFrame(ByVal a As Long, ByVal t As Long) As Long
    AnimFrame = AnimFrames(AnimStart(a) + ((t \ AnimTPF(a)) Mod AnimLen(a)))
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

Private Sub LoadStageFile(ByVal fileName As String)
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
                StageW = CLng(parts(2)): StageH = CLng(parts(3))
                ReDim StagePix(0 To StageW - 1, 0 To StageH - 1)
                For y = 0 To StageH - 1
                    i = i + 1
                    b = StrConv(lines(i), vbFromUnicode)
                    For x = 0 To StageW - 1
                        StagePix(x, y) = pal(b(x))
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
        If mFramesUsed > UBound(AnimFrames) Then ReDim Preserve AnimFrames(0 To UBound(AnimFrames) * 2 + 1)
        AnimFrames(mFramesUsed) = SpriteId(fileName & "." & parts(k))
        mFramesUsed = mFramesUsed + 1
    Next k
    mAnimIdx.Add a, fileName & "." & parts(1)
    AnimCount = AnimCount + 1
End Sub

' ---------- helpers ----------

Private Function DataPath(ByVal fileName As String) As String
    DataPath = ThisWorkbook.Path & "\" & DATA_DIR & "\" & fileName
End Function

Private Function ReadLines(ByVal fullPath As String) As String()
    Dim f As Integer, s As String
    If Len(Dir(fullPath)) = 0 Then Err.Raise vbObjectError + 515, "Data", "File not found: " & fullPath
    f = FreeFile
    Open fullPath For Binary Access Read As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    s = Replace(s, vbCr, "")
    ReadLines = Split(s, vbLf)
End Function

Private Function HexColor(ByVal hx As String) As Long
    HexColor = RGB(CLng("&H" & Mid$(hx, 1, 2)), CLng("&H" & Mid$(hx, 3, 2)), CLng("&H" & Mid$(hx, 5, 2)))
End Function
