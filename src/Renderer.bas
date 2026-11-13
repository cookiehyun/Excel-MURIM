Attribute VB_Name = "Renderer"
Option Explicit

' Cell-pixel renderer.
' Draw into a back buffer, then Flush() paints only cells that differ from
' what is already on the sheet (front buffer).
' Speed-ups (every Excel call is expensive, VBA loops are cheap but not free):
'   - BeginFrame copies the background with one array assignment
'   - every draw call records which part of each row it touched; Flush only
'     compares rows/columns touched this frame or the previous one
'   - horizontal runs of one color are merged, then identical runs on
'     consecutive rows are merged into rectangles ("B3:F9")
'   - all rectangles of one color go into one multi-area address
'     ("A1:C1,F4:G9,...") so a single Range(...).Interior.Color call paints them
' Headless = True (test harness) skips every sheet operation and instead checks
' that the painted cells really end up equal to the back buffer.

Public Const SCREEN_W As Long = 192
Public Const SCREEN_H As Long = 100
Public Const GROUND_Y As Long = 88        ' first floor row; feet stand on GROUND_Y - 1 (must match the stage files)
Private Const PIXEL_PT As Double = 3.75   ' cell size in points (= 5 screen pixels at 100% zoom)
Private Const ORG_ROW As Long = 1         ' sheet row of pixel y = 0
Private Const ORG_COL As Long = 1         ' sheet col of pixel x = 0
Private Const GAME_SHEET As String = "Game"
Private Const ADDR_MAX As Long = 240      ' Range() address strings must stay under 255 chars
Private Const MAX_SLOTS As Long = 64
Private Const ERR_BASE As Long = -2147221504   ' = vbObjectError

Public Headless As Boolean

Private mWs As Worksheet
Private mBack() As Long
Private mFront() As Long
Private mBg() As Long
Private mBgDark() As Long
Private mPaintCalls As Long
Private mRuns As Long
Private mRects As Long
Private mColName(0 To SCREEN_W - 1) As String
Private mSlotColor(0 To MAX_SLOTS - 1) As Long
Private mSlotAddr(0 To MAX_SLOTS - 1) As String
Private mSlots As Long
Public KeepScreenUpdating As Boolean      ' benchmark option: never switch ScreenUpdating off
' dirty spans per row: this frame (mD0..mD1) and the previous frame (mP0..mP1); x0 > x1 = clean
Private mD0(0 To SCREEN_H - 1) As Long
Private mD1(0 To SCREEN_H - 1) As Long
Private mP0(0 To SCREEN_H - 1) As Long
Private mP1(0 To SCREEN_H - 1) As Long
Private mFull As Boolean                  ' background changed: compare the whole screen
Private mLastDark As Long                 ' -1 unknown, 0 normal, 1 dark
' rectangle merging: runs of the previous row that may still grow downward
Private mOX0(0 To SCREEN_W) As Long, mOX1(0 To SCREEN_W) As Long, mOC(0 To SCREEN_W) As Long
Private mOY0(0 To SCREEN_W) As Long, mOn As Long
Private mNX0(0 To SCREEN_W) As Long, mNX1(0 To SCREEN_W) As Long, mNC(0 To SCREEN_W) As Long
Private mNY0(0 To SCREEN_W) As Long, mNn As Long
Private mSim() As Long                    ' headless: what the sheet would show

Public Sub Init()
    Dim x As Long, y As Long
    If Not Headless Then
        Set mWs = GetSheet(GAME_SHEET)
        mWs.Cells.Interior.Color = RGB(16, 16, 20)     ' dark frame around the play area
        SquareCells mWs, ORG_ROW, ORG_COL, SCREEN_H, SCREEN_W, PIXEL_PT
        mWs.Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        mWs.Cells(ORG_ROW + SCREEN_H + 2, ORG_COL).Select   ' keep selection box off the play area
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If
    For x = 0 To SCREEN_W - 1
        mColName(x) = ColumnLetters(ORG_COL + x)
    Next x

    ReDim mBack(0 To SCREEN_W - 1, 0 To SCREEN_H - 1)
    ReDim mFront(0 To SCREEN_W - 1, 0 To SCREEN_H - 1)
    ReDim mBg(0 To SCREEN_W - 1, 0 To SCREEN_H - 1)
    ReDim mBgDark(0 To SCREEN_W - 1, 0 To SCREEN_H - 1)
    ReDim mSim(0 To SCREEN_W - 1, 0 To SCREEN_H - 1)
    mLastDark = -1
    SetStage 0
    For y = 0 To SCREEN_H - 1
        For x = 0 To SCREEN_W - 1
            mFront(x, y) = -1     ' force full paint on first frame
            mSim(x, y) = -1
        Next x
        mD0(y) = SCREEN_W: mD1(y) = -1
        mP0(y) = SCREEN_W: mP1(y) = -1
    Next y
    mFull = True
End Sub

' Repaint every cell on the next Flush (e.g. after Excel was in the background).
Public Sub Invalidate()
    Dim x As Long, y As Long
    For y = 0 To SCREEN_H - 1
        For x = 0 To SCREEN_W - 1
            mFront(x, y) = -1
        Next x
    Next y
    mFull = True
End Sub

' Record that row y was drawn between x0 and x1 this frame.
Private Sub Mark(ByVal y As Long, ByVal x0 As Long, ByVal x1 As Long)
    If x0 < 0 Then x0 = 0
    If x1 > SCREEN_W - 1 Then x1 = SCREEN_W - 1
    If x0 > x1 Then Exit Sub
    If x0 < mD0(y) Then mD0(y) = x0
    If x1 > mD1(y) Then mD1(y) = x1
End Sub

Public Property Get PaintCalls() As Long
    PaintCalls = mPaintCalls
End Property

Public Property Get RectCount() As Long
    RectCount = mRects
End Property

Public Property Get RunCount() As Long
    RunCount = mRuns
End Property

Public Function BackPixel(ByVal x As Long, ByVal y As Long) As Long
    BackPixel = mBack(x, y)
End Function

' dark = True draws the dimmed stage (super freeze, menu).
Public Sub BeginFrame(Optional ByVal dark As Boolean = False)
    Dim d As Long
    If dark Then d = 1 Else d = 0
    If Headless Then
        CopyLoop d                         ' LibreOffice shares arrays on assignment: copy element by element
    ElseIf d = 1 Then
        mBack = mBgDark                    ' Excel copies the whole array natively (much faster than a loop)
    Else
        mBack = mBg
    End If
    If d <> mLastDark Then mFull = True     ' switched between normal and dimmed stage
    mLastDark = d
End Sub

Private Sub CopyLoop(ByVal d As Long)
    Dim x As Long, y As Long
    For y = 0 To SCREEN_H - 1
        For x = 0 To SCREEN_W - 1
            If d = 1 Then mBack(x, y) = mBgDark(x, y) Else mBack(x, y) = mBg(x, y)
        Next x
    Next y
End Sub

Private Function Darken(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = c And &HFF&
    g = (c \ &H100&) And &HFF&
    b = (c \ &H10000) And &HFF&
    Darken = RGB(r * 0.3, g * 0.3, b * 0.3 + 20)
End Function

Public Sub PutPixel(ByVal x As Long, ByVal y As Long, ByVal color As Long)
    If x >= 0 And x < SCREEN_W And y >= 0 And y < SCREEN_H Then
        mBack(x, y) = color
        If x < mD0(y) Then mD0(y) = x
        If x > mD1(y) Then mD1(y) = x
    End If
End Sub

Public Sub FillRect(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal color As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, i As Long, j As Long
    x0 = x: y0 = y: x1 = x + w - 1: y1 = y + h - 1
    If x0 < 0 Then x0 = 0
    If y0 < 0 Then y0 = 0
    If x1 > SCREEN_W - 1 Then x1 = SCREEN_W - 1
    If y1 > SCREEN_H - 1 Then y1 = SCREEN_H - 1
    For j = y0 To y1
        For i = x0 To x1
            mBack(i, j) = color
        Next i
        Mark j, x0, x1
    Next j
End Sub

' 1px rectangle outline (debug boxes).
Public Sub DrawRectOutline(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal color As Long)
    If w <= 0 Or h <= 0 Then Exit Sub
    FillRect x, y, w, 1, color
    FillRect x, y + h - 1, w, 1, color
    FillRect x, y, 1, h, color
    FillRect x + w - 1, y, 1, h, color
End Sub

' Draw sprite so that its anchor (ground point under the hip) lands on (x, y).
' Sprites face right; flip = True mirrors around the anchor.
' tint >= 0 paints every opaque pixel with that color (hit flash).
Public Sub DrawSprite(ByVal id As Long, ByVal x As Long, ByVal y As Long, ByVal flip As Boolean, _
                      Optional ByVal tint As Long = -1)
    Dim w As Long, h As Long, ax As Long, ay As Long, off As Long
    Dim sx As Long, sy As Long, px As Long, py As Long, c As Long
    w = Data.SprW(id): h = Data.SprH(id)
    ax = Data.SprAX(id): ay = Data.SprAY(id): off = Data.SprOff(id)
    For sy = 0 To h - 1
        py = y - ay + sy
        If py >= 0 And py < SCREEN_H Then
            If flip Then
                Mark py, x + ax - (w - 1), x + ax
            Else
                Mark py, x - ax, x - ax + w - 1
            End If
            For sx = 0 To w - 1
                c = Data.SprPix(off + sy * w + sx)
                If c >= 0 Then
                    If flip Then px = x + ax - sx Else px = x - ax + sx
                    If px >= 0 And px < SCREEN_W Then
                        If tint >= 0 Then mBack(px, py) = tint Else mBack(px, py) = c
                    End If
                End If
            Next sx
        End If
    Next sy
End Sub

' Draw a sprite enlarged by an integer factor, clipped to (cx0, cy0)-(cx1, cy1).
' Used by the select screens (portraits and the big character view).
Public Sub DrawSpriteScaled(ByVal id As Long, ByVal x As Long, ByVal y As Long, ByVal flip As Boolean, _
                            ByVal k As Long, ByVal cx0 As Long, ByVal cy0 As Long, ByVal cx1 As Long, ByVal cy1 As Long, _
                            Optional ByVal tint As Long = -1)
    Dim w As Long, h As Long, ax As Long, ay As Long, off As Long
    Dim sx As Long, sy As Long, px As Long, py As Long, c As Long, i As Long, j As Long
    w = Data.SprW(id): h = Data.SprH(id)
    ax = Data.SprAX(id): ay = Data.SprAY(id): off = Data.SprOff(id)
    If cx0 < 0 Then cx0 = 0
    If cy0 < 0 Then cy0 = 0
    If cx1 > SCREEN_W - 1 Then cx1 = SCREEN_W - 1
    If cy1 > SCREEN_H - 1 Then cy1 = SCREEN_H - 1
    For j = cy0 To cy1                     ' the clip box bounds everything this call can touch
        Mark j, cx0, cx1
    Next j
    For sy = 0 To h - 1
        For sx = 0 To w - 1
            c = Data.SprPix(off + sy * w + sx)
            If c >= 0 Then
                If tint >= 0 Then c = tint
                If flip Then px = x + (ax - sx) * k - (k - 1) Else px = x + (sx - ax) * k
                py = y + (sy - ay) * k
                For j = py To py + k - 1
                    If j >= cy0 And j <= cy1 Then
                        For i = px To px + k - 1
                            If i >= cx0 And i <= cx1 Then mBack(i, j) = c
                        Next i
                    End If
                Next j
            End If
        Next sx
    Next sy
End Sub

' Paint differences to the sheet. Returns number of Range paint calls.
Public Function Flush() As Long
    Dim x As Long, y As Long, x0 As Long, c As Long, k As Long, a0 As Long, a1 As Long
    mPaintCalls = 0: mRuns = 0: mRects = 0: mSlots = 0: mOn = 0
    If Not Headless And Not KeepScreenUpdating Then Application.ScreenUpdating = False
    For y = 0 To SCREEN_H - 1
        mNn = 0
        ' columns that can differ: drawn now, drawn last frame, or everything after a background change
        If mFull Then
            a0 = 0: a1 = SCREEN_W - 1
        Else
            a0 = mD0(y): a1 = mD1(y)
            If mP0(y) < a0 Then a0 = mP0(y)
            If mP1(y) > a1 Then a1 = mP1(y)
        End If
        x = a0
        Do While x <= a1
            If mBack(x, y) <> mFront(x, y) Then
                c = mBack(x, y)
                x0 = x
                ' extend run while color matches (painting already-correct cells is harmless)
                Do While x < SCREEN_W
                    If mBack(x, y) <> c Then Exit Do
                    mFront(x, y) = c
                    x = x + 1
                Loop
                AddRun c, y, x0, x - 1
            Else
                x = x + 1
            End If
        Loop
        CloseOpen y                        ' runs of row y-1 that did not continue into row y
        mP0(y) = mD0(y): mP1(y) = mD1(y)
        mD0(y) = SCREEN_W: mD1(y) = -1
    Next y
    CloseOpen SCREEN_H
    For k = 0 To mSlots - 1
        PaintSlot k
    Next k
    If Not Headless And Not KeepScreenUpdating Then Application.ScreenUpdating = True
    mFull = False
    If Headless Then VerifySim
    Flush = mPaintCalls
End Function

' A run on row y: grow the matching rectangle from row y-1, or start a new one.
Private Sub AddRun(ByVal c As Long, ByVal y As Long, ByVal x0 As Long, ByVal x1 As Long)
    Dim k As Long
    mRuns = mRuns + 1
    For k = 0 To mOn - 1
        If mOX0(k) = x0 Then
            If mOX1(k) = x1 And mOC(k) = c Then
                mNX0(mNn) = x0: mNX1(mNn) = x1: mNC(mNn) = c: mNY0(mNn) = mOY0(k)
                mNn = mNn + 1
                mOX0(k) = -1                ' taken over
                Exit Sub
            End If
        End If
    Next k
    mNX0(mNn) = x0: mNX1(mNn) = x1: mNC(mNn) = c: mNY0(mNn) = y
    mNn = mNn + 1
End Sub

' Emit the rectangles that ended on row y-1, then make row y's runs the open set.
Private Sub CloseOpen(ByVal y As Long)
    Dim k As Long
    For k = 0 To mOn - 1
        If mOX0(k) >= 0 Then AddRect mOC(k), mOY0(k), y - 1, mOX0(k), mOX1(k)
    Next k
    For k = 0 To mNn - 1
        mOX0(k) = mNX0(k): mOX1(k) = mNX1(k): mOC(k) = mNC(k): mOY0(k) = mNY0(k)
    Next k
    mOn = mNn
    mNn = 0
End Sub

Private Sub AddRect(ByVal c As Long, ByVal y0 As Long, ByVal y1 As Long, ByVal x0 As Long, ByVal x1 As Long)
    Dim a As String, k As Long, i As Long, j As Long
    mRects = mRects + 1
    If x0 = x1 And y0 = y1 Then
        a = mColName(x0) & CStr(ORG_ROW + y0)
    Else
        a = mColName(x0) & CStr(ORG_ROW + y0) & ":" & mColName(x1) & CStr(ORG_ROW + y1)
    End If
    If Headless Then
        For j = y0 To y1
            For i = x0 To x1
                mSim(i, j) = c
            Next i
        Next j
    End If
    For k = 0 To mSlots - 1
        If mSlotColor(k) = c Then Exit For
    Next k
    If k = mSlots Then
        If mSlots = MAX_SLOTS Then
            ' too many colors this frame: flush everything and start over
            For k = 0 To mSlots - 1
                PaintSlot k
            Next k
            mSlots = 0
            k = 0
        End If
        mSlotColor(k) = c
        mSlotAddr(k) = ""
        mSlots = mSlots + 1
    End If
    If Len(mSlotAddr(k)) + Len(a) + 1 > ADDR_MAX Then PaintSlot k
    If Len(mSlotAddr(k)) = 0 Then mSlotAddr(k) = a Else mSlotAddr(k) = mSlotAddr(k) & "," & a
End Sub

' Test harness: the simulated sheet must equal the back buffer after every flush.
Private Sub VerifySim()
    Dim x As Long, y As Long
    For y = 0 To SCREEN_H - 1
        For x = 0 To SCREEN_W - 1
            If mSim(x, y) <> mBack(x, y) Then
                Err.Raise ERR_BASE + 600, "Renderer", "paint mismatch at " & x & "," & y
            End If
        Next x
    Next y
End Sub

Private Sub PaintSlot(ByVal k As Long)
    If Len(mSlotAddr(k)) = 0 Then Exit Sub
    If Not Headless Then mWs.Range(mSlotAddr(k)).Interior.Color = mSlotColor(k)
    mPaintCalls = mPaintCalls + 1
    mSlotAddr(k) = ""
End Sub

' ---------- setup helpers (also used by Gallery) ----------

Public Function GetSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = sheetName
    End If
    Set GetSheet = ws
End Function

' Make an nRows x nCols block of square cells, pt points per side.
Public Sub SquareCells(ByVal ws As Worksheet, ByVal r0 As Long, ByVal c0 As Long, _
                       ByVal nRows As Long, ByVal nCols As Long, ByVal pt As Double)
    Dim cw As Double, w As Double, i As Long
    ws.Rows(r0).Resize(nRows).RowHeight = pt
    ' ColumnWidth is in character units; iterate until Width (points) ~= pt
    cw = 1
    For i = 1 To 6
        ws.Columns(c0).ColumnWidth = cw
        w = ws.Columns(c0).Width
        If w <= 0 Then Exit For
        cw = cw * pt / w
        If cw < 0.05 Then cw = 0.05
    Next i
    ws.Columns(c0).Resize(, nCols).ColumnWidth = cw
End Sub

Private Function ColumnLetters(ByVal col As Long) As String
    Dim s As String, n As Long
    n = col
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColumnLetters = s
End Function

' Switch the background to stage k (the diff renderer repaints what changed).
Public Sub SetStage(ByVal k As Long)
    Dim x As Long, y As Long
    mFull = True
    BuildBackground k
    For y = 0 To SCREEN_H - 1
        For x = 0 To SCREEN_W - 1
            mBgDark(x, y) = Darken(mBg(x, y))
        Next x
    Next y
End Sub

' Half-size (or 1/stp) preview of stage k with its top-left at (x0, y0).
Public Sub DrawStageThumb(ByVal k As Long, ByVal x0 As Long, ByVal y0 As Long, ByVal stp As Long)
    Dim x As Long, y As Long
    For y = 0 To SCREEN_H \ stp - 1
        For x = 0 To SCREEN_W \ stp - 1
            PutPixel x0 + x, y0 + y, Data.StagePix(k, x * stp, y * stp)
        Next x
    Next y
End Sub

Private Sub BuildBackground(ByVal k As Long)
    Dim x As Long, y As Long, t As Double, c As Long
    If Data.StageW = SCREEN_W And Data.StageH = SCREEN_H And k < Data.StageCount Then
        For y = 0 To SCREEN_H - 1
            For x = 0 To SCREEN_W - 1
                mBg(x, y) = Data.StagePix(k, x, y)
            Next x
        Next y
        Exit Sub
    End If
    ' fallback: plain gradient if no stage data matches the screen size
    For y = 0 To SCREEN_H - 1
        If y < GROUND_Y Then
            t = y / (GROUND_Y - 1)
            c = RGB(40 + 60 * t, 60 + 80 * t, 120 + 90 * t)
        Else
            c = RGB(60, 45, 35)
        End If
        For x = 0 To SCREEN_W - 1
            mBg(x, y) = c
        Next x
    Next y
End Sub
