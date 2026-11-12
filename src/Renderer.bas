Attribute VB_Name = "Renderer"
Option Explicit

' Cell-pixel renderer.
' Draw into a back buffer, then Flush() paints only cells that differ from
' what is already on the sheet (front buffer).
' Paint batching: horizontal runs of one color are merged, and all runs of the
' same color are joined into one multi-area address ("A1:C1,F4:G4,...") so a
' single Range(...).Interior.Color call paints many runs.
' Headless = True (test harness) skips every sheet operation.

Public Const SCREEN_W As Long = 192
Public Const SCREEN_H As Long = 100
Public Const GROUND_Y As Long = 88        ' first floor row; feet stand on GROUND_Y - 1 (must match the stage files)
Private Const PIXEL_PT As Double = 3.75   ' cell size in points (= 5 screen pixels at 100% zoom)
Private Const ORG_ROW As Long = 1         ' sheet row of pixel y = 0
Private Const ORG_COL As Long = 1         ' sheet col of pixel x = 0
Private Const GAME_SHEET As String = "Game"
Private Const ADDR_MAX As Long = 240      ' Range() address strings must stay under 255 chars
Private Const MAX_SLOTS As Long = 64

Public Headless As Boolean

Private mWs As Worksheet
Private mBack() As Long
Private mFront() As Long
Private mBg() As Long
Private mBgDark() As Long
Private mPaintCalls As Long
Private mRuns As Long
Private mColName(0 To SCREEN_W - 1) As String
Private mSlotColor(0 To MAX_SLOTS - 1) As Long
Private mSlotAddr(0 To MAX_SLOTS - 1) As String
Private mSlots As Long

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
    SetStage 0
    For y = 0 To SCREEN_H - 1
        For x = 0 To SCREEN_W - 1
            mFront(x, y) = -1     ' force full paint on first frame
        Next x
    Next y
End Sub

Public Property Get PaintCalls() As Long
    PaintCalls = mPaintCalls
End Property

Public Property Get RunCount() As Long
    RunCount = mRuns
End Property

Public Function BackPixel(ByVal x As Long, ByVal y As Long) As Long
    BackPixel = mBack(x, y)
End Function

' dark = True draws the dimmed stage (super freeze, menu).
Public Sub BeginFrame(Optional ByVal dark As Boolean = False)
    Dim x As Long, y As Long
    If dark Then
        For y = 0 To SCREEN_H - 1
            For x = 0 To SCREEN_W - 1
                mBack(x, y) = mBgDark(x, y)
            Next x
        Next y
    Else
        For y = 0 To SCREEN_H - 1
            For x = 0 To SCREEN_W - 1
                mBack(x, y) = mBg(x, y)
            Next x
        Next y
    End If
End Sub

Private Function Darken(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = c And &HFF&
    g = (c \ &H100&) And &HFF&
    b = (c \ &H10000) And &HFF&
    Darken = RGB(r * 0.3, g * 0.3, b * 0.3 + 20)
End Function

Public Sub PutPixel(ByVal x As Long, ByVal y As Long, ByVal color As Long)
    If x >= 0 And x < SCREEN_W And y >= 0 And y < SCREEN_H Then mBack(x, y) = color
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
    Dim x As Long, y As Long, x0 As Long, c As Long, k As Long
    mPaintCalls = 0: mRuns = 0: mSlots = 0
    If Not Headless Then Application.ScreenUpdating = False
    For y = 0 To SCREEN_H - 1
        x = 0
        Do While x < SCREEN_W
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
    Next y
    For k = 0 To mSlots - 1
        PaintSlot k
    Next k
    If Not Headless Then Application.ScreenUpdating = True
    Flush = mPaintCalls
End Function

Private Sub AddRun(ByVal c As Long, ByVal y As Long, ByVal x0 As Long, ByVal x1 As Long)
    Dim a As String, k As Long, r As String
    r = CStr(ORG_ROW + y)
    If x0 = x1 Then a = mColName(x0) & r Else a = mColName(x0) & r & ":" & mColName(x1) & r
    mRuns = mRuns + 1

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
