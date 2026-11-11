Attribute VB_Name = "Renderer"
Option Explicit

' Cell-pixel renderer.
' Draw into a back buffer, then Flush() paints only cells that differ from
' what is already on the sheet (front buffer). Horizontal runs of the same
' color are merged into one Range.Interior.Color call.

Public Const SCREEN_W As Long = 120
Public Const SCREEN_H As Long = 64
Public Const GROUND_Y As Long = 56        ' first floor row; feet stand on GROUND_Y - 1
Private Const PIXEL_PT As Double = 5      ' cell size in points
Private Const ORG_ROW As Long = 1         ' sheet row of pixel y = 0
Private Const ORG_COL As Long = 1         ' sheet col of pixel x = 0
Private Const GAME_SHEET As String = "Game"

Private mWs As Worksheet
Private mBack() As Long
Private mFront() As Long
Private mBg() As Long
Private mPaintCalls As Long

Public Sub Init()
    Set mWs = GetSheet(GAME_SHEET)
    mWs.Cells.Interior.Color = RGB(16, 16, 20)     ' dark frame around the play area
    SquareCells mWs, ORG_ROW, ORG_COL, SCREEN_H, SCREEN_W, PIXEL_PT
    mWs.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    mWs.Cells(ORG_ROW + SCREEN_H + 2, ORG_COL).Select   ' keep selection box off the play area
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    ReDim mBack(0 To SCREEN_W - 1, 0 To SCREEN_H - 1)
    ReDim mFront(0 To SCREEN_W - 1, 0 To SCREEN_H - 1)
    ReDim mBg(0 To SCREEN_W - 1, 0 To SCREEN_H - 1)
    BuildBackground

    Dim x As Long, y As Long
    For y = 0 To SCREEN_H - 1
        For x = 0 To SCREEN_W - 1
            mFront(x, y) = -1     ' force full paint on first frame
        Next x
    Next y
End Sub

Public Property Get PaintCalls() As Long
    PaintCalls = mPaintCalls
End Property

Public Sub BeginFrame()
    Dim x As Long, y As Long
    For y = 0 To SCREEN_H - 1
        For x = 0 To SCREEN_W - 1
            mBack(x, y) = mBg(x, y)
        Next x
    Next y
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
Public Sub DrawSprite(ByVal id As Long, ByVal x As Long, ByVal y As Long, ByVal flip As Boolean)
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
                    If px >= 0 And px < SCREEN_W Then mBack(px, py) = c
                End If
            Next sx
        End If
    Next sy
End Sub

' Paint differences to the sheet. Returns number of Range paint calls.
Public Function Flush() As Long
    Dim x As Long, y As Long, x0 As Long, c As Long
    mPaintCalls = 0
    Application.ScreenUpdating = False
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
                mWs.Range(mWs.Cells(ORG_ROW + y, ORG_COL + x0), _
                          mWs.Cells(ORG_ROW + y, ORG_COL + x - 1)).Interior.Color = c
                mPaintCalls = mPaintCalls + 1
            Else
                x = x + 1
            End If
        Loop
    Next y
    Application.ScreenUpdating = True
    Flush = mPaintCalls
End Function

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

Private Sub BuildBackground()
    Dim x As Long, y As Long, t As Double, c As Long
    If Data.StageW = SCREEN_W And Data.StageH = SCREEN_H Then
        For y = 0 To SCREEN_H - 1
            For x = 0 To SCREEN_W - 1
                mBg(x, y) = Data.StagePix(x, y)
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
