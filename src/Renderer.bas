Attribute VB_Name = "Renderer"
Option Explicit

' Cell-pixel renderer.
' Draw into a back buffer, then Flush() paints only cells that differ from
' what is already on the sheet (front buffer). Horizontal runs of the same
' color are merged into one Range.Interior.Color call.

Public Const SCREEN_W As Long = 96
Public Const SCREEN_H As Long = 54
Public Const GROUND_Y As Long = 46        ' first ground row (feet stand on GROUND_Y - 1)
Private Const PIXEL_PT As Double = 6      ' cell size in points
Private Const ORG_ROW As Long = 1         ' sheet row of pixel y = 0
Private Const ORG_COL As Long = 1         ' sheet col of pixel x = 0
Private Const GAME_SHEET As String = "Game"

Private mWs As Worksheet
Private mBack() As Long
Private mFront() As Long
Private mBg() As Long
Private mPaintCalls As Long

Public Sub Init()
    Set mWs = GetGameSheet()
    SetupGrid
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

' ---------- setup ----------

Private Function GetGameSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(GAME_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = GAME_SHEET
    End If
    Set GetGameSheet = ws
End Function

Private Sub SetupGrid()
    Dim cw As Double, w As Double, i As Long
    With mWs
        .Cells.Interior.Color = RGB(16, 16, 20)   ' dark frame around the play area
        .Rows(ORG_ROW).Resize(SCREEN_H).RowHeight = PIXEL_PT

        ' ColumnWidth is in character units; iterate until Width (points) ~= PIXEL_PT
        cw = 1
        For i = 1 To 6
            .Columns(ORG_COL).ColumnWidth = cw
            w = .Columns(ORG_COL).Width
            If w <= 0 Then Exit For
            cw = cw * PIXEL_PT / w
            If cw < 0.05 Then cw = 0.05
        Next i
        .Columns(ORG_COL).Resize(, SCREEN_W).ColumnWidth = cw
    End With

    mWs.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    mWs.Cells(ORG_ROW + SCREEN_H + 2, ORG_COL).Select   ' keep selection box off the play area
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

Private Sub BuildBackground()
    Dim x As Long, y As Long, t As Double, c As Long
    For y = 0 To SCREEN_H - 1
        If y < GROUND_Y Then
            t = y / (GROUND_Y - 1)                      ' sky gradient top -> horizon
            c = RGB(40 + 60 * t, 60 + 80 * t, 120 + 90 * t)
        ElseIf y = GROUND_Y Then
            c = RGB(90, 70, 50)                         ' ground edge
        Else
            c = RGB(60, 45, 35)
        End If
        For x = 0 To SCREEN_W - 1
            mBg(x, y) = c
        Next x
    Next y
End Sub
