Attribute VB_Name = "Gallery"
Option Explicit

' Lays out every loaded sprite on a "Gallery" sheet for visual checking.
' Run: Alt+F8 -> ShowGallery

Private Const PX As Double = 4        ' cell size (points)
Private Const PER_ROW As Long = 8
Private Const CELL_W As Long = 52     ' pixels per gallery slot
Private Const CELL_H As Long = 48     ' sprite rows per slot (label row is extra)

Public Sub ShowGallery()
    Dim ws As Worksheet, id As Long, slotRow As Long, slotCol As Long
    Dim r0 As Long, c0 As Long, nSlots As Long, nRowsTotal As Long

    On Error GoTo Fail
    Data.LoadAll
    Set ws = Renderer.GetSheet("Gallery")

    Application.ScreenUpdating = False
    ws.Cells.Clear
    nSlots = (Data.SprCount + PER_ROW - 1) \ PER_ROW
    nRowsTotal = nSlots * (CELL_H + 1)
    Renderer.SquareCells ws, 1, 1, nRowsTotal, PER_ROW * CELL_W, PX
    ws.Range(ws.Cells(1, 1), ws.Cells(nRowsTotal, PER_ROW * CELL_W)).Interior.Color = RGB(70, 80, 100)

    For id = 0 To Data.SprCount - 1
        slotRow = id \ PER_ROW
        slotCol = id Mod PER_ROW
        r0 = 1 + slotRow * (CELL_H + 1)
        c0 = 1 + slotCol * CELL_W
        With ws.Cells(r0, c0)
            .Value = Data.SprName(id)
            .Font.Size = 7
            .Font.Color = RGB(255, 255, 255)
        End With
        ws.Rows(r0).RowHeight = 10
        ' bottom-align sprites inside the slot so feet line up
        PaintSprite ws, id, r0 + 1 + (CELL_H - Data.SprH(id)), c0 + 2
    Next id

    Application.ScreenUpdating = True
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Gallery"
End Sub

Private Sub PaintSprite(ByVal ws As Worksheet, ByVal id As Long, ByVal r0 As Long, ByVal c0 As Long)
    Dim w As Long, h As Long, off As Long, x As Long, y As Long, x0 As Long, c As Long
    w = Data.SprW(id): h = Data.SprH(id): off = Data.SprOff(id)
    For y = 0 To h - 1
        x = 0
        Do While x < w
            c = Data.SprPix(off + y * w + x)
            If c >= 0 Then
                x0 = x
                Do While x < w
                    If Data.SprPix(off + y * w + x) <> c Then Exit Do
                    x = x + 1
                Loop
                ws.Range(ws.Cells(r0 + y, c0 + x0), ws.Cells(r0 + y, c0 + x - 1)).Interior.Color = c
            Else
                x = x + 1
            End If
        Loop
    Next y
End Sub
