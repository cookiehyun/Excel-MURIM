Attribute VB_Name = "KeyConfig"
Option Explicit

' Multiplayer key settings: defaults, data\keys.txt (saved / loaded) and the KEY CONFIG screen.
'   up / down = row, left / right = 1P / 2P column, A = change (then press the new key, Esc = cancel), S = save + back
' A key that is already used by the other slot / player is swapped, so no two controls share a key.

Private Const KEY_FILE As String = "keys.txt"
Private Const ROWS As Long = 9                ' 8 controls + RESET

Private mRow As Long
Private mCol As Long
Private mWait As Boolean                      ' waiting for the new key
Private mTick As Long
Private mMsg As String
Private mMsgT As Long

' ---------------------------------------------------------------- defaults / file

Public Sub ResetDefaults()
    SetSet 1, Array(&H41, &H44, &H57, &H53, &H52, &H54, &H46, &H47)   ' 1P: A D W S | R T F G
    SetSet 2, Array(&H25, &H27, &H26, &H28, &H4B, &H4C, &HBC, &HBE)   ' 2P: arrows | K L , .
End Sub

Private Sub SetSet(ByVal ks As Long, ByVal vks As Variant)
    Dim k As Long
    For k = 0 To 7
        KeyInput.SetVK KeyInput.KeyOf(ks, k), vks(k)
    Next k
End Sub

' Defaults, then data\keys.txt if it exists ("1 vk vk vk vk vk vk vk vk" / "2 ...").
Public Sub Load()
    Dim lines As Variant, i As Long, p() As String, k As Long, ks As Long
    ResetDefaults
    On Error GoTo NoFile
    lines = Data.ReadFileLines(KEY_FILE)
    For i = 0 To UBound(lines)
        p = Split(Trim$(lines(i)), " ")
        If UBound(p) >= 8 Then
            ks = CLng(p(0))
            If ks = 1 Or ks = 2 Then
                For k = 0 To 7
                    KeyInput.SetVK KeyInput.KeyOf(ks, k), CLng(p(k + 1))
                Next k
            End If
        End If
    Next i
NoFile:
End Sub

Public Sub Save()
    Dim f As Integer, ks As Long, k As Long, ln As String
    On Error GoTo Fail
    f = FreeFile
    Open Data.FilePath(KEY_FILE) For Output As #f
    For ks = 1 To 2
        ln = CStr(ks)
        For k = 0 To 7
            ln = ln & " " & KeyInput.GetVK(KeyInput.KeyOf(ks, k))
        Next k
        Print #f, ln
    Next ks
    Close #f
    Exit Sub
Fail:
    mMsg = "COULD NOT SAVE": mMsgT = 40
End Sub

' ---------------------------------------------------------------- screen

Public Sub Begin()
    mRow = 0: mCol = 0: mWait = False: mMsgT = 0
    KeyInput.ScanKey                                    ' forget keys that are already down
End Sub

' Returns True when the screen is closed.
Public Function Tick() As Boolean
    Dim vk As Long
    mTick = mTick + 1
    If mMsgT > 0 Then mMsgT = mMsgT - 1
    If mWait Then
        vk = KeyInput.ScanKey()
        If vk = &H1B Then                               ' Esc = cancel
            mWait = False
        ElseIf vk > 0 Then
            Assign 1 + mCol, mRow, vk
            mWait = False
        End If
        Exit Function
    End If
    KeyInput.ScanKey                                    ' keep the scan state fresh
    If KeyInput.IsPressed(BTN_UP) Then mRow = (mRow + ROWS - 1) Mod ROWS
    If KeyInput.IsPressed(BTN_DOWN) Then mRow = (mRow + 1) Mod ROWS
    If KeyInput.IsPressed(BTN_LEFT) Or KeyInput.IsPressed(BTN_RIGHT) Then mCol = 1 - mCol
    If KeyInput.IsPressed(BTN_LP) Then
        If mRow = ROWS - 1 Then
            ResetDefaults
            mMsg = "DEFAULTS RESTORED": mMsgT = 30
        Else
            mWait = True
        End If
    ElseIf KeyInput.IsPressed(BTN_HP) Or KeyInput.IsPressed(BTN_QUIT) Then
        Save
        Tick = True
    End If
End Function

' Give control (ks, logical) the key vk; whoever had vk gets this control's old key.
Private Sub Assign(ByVal ks As Long, ByVal logical As Long, ByVal vk As Long)
    Dim b As Long, old As Long, s As Long, k As Long, ob As Long
    b = KeyInput.KeyOf(ks, logical)
    old = KeyInput.GetVK(b)
    For s = 1 To 2
        For k = 0 To 7
            ob = KeyInput.KeyOf(s, k)
            If ob <> b And KeyInput.GetVK(ob) = vk Then
                KeyInput.SetVK ob, old
                mMsg = "SWAPPED WITH " & s & "P " & ControlName(k): mMsgT = 40
            End If
        Next k
    Next s
    KeyInput.SetVK b, vk
End Sub

Private Function ControlName(ByVal k As Long) As String
    ControlName = Choose(k + 1, "LEFT", "RIGHT", "UP", "DOWN", "LIGHT P", "HEAVY P", "LIGHT K", "HEAVY K")
End Function

Public Function KeyName(ByVal vk As Long) As String
    Select Case vk
    Case &H30 To &H39, &H41 To &H5A: KeyName = Chr$(vk)
    Case &H60 To &H69: KeyName = "NUM" & (vk - &H60)
    Case &H25: KeyName = "LEFT"
    Case &H26: KeyName = "UP"
    Case &H27: KeyName = "RIGHT"
    Case &H28: KeyName = "DOWN"
    Case &H20: KeyName = "SPACE"
    Case &HD: KeyName = "ENTER"
    Case &H9: KeyName = "TAB"
    Case &H10, &HA0, &HA1: KeyName = "SHIFT"
    Case &H11, &HA2, &HA3: KeyName = "CTRL"
    Case &H12, &HA4, &HA5: KeyName = "ALT"
    Case &HBC: KeyName = ","
    Case &HBE: KeyName = "."
    Case &HBF: KeyName = "/"
    Case &HBD: KeyName = "-"
    Case &HBB: KeyName = "="
    Case &HBA: KeyName = "SEMI"
    Case &HDE: KeyName = "QUOTE"
    Case &HDB: KeyName = "LBRKT"
    Case &HDD: KeyName = "RBRKT"
    Case &H6A: KeyName = "NUM*"
    Case &H6B: KeyName = "NUMPLUS"
    Case &H6D: KeyName = "NUM-"
    Case &H6E: KeyName = "NUM."
    Case &H6F: KeyName = "NUM/"
    Case Else: KeyName = "KEY" & vk
    End Select
End Function

Public Sub Render()
    Dim r As Long, s As Long, y As Long, x As Long, c As Long, cx As Long, lbl As String, sel As Boolean
    Dim cGold As Long, cDim As Long, cP1 As Long, cP2 As Long
    cGold = RGB(240, 190, 80): cDim = RGB(130, 120, 150)
    cP1 = RGB(90, 170, 255): cP2 = RGB(255, 90, 80)
    cx = SCREEN_W \ 2
    Renderer.FillRect 0, 0, SCREEN_W, SCREEN_H, RGB(12, 10, 22)
    PixelFont.DrawBig "KEY CONFIG", cx, 2, cGold, 1
    PixelFont.DrawText "1P", 96, 12, cP1, 1, 1
    PixelFont.DrawText "2P", 150, 12, cP2, 1, 1
    For r = 0 To ROWS - 1
        y = 19 + r * 7
        If r = ROWS - 1 Then
            sel = (r = mRow)
            If sel Then Renderer.FillRect 70, y - 2, 106, 8, RGB(40, 28, 20)
            PixelFont.DrawText "RESET TO DEFAULTS", 123, y, IIf(sel, RGB(255, 236, 180), cDim), 1, 1
        Else
            PixelFont.DrawText ControlName(r), 12, y, cDim
            For s = 0 To 1
                x = 96 + s * 54
                sel = (r = mRow And s = mCol)
                If sel Then
                    Renderer.FillRect x - 24, y - 2, 48, 8, RGB(40, 28, 20)
                    Renderer.DrawRectOutline x - 24, y - 2, 48, 8, cGold
                End If
                If sel And mWait Then
                    If (mTick \ 4) Mod 2 = 0 Then lbl = "PRESS KEY" Else lbl = ""
                    c = RGB(255, 200, 90)
                Else
                    lbl = KeyName(KeyInput.GetVK(KeyInput.KeyOf(1 + s, r)))
                    If sel Then c = RGB(255, 236, 180) Else c = RGB(220, 210, 230)
                End If
                PixelFont.DrawText lbl, x, y, c, 1, 1
            Next s
        End If
    Next r
    If mMsgT > 0 Then PixelFont.DrawText mMsg, cx, 84, RGB(150, 220, 255), 1, 1
    If mWait Then
        PixelFont.DrawText "PRESS NEW KEY   ESC = CANCEL", cx, 93, RGB(150, 140, 170), 1, 1
    Else
        PixelFont.DrawText "A = CHANGE   S = SAVE   ARROWS = MOVE", cx, 93, RGB(150, 140, 170), 1, 1
    End If
End Sub
