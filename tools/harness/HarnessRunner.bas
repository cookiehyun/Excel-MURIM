Attribute VB_Name = "HarnessRunner"
Option Explicit

' Drives the real game modules tick by tick from a text script (headless).
' Script lines:
'   <ticks> <keys>      keys joined with "+": L R U D A S Z X TAB H J DD F1, "-" = none
'   CMD control <i>     which fighter the keys control
'   CMD start <human> <level> <watch> <seed>   start a match (skips the menu)
'   CMD dummy <mode>    0 CPU, 1 idle, 2 guard, 3 attack
'   CMD meter <i> <n>   set super gauge
'   CMD place <x0> <x1> set positions
' Output file: "T <tick> <paintCalls> <runs> <hitstop> | <fighter0> | <fighter1>" per tick,
'              "F <tick>" + SCREEN_H rows of hex BGR pixels for dumped frames.

Public Function RunScript(ByVal dataDir As String, ByVal script As String, ByVal outPath As String, ByVal dumpEvery As Long) As String
    Dim lines() As String, p() As String, ks() As String, ln As String, keys As String
    Dim i As Long, n As Long, t As Long, k As Long, f As Integer, tick As Long

    On Error GoTo Fail
    Data.BaseDir = dataDir
    Game.GameInit True
    f = FreeFile
    Open outPath For Output As #f
    lines = Split(Replace(script, vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        ln = Trim(lines(i))
        If Len(ln) > 0 And Left(ln, 1) <> "#" Then
            p = Split(ln, " ")
            If p(0) = "CMD" Then
                Select Case p(1)
                Case "start"
                    ' CMD start <human 0/1> <level 0-2> <watch 0/1> <seed>
                    Game.HumanIndex = CLng(p(2))
                    Game.CpuLevel = CLng(p(3))
                    Game.WatchMode = (p(4) = "1")
                    Game.AutoRestart = True
                    Game.StartMatch CLng(p(5))
                Case "meter"
                    Fighter.F(CLng(p(2))).meter = CLng(p(3))
                Case "control": Fighter.SetControl CLng(p(2))
                Case "dummy": Fighter.SetDummy CLng(p(2))
                Case "place"
                    Fighter.F(0).x = CDbl(p(2))
                    Fighter.F(1).x = CDbl(p(3))
                End Select
            Else
                n = CLng(p(0))
                keys = "-"
                If UBound(p) >= 1 Then keys = p(1)
                For t = 1 To n
                    KeyInput.ClearNext
                    If keys <> "-" Then
                        ks = Split(keys, "+")
                        For k = 0 To UBound(ks)
                            KeyInput.SetNextByName ks(k)
                        Next k
                    End If
                    Game.GameTick
                    Game.GameRender
                    tick = tick + 1
                    Print #f, "T " & tick & " " & Renderer.PaintCalls & " " & Renderer.RunCount & " " & Combat.HitStop & _
                              " " & Game.GameState & " " & Combat.ComboCount(0) & " " & Combat.ComboCount(1) & _
                              " | " & FighterLine(0) & " | " & FighterLine(1)
                    If dumpEvery > 0 Then
                        If tick Mod dumpEvery = 0 Then DumpFrame f, tick
                    End If
                Next t
            End If
        End If
    Next i
    Close #f
    RunScript = "ok " & tick
    Exit Function
Fail:
    RunScript = "ERR " & Err.Number & " " & Err.Description & " (tick " & tick & ", script line " & (i + 1) & ")"
    On Error Resume Next
    Close #f
End Function

Private Function FighterLine(ByVal i As Long) As String
    With Fighter.F(i)
        FighterLine = .prefix & " st=" & .state & " t=" & .stTick & " x=" & Format(.x, "0.0") & _
            " y=" & Format(.footY, "0.0") & " hp=" & .health & " stun=" & .stun & " chain=" & .chain & _
            " face=" & .facing & " mv=" & .mv & " meter=" & .meter & " dir=" & .inDir & " press=" & .inPress
    End With
End Function

Private Sub DumpFrame(ByVal f As Integer, ByVal tick As Long)
    Dim x As Long, y As Long, s As String
    Print #f, "F " & tick
    For y = 0 To SCREEN_H - 1
        s = ""
        For x = 0 To SCREEN_W - 1
            s = s & Right("00000" & Hex(Renderer.BackPixel(x, y)), 6)
        Next x
        Print #f, s
    Next y
End Sub
