Attribute VB_Name = "FX"
Option Explicit

' Short-lived visual effects (hit / guard sparks), drawn procedurally.

Public Const FX_HIT As Long = 0
Public Const FX_HEAVY As Long = 1
Public Const FX_BLOCK As Long = 2
Public Const FX_SLAM As Long = 3       ' ground impact (shock ring + dust)
Public Const FX_SLASH As Long = 4      ' sword cut streak
Public Const FX_DARK As Long = 5       ' dark energy burst (teleport smoke, demon palm)

Private Const MAX_FX As Long = 8
Private Const FX_LIFE As Long = 4
Private Const SLAM_LIFE As Long = 8

Private mX(0 To MAX_FX - 1) As Long
Private mY(0 To MAX_FX - 1) As Long
Private mType(0 To MAX_FX - 1) As Long
Private mAge(0 To MAX_FX - 1) As Long

Public Sub Init()
    Dim k As Long
    For k = 0 To MAX_FX - 1
        mAge(k) = FX_LIFE
    Next k
End Sub

Public Sub Spawn(ByVal x As Long, ByVal y As Long, ByVal kind As Long)
    Dim k As Long, oldest As Long
    oldest = 0
    For k = 0 To MAX_FX - 1
        If mAge(k) >= Life(k) Then Exit For
        If mAge(k) > mAge(oldest) Then oldest = k
    Next k
    If k = MAX_FX Then k = oldest
    mX(k) = x: mY(k) = y: mType(k) = kind: mAge(k) = 0
End Sub

Public Sub Update()
    Dim k As Long
    For k = 0 To MAX_FX - 1
        If mAge(k) < Life(k) Then mAge(k) = mAge(k) + 1
    Next k
End Sub

Public Sub Draw()
    Dim k As Long
    For k = 0 To MAX_FX - 1
        If mAge(k) < Life(k) Then
            If mType(k) = FX_SLAM Then
                Slam mX(k), mY(k), mAge(k)
            ElseIf mType(k) = FX_SLASH Then
                SlashFx mX(k), mY(k), mAge(k), k
            ElseIf mType(k) = FX_DARK Then
                DarkFx mX(k), mY(k), mAge(k)
            Else
                Spark mX(k), mY(k), mAge(k), mType(k)
            End If
        End If
    Next k
End Sub

Private Function Life(ByVal k As Long) As Long
    If mType(k) = FX_SLAM Then Life = SLAM_LIFE Else Life = FX_LIFE
End Function

' Purple-black burst: dark core, violet ring, drifting motes.
Private Sub DarkFx(ByVal x As Long, ByVal y As Long, ByVal age As Long)
    Dim k As Long, a As Double, r As Long
    r = 3 + age * 3
    If age < 2 Then Renderer.FillRect x - 3, y - 3, 7, 7, RGB(30, 10, 40)
    For k = 0 To 15
        a = k * 3.14159 / 8 + age * 0.3
        Renderer.PutPixel x + Int(Cos(a) * r), y + Int(Sin(a) * r), IIf(age < 2, RGB(200, 120, 255), RGB(120, 60, 170))
    Next k
    For k = 0 To 3
        Renderer.PutPixel x - 5 + k * 3, y - 4 - age * 2 - (k Mod 2) * 2, RGB(150, 90, 200)
    Next k
End Sub

' Diagonal cut streak that flashes and thins out; direction varies per slot.
Private Sub SlashFx(ByVal x As Long, ByVal y As Long, ByVal age As Long, ByVal k As Long)
    Dim d As Long, sx As Long, n As Long, c As Long
    If k Mod 2 = 0 Then sx = 1 Else sx = -1
    n = 9 - age * 2
    If age = 0 Then c = RGB(255, 255, 255) Else c = RGB(190, 230, 255)
    For d = -n To n
        Renderer.PutPixel x + d, y + sx * (d \ 2), c
        If age < 2 Then Renderer.PutPixel x + d, y + sx * (d \ 2) + 1, RGB(140, 200, 255)
    Next d
End Sub

' Ground impact: flattened shock ring that widens, plus dust puffs and a white core.
Private Sub Slam(ByVal x As Long, ByVal y As Long, ByVal age As Long)
    Dim r As Long, k As Long, a As Double, c As Long, dx As Long
    r = 4 + age * 4
    If age < 3 Then c = RGB(255, 255, 255) Else c = RGB(255, 210, 120)
    For k = 0 To 31
        a = k * 3.14159 / 16
        Renderer.PutPixel x + Int(Cos(a) * r), y + Int(Sin(a) * r * 0.35), c
    Next k
    If age < 2 Then Renderer.FillRect x - 6, y - 10, 13, 10, RGB(255, 255, 255)
    For dx = -2 To 2
        If dx <> 0 Then
            Renderer.FillRect x + dx * (6 + age * 3), y - 2 - age \ 2, 3, 2, RGB(170, 150, 120)
            Renderer.PutPixel x + dx * (5 + age * 3), y - 4 - age, RGB(200, 185, 150)
        End If
    Next dx
End Sub

Private Sub Spark(ByVal x As Long, ByVal y As Long, ByVal age As Long, ByVal kind As Long)
    Dim cCore As Long, cMid As Long, cOut As Long, r As Long, k As Long
    If kind = FX_BLOCK Then
        cCore = RGB(255, 255, 255): cMid = RGB(150, 215, 255): cOut = RGB(60, 120, 255)
    Else
        cCore = RGB(255, 255, 255): cMid = RGB(255, 224, 102): cOut = RGB(255, 130, 50)
    End If
    If kind = FX_HEAVY Then r = 2 Else r = 0
    Select Case age
    Case 0
        Ray x, y, 0, 2 + r, cCore
        Renderer.PutPixel x, y, cCore
    Case 1
        Ray x, y, 0, 3 + r, cMid
        Diag x, y, 2 + r, cMid
        Renderer.FillRect x - 1, y - 1, 3, 3, cCore
    Case 2
        Ray x, y, 3 + r, 5 + r, cOut
        Diag x, y, 3 + r, cOut
        Renderer.PutPixel x, y, cMid
    Case Else
        For k = -1 To 1 Step 2
            Renderer.PutPixel x + k * (6 + r), y, cOut
            Renderer.PutPixel x, y + k * (6 + r), cOut
        Next k
    End Select
End Sub

' + shaped rays from radius r0 to r1
Private Sub Ray(ByVal x As Long, ByVal y As Long, ByVal r0 As Long, ByVal r1 As Long, ByVal c As Long)
    Dim d As Long
    For d = r0 To r1
        Renderer.PutPixel x + d, y, c
        Renderer.PutPixel x - d, y, c
        Renderer.PutPixel x, y + d, c
        Renderer.PutPixel x, y - d, c
    Next d
End Sub

' x shaped dots at radius r
Private Sub Diag(ByVal x As Long, ByVal y As Long, ByVal r As Long, ByVal c As Long)
    Renderer.PutPixel x + r, y + r, c
    Renderer.PutPixel x - r, y + r, c
    Renderer.PutPixel x + r, y - r, c
    Renderer.PutPixel x - r, y - r, c
End Sub
