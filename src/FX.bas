Attribute VB_Name = "FX"
Option Explicit

' Short-lived visual effects (hit / guard sparks), drawn procedurally.

Public Const FX_HIT As Long = 0
Public Const FX_HEAVY As Long = 1
Public Const FX_BLOCK As Long = 2

Private Const MAX_FX As Long = 8
Private Const FX_LIFE As Long = 4

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
        If mAge(k) >= FX_LIFE Then Exit For
        If mAge(k) > mAge(oldest) Then oldest = k
    Next k
    If k = MAX_FX Then k = oldest
    mX(k) = x: mY(k) = y: mType(k) = kind: mAge(k) = 0
End Sub

Public Sub Update()
    Dim k As Long
    For k = 0 To MAX_FX - 1
        If mAge(k) < FX_LIFE Then mAge(k) = mAge(k) + 1
    Next k
End Sub

Public Sub Draw()
    Dim k As Long
    For k = 0 To MAX_FX - 1
        If mAge(k) < FX_LIFE Then Spark mX(k), mY(k), mAge(k), mType(k)
    Next k
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
