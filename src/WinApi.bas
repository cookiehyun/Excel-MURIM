Attribute VB_Name = "WinApi"
Option Explicit

' Windows API wrappers that are not input related.
' (The test harness swaps this module for a stub.)
'   SleepMs        : Sleep, made accurate to ~1 ms while the game runs (timeBeginPeriod)
'   NowSec         : high resolution clock (QueryPerformanceCounter) in seconds
'   HighResTimer   : turn the 1 ms Windows timer resolution on / off

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
Private Declare PtrSafe Function timeBeginPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Private Declare PtrSafe Function timeEndPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
Private Declare Function timeBeginPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
#End If

Private mFreq As Currency
Private mHiRes As Boolean

Public Sub SleepMs(ByVal ms As Long)
    If ms > 0 Then Sleep ms
End Sub

' Seconds from an arbitrary start, microsecond resolution (VBA's Timer is much coarser).
Public Function NowSec() As Double
    Dim c As Currency
    If mFreq = 0 Then QueryPerformanceFrequency mFreq
    QueryPerformanceCounter c
    NowSec = c / mFreq
End Function

Public Sub HighResTimer(ByVal turnOn As Boolean)
    If turnOn And Not mHiRes Then
        timeBeginPeriod 1
        mHiRes = True
    ElseIf Not turnOn And mHiRes Then
        timeEndPeriod 1
        mHiRes = False
    End If
End Sub
