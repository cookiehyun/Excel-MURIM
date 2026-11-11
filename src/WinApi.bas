Attribute VB_Name = "WinApi"
Option Explicit

' Windows API wrappers that are not input related.
' (The test harness swaps this module for a stub.)

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Sub SleepMs(ByVal ms As Long)
    Sleep ms
End Sub
