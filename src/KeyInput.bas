Attribute VB_Name = "KeyInput"
Option Explicit

' Real-time keyboard input via GetAsyncKeyState.
' Poll() once per logic tick, then query IsDown / IsPressed.

#If VBA7 Then
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

' Logical buttons
Public Const BTN_LEFT As Long = 0
Public Const BTN_RIGHT As Long = 1
Public Const BTN_UP As Long = 2
Public Const BTN_DOWN As Long = 3
Public Const BTN_LP As Long = 4   ' light punch
Public Const BTN_HP As Long = 5   ' heavy punch
Public Const BTN_LK As Long = 6   ' light kick
Public Const BTN_HK As Long = 7   ' heavy kick
Public Const BTN_QUIT As Long = 8
Public Const BTN_SWAP As Long = 9       ' debug: switch controlled fighter (Tab)
Public Const BTN_DBG_HIT As Long = 10   ' debug: hit the other fighter (H)
Public Const BTN_DBG_DOWN As Long = 11  ' debug: knock the other fighter down (J)
Public Const BTN_DBG_DUMMY As Long = 12 ' debug: toggle dummy attacks (D)
Public Const BTN_DBG_BOX As Long = 13   ' debug: toggle hitbox display (F1)
Public Const BTN_COUNT As Long = 14

Private mVK(0 To BTN_COUNT - 1) As Long
Private mCur(0 To BTN_COUNT - 1) As Boolean
Private mPrev(0 To BTN_COUNT - 1) As Boolean
Private mPid As Long
Private mHadFocus As Boolean
Private mFocusLost As Boolean

Public Sub Init()
    mVK(BTN_LEFT) = &H25    ' Left arrow
    mVK(BTN_UP) = &H26      ' Up arrow
    mVK(BTN_RIGHT) = &H27   ' Right arrow
    mVK(BTN_DOWN) = &H28    ' Down arrow
    mVK(BTN_LP) = &H41      ' A
    mVK(BTN_HP) = &H53      ' S
    mVK(BTN_LK) = &H5A      ' Z
    mVK(BTN_HK) = &H58      ' X
    mVK(BTN_QUIT) = &H1B    ' Esc
    mVK(BTN_SWAP) = &H9     ' Tab
    mVK(BTN_DBG_HIT) = &H48   ' H
    mVK(BTN_DBG_DOWN) = &H4A  ' J
    mVK(BTN_DBG_DUMMY) = &H44 ' D
    mVK(BTN_DBG_BOX) = &H70   ' F1
    mPid = GetCurrentProcessId()
    mHadFocus = True
    mFocusLost = False

    Dim i As Long
    For i = 0 To BTN_COUNT - 1
        GetAsyncKeyState mVK(i)   ' flush stale "pressed since last call" bits
        mCur(i) = False
        mPrev(i) = False
    Next i
End Sub

Public Sub Poll()
    Dim focused As Boolean, i As Long, s As Integer
    focused = AppHasFocus()
    mFocusLost = mHadFocus And Not focused       ' Excel just lost the foreground
    mHadFocus = focused
    For i = 0 To BTN_COUNT - 1
        mPrev(i) = mCur(i)
        If focused Then
            s = GetAsyncKeyState(mVK(i))
            ' high bit = held now, low bit = tapped since last call (catches taps shorter than a tick)
            mCur(i) = ((s And &H8000) <> 0) Or ((s And 1) <> 0)
        Else
            mCur(i) = False
        End If
    Next i
End Sub

' True on the one tick another window took the foreground (used for auto-pause).
Public Function FocusLost() As Boolean
    FocusLost = mFocusLost
End Function

Public Function IsDown(ByVal b As Long) As Boolean
    IsDown = mCur(b)
End Function

Public Function IsPressed(ByVal b As Long) As Boolean
    IsPressed = mCur(b) And Not mPrev(b)
End Function

' Ignore keys while another app is in front (GetAsyncKeyState is global).
Private Function AppHasFocus() As Boolean
    Dim pid As Long
    GetWindowThreadProcessId GetForegroundWindow(), pid
    AppHasFocus = (pid = mPid)
End Function
