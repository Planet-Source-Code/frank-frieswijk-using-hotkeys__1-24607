Attribute VB_Name = "modHOTKEY"
Option Explicit

'********************************************************
'   DECLARATIONS NEEDED TO INTERCEPT WINDOW MESSAGES    *
'********************************************************

Public Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal OldwndProc As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = -4
Public OldwndProc As Long

'********************************************************
'       DECLARATIONS NEEDED TO CREATE THE HOTKEY        *
'********************************************************

Public Declare Function RegisterHotKey Lib "USER32" (ByVal hWnd As Long, ByVal HotKeyID As Long, ByVal fsModifiers As Long, ByVal vKey As Long) As Long
Public Declare Function UnregisterHotKey Lib "USER32" (ByVal hWnd As Long, ByVal HotKeyID As Long) As Long

Public Const WM_HOTKEY = &H312
Public Const WM_NCDESTROY = &H82

Public HotKeyID(12) As Long

Public Function WindowProc(ByVal hWnd As Long, ByVal WindowMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'This is where all the messages for this form as directed to
'We will need to check for the WindowMessage WM_HOTKEY
'to see if a hotkey is pressed and then we need
'To check the wParam to see which HotKey (1-12) has been pressed

Select Case WindowMsg
    Case WM_HOTKEY
        Select Case wParam
            'This is where you put the code you want to start
            'whenever someone has pressed a hotkey
            Case HotKeyID(1)
                MsgBox "Hotkey F1 has been pressed."
            Case HotKeyID(2)
                MsgBox "Hotkey F2 has been pressed."
            Case HotKeyID(3)
                MsgBox "Hotkey F3 has been pressed."
            Case HotKeyID(4)
                MsgBox "Hotkey F4 has been pressed."
            Case HotKeyID(5)
                MsgBox "Hotkey F5 has been pressed."
            Case HotKeyID(6)
                MsgBox "Hotkey F6 has been pressed."
            Case HotKeyID(7)
                MsgBox "Hotkey F7 has been pressed."
            Case HotKeyID(8)
                MsgBox "Hotkey F8 has been pressed."
            Case HotKeyID(9)
                MsgBox "Hotkey F9 has been pressed."
            Case HotKeyID(10)
                MsgBox "Hotkey F10 has been pressed."
            Case HotKeyID(11)
                MsgBox "Hotkey F11 has been pressed."
            Case HotKeyID(12)
                MsgBox "Hotkey F12 has been pressed."
        End Select
End Select

'No matter what happens we *always* end with the normal
'window procedure to finish/handle the message by
'calling the CallWindowProc.
WindowProc = CallWindowProc(OldwndProc, hWnd, WindowMsg, wParam, lParam)

End Function


