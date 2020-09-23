Attribute VB_Name = "modTest"
Option Explicit

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_WNDPROC As Long = -4
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private hOldWndProc As Long

Public Function SubclassForm(hwnd As Long) As Long
' simple subclass on/off routine

    If hOldWndProc Then
        SetWindowLong hwnd, GWL_WNDPROC, hOldWndProc
        hOldWndProc = 0
    Else
        hOldWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf NewWndProc)
        SubclassForm = hOldWndProc
    End If
    
End Function

Private Function NewWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' sample routine only; no real purpose

    NewWndProc = CallWindowProc(hOldWndProc, hwnd, uMsg, wParam, lParam)

End Function

