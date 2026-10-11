Attribute VB_Name = "MouseWheelGridScroller"
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' The mouse wheel can normally not be used to scroll a DataGrid. '
' This module allows it to do so. '
Global int_MouseOverGrid As Integer

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Public Const GWL_WNDPROC = (-4)
Public lpPrevWndProc As Long
Const WM_MOUSEWHEEL = &H20A
Const WHEEL_DELTA = 120
Dim Count As Integer

Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' This allows scrolling grids with a mouse scroll wheel. '
    If Msg = WM_MOUSEWHEEL Then
        Dim Delta As Long
        Static Travel As Long
        Delta = HiWord(wParam)
        Travel = Travel + Delta
        MouseWheel Travel \ WHEEL_DELTA, LoWord(lParam), HiWord(lParam)
        Travel = Travel Mod WHEEL_DELTA
    End If
    WndProc = CallWindowProc(lpPrevWndProc, hwnd, Msg, wParam, lParam)
End Function


Sub MouseWheel(Travel As Integer, x As Long, Y As Long)
    'scroll three rows per wheel step (windows default)
    BBEditor.Grid1(int_MouseOverGrid).Scroll 0, -Travel * 3
End Sub


Function HiWord(DWord As Long) As Integer
    CopyMemory HiWord, ByVal VarPtr(DWord) + 2, 2
End Function


Function LoWord(DWord As Long) As Integer
    CopyMemory LoWord, DWord, 2
End Function
