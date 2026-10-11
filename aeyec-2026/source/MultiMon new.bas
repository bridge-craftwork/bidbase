Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hdc As Long, _
                                                           lprcClip As Any, _
                                                           ByVal lpfnEnum As Long, _
                                                           dwData As Any) _
                                                           As Long
Private Declare Function MonitorFromRect Lib "user32" (ByRef lprc As RECT, _
                                                       ByVal dwFlags As Long) _
                                                       As Long
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, _
                                                                              ByRef lpmi As MONITORINFO) _
                                                                              As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) _
                                                     As Long
Private Declare Function UnionRect Lib "user32" (lprcDst As RECT, _
                                                 lprcSrc1 As RECT, _
                                                 lprcSrc2 As RECT) _
                                                 As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) _
                                                  As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long, _
                                                  ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, _
                                                  ByVal bRepaint As Long) _
                                                  As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Const MONITOR_DEFAULTTONEAREST = &H2
Dim rcMonitors() As RECT 'coordinate array for all monitors
Dim rcVS         As RECT 'coordinates for Virtual Screen


Function EnumMonitors(F As Form) As Long
    Dim N As Long
    
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, N
    
    With F
        .Move .Left, .Top, _
        (rcVS.Right - rcVS.Left) * 2 + .Width - .ScaleWidth, _
        (rcVS.Bottom - rcVS.Top) * 2 + .Height - .ScaleHeight
        
        F.Scale (rcVS.Left, rcVS.Top)-(rcVS.Right, rcVS.Bottom)
        
        .Caption = N & " Monitor" & IIf(N > 1, "s", vbNullString)
        .lblMonitors(0).Appearance = 0 'Flat
        .lblMonitors(0).BorderStyle = 1 'FixedSingle
    End With
    
    For N = 0 To N - 1
        If N Then
            Load F.lblMonitors(N)
            F.lblMonitors(N).Visible = True
        End If
        
        With rcMonitors(N)
            F.lblMonitors(N).Move .Left, .Top, .Right - .Left, .Bottom - .Top
            F.lblMonitors(N).Caption = "Monitor " & N + 1 & vbLf & _
                .Right - .Left & " x " & .Bottom - .Top & vbLf & _
                "(" & .Left & ", " & .Top & ")-(" & .Right & ", " & .Bottom & ")"
        End With
    Next
End Function


Private Function MonitorEnumProc(ByVal hMonitor As Long, _
                                 ByVal hdcMonitor As Long, _
                                 lprcMonitor As RECT, _
                                 dwData As Long) _
                                As Long
    ReDim Preserve rcMonitors(dwData)
    
    rcMonitors(dwData) = lprcMonitor
    UnionRect rcVS, rcVS, lprcMonitor 'merge all monitors together to get the virtual screen coordinates
    dwData = dwData + 1 'increase monitor count
    MonitorEnumProc = 1 'continue
End Function


Sub SavePosition(hwnd As Long)
    Dim rc As RECT
    
    GetWindowRect hwnd, rc 'save position in pixel units
    SaveSetting "Multi Monitor Demo", "Position", "Left", rc.Left
    SaveSetting "Multi Monitor Demo", "Position", "Top", rc.Top
End Sub


Sub LoadPosition(hwnd As Long)
    Dim rc As RECT, _
        Left As Long, _
        Top As Long, _
        hMonitor As Long, _
        mi As MONITORINFO
    
    GetWindowRect hwnd, rc ' Obtain the window rectangle '
    ' Move the window rectangle to position saved previously "
    Left = GetSetting("Multi Monitor Demo", "Position", "Left", rc.Left)
    Top = GetSetting("Multi Monitor Demo", "Position", "Top", rc.Left)
    
    OffsetRect rc, Left - rc.Left, Top - rc.Top
    
    ' Find the monitor closest to window rectangle '
    hMonitor = MonitorFromRect(rc, MONITOR_DEFAULTTONEAREST)
    
    ' Get info about monitor coordinates and working area '
    mi.cbSize = Len(mi)
    GetMonitorInfo hMonitor, mi
    
    ' Adjust the window rectangle so it fits '
    ' inside the work area of the monitor    '
    If rc.Left < mi.rcWork.Left Then OffsetRect rc, mi.rcWork.Left - rc.Left, 0
    If rc.Right > mi.rcWork.Right Then OffsetRect rc, mi.rcWork.Right - rc.Right, 0
    If rc.Top < mi.rcWork.Top Then OffsetRect rc, 0, mi.rcWork.Top - rc.Top
    If rc.Bottom > mi.rcWork.Bottom Then OffsetRect rc, 0, mi.rcWork.Bottom - rc.Bottom
    
    ' Move the window to new calculated position '
    MoveWindow hwnd, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, 0
End Sub

