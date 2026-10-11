Attribute VB_Name = "MultiMonitor"
DefInt A-Z

Public Const SM_XVIRTUALSCREEN = 76    'virtual desktop left
Public Const SM_YVIRTUALSCREEN = 77    'virtual top
Public Const SM_CXVIRTUALSCREEN = 78   'virtual width
Public Const SM_CYVIRTUALSCREEN = 79   'virtual height
Public Const SM_CMONITORS = 80         'number of monitors
Public Const SM_SAMEDISPLAYFORMAT = 81

Public Const MONITOR_DEFAULTTONULL As Long = &H0    'If the monitor is not found, return 0
Public Const MONITOR_DEFAULTTOPRIMARY As Long = &H1 'If the monitor is not found, return the primary monitor
Public Const MONITOR_DEFAULTTONEAREST As Long = &H2 'If the monitor is not found, return the nearest monitor

Global hMonitor!, hMonitor2!, hMonitor3!
Global NumMonitors
Global VirtualScrLeft!, VirtualScrTop!
Global VirtualScrHeight!, VirtualScrWidth!
Global Screen1Left!, Screen1Top!, Screen1Height!, Screen1Width!
Global Screen2Left!, Screen2Top!, Screen2Height!, Screen2Width!
Global Screen3Left!, Screen3Top!, Screen3Height!, Screen3Width!

Public Declare Function MonitorFromWindow Lib "user32" _
  (ByVal hwnd As Long, ByVal dwFlags As Long) As Long

Public Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   source As Any, _
   ByVal Length As Long)
   
Public Declare Function EnumDisplayMonitors Lib "user32" _
  (ByVal hdc As Long, _
   ByVal lprcClip As Long, _
   ByVal lpfnEnum As Long, _
   dwData As Any) As Long

Public Declare Function GetSystemMetrics Lib "user32" _
   (ByVal nIndex As Long) As Long
   
'Menu Colors:
''''''''''''
Public Const MIM_BACKGROUND As Long = &H2
Public Const MIM_APPLYTOSUBMENUS As Long = &H80000000

Type MENUINFO
   cbSize As Long
   fMask As Long
   dwStyle As Long
   cyMax As Long
   hbrBack As Long
   dwContextHelpID As Long
   dwMenuData As Long
End Type

Declare Function DrawMenuBar Lib "user32" _
            (ByVal hwnd As Long) _
        As Long

Declare Function GetMenu Lib "user32" _
            (ByVal hwnd As Long) _
        As Long

Declare Function GetSystemMenu Lib "user32" _
            (ByVal hwnd As Long, _
             ByVal bRevert As Long) _
        As Long
   
Declare Function GetSubMenu Lib "user32.dll" _
            (ByVal hMenu As Long, _
             ByVal nPos As Long) _
        As Long

Declare Function SetMenuInfo Lib "user32" _
            (ByVal hMenu As Long, _
             mi As MENUINFO) _
        As Long

Declare Function CreateSolidBrush Lib "gdi32" _
            (ByVal crColor As Long) _
        As Long

Declare Function OleTranslateColor Lib "olepro32.dll" _
            (ByVal OLE_COLOR As Long, _
             ByVal HPALETTE As Long, _
             pccolorref As Long) _
        As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Find text in ComboBox list:
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
            (ByVal hwnd As Long, _
             ByVal wMsg As Long, _
             ByVal wParam As Long, _
             lParam As Any) _
        As Long

Const CB_FINDSTRING = &H14C
Const CB_FINDSTRINGEXACT = &H158
Const LB_FINDSTRING = &H18F
Const LB_FINDSTRINGEXACT = &H1A2

    
Function GetDesktopMaximumHeight() As Long
  'Return the maximum height of all monitors on the desktop.
  'If only one monitor, return screen height.
   If IsMultiMonitorSystem() Then
      GetDesktopMaximumHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
   Else
      GetDesktopMaximumHeight = Screen.Height \ Screen.TwipsPerPixelY
   End If
End Function

Private Function IsMultiMonitorSystem() As Boolean
  'Returns True if a multi-monitor system
   IsMultiMonitorSystem = GetSystemMetrics(SM_CMONITORS) > 1
End Function

Public Function MonitorEnumProc(ByVal hMonitor As Long, _
                                ByVal hdcMonitor As Long, _
                                ByVal rcMonitor As Long, _
                                ByVal dwData As Long) As Long
  Dim rc As RECT
  'copy the rectangle data from
  'rcMonitor into a RECT type
  CopyMemory rc, ByVal rcMonitor, Len(rc)
  If Screen1Width! = 0 Then
    Screen1Left! = BBEditor.ScaleX(rc.Left, vbPixels, vbTwips)
    Screen1Top! = BBEditor.ScaleX(rc.Top, vbPixels, vbTwips)
    Screen1Width! = BBEditor.ScaleX(rc.Right, vbPixels, vbTwips) - Screen1Left!
    Screen1Height! = BBEditor.ScaleX(rc.Bottom, vbPixels, vbTwips) - Screen1Top!
  ElseIf Screen2Width! < 1 And _
         ((hMonitor = hMonitor2! And _
           hMonitor2! < 9999999) Or _
          (Abs(hMonitor - hMonitor2!) < 100 _
           And hMonitor2! >= 9999999)) _
  Then
    Screen2Left! = BBEditor.ScaleX(rc.Left, vbPixels, vbTwips)
    Screen2Top! = BBEditor.ScaleX(rc.Top, vbPixels, vbTwips)
    Screen2Width! = BBEditor.ScaleX(rc.Right, vbPixels, vbTwips) - Screen2Left!
    Screen2Height! = BBEditor.ScaleX(rc.Bottom, vbPixels, vbTwips) - Screen2Top!
  End If
  
  'return 1 to continue processing
  'monitor information, or return
  '0 to stop. Returning 1 obtains
  'info for all physical and invisible
  'pseudo-monitors associated with
  'mirroring drivers
   MonitorEnumProc = 1
End Function

Public Function GetMonitorByWindow(ByVal hwnd As Long) As Long
  'Returns a handle to the display monitor
  'that has the largest area of intersection
  'with the bounding rectangle of a specified
  'window.
  '
  'If the window specified by hwnd intersects
  'one or more display monitor rectangles,
  'the return value is an HMONITOR handle
  'to the display monitor that has the
  'largest area of intersection with the
  'window.
  '
  'If the window is currently minimized,
  'the window occupied before minimization
  'is returned.
  '
  'If the window does not intersect a display
  'monitor, the return value depends on the
  'value of dwFlags. Available values are:
  'MONITOR_DEFAULTTONEAREST: Returns a handle to the
  '                          display monitor that is
  '                          nearest to the window.
  'MONITOR_DEFAULTTONULL:    Returns NULL
  'MONITOR_DEFAULTTOPRIMARY: Returns a handle to the
  '                          primary display monitor.
  
   GetMonitorByWindow = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
End Function


Public Sub SetupMultipleMonitors()
  NumMonitors = GetSystemMetrics(SM_CMONITORS)
  VirtualScrLeft! = BBEditor.ScaleX(GetSystemMetrics(SM_XVIRTUALSCREEN), vbPixels, vbTwips)
  VirtualScrTop! = BBEditor.ScaleX(GetSystemMetrics(SM_YVIRTUALSCREEN), vbPixels, vbTwips)
  VirtualScrHeight! = BBEditor.ScaleX(GetSystemMetrics(SM_CYVIRTUALSCREEN), vbPixels, vbTwips)
  VirtualScrWidth! = BBEditor.ScaleX(GetSystemMetrics(SM_CXVIRTUALSCREEN), vbPixels, vbTwips)
End Sub

