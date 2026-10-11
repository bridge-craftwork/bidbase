Attribute VB_Name = "MultiMonitor"
Option Explicit

Global ItsMe As String

Public Const SM_XVIRTUALSCREEN = 76   'virtual desktop left
Public Const SM_YVIRTUALSCREEN = 77    'virtual top
Public Const SM_CXVIRTUALSCREEN = 78   'virtual width
Public Const SM_CYVIRTUALSCREEN = 79   'virtual height
Public Const SM_CMONITORS = 80         'number of monitors
Public Const SM_SAMEDISPLAYFORMAT = 81

Public Const MONITOR_DEFAULTTONULL As Long = &H0    'If the monitor is not found, return 0
Public Const MONITOR_DEFAULTTOPRIMARY As Long = &H1 'If the monitor is not found, return the primary monitor
Public Const MONITOR_DEFAULTTONEAREST As Long = &H2 'If the monitor is not found, return the nearest monitor

Global lng_hMonitor As Long
Global lng_hMonitor2 As Long
Global lng_hMonitor3 As Long
Global int_NumMonitors As Integer
Global lng_VirtualScrLeft As Long
Global lng_VirtualScrTop As Long
Global lng_VirtualScrHeight As Long
Global lng_VirtualScrWidth As Long
Global lng_Screen1Left As Long
Global lng_Screen1Top As Long
Global lng_Screen1Height As Long
Global lng_Screen1Width As Long
Global lng_Screen2Left As Long
Global lng_Screen2Top As Long
Global lng_Screen2Height As Long
Global lng_Screen2Width As Long
Global lng_Screen3Left As Long
Global lng_Screen3Top As Long
Global lng_Screen3Height As Long
Global lng_Screen3Width As Long

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
  If lng_Screen1Width = 0 Then
    lng_Screen1Left = BidBaseFileViewer.ScaleX(rc.Left, vbPixels, vbTwips)
    lng_Screen1Top = BidBaseFileViewer.ScaleX(rc.Top, vbPixels, vbTwips)
    lng_Screen1Width = BidBaseFileViewer.ScaleX(rc.Right, vbPixels, vbTwips) - lng_Screen1Left
    lng_Screen1Height = BidBaseFileViewer.ScaleX(rc.Bottom, vbPixels, vbTwips) - lng_Screen1Top
  End If
  
  If lng_Screen2Width < 1 And _
     ((hMonitor = lng_hMonitor2 And _
       lng_hMonitor2 < 9999999) Or _
      (Abs(hMonitor - lng_hMonitor2) < 100 And _
       lng_hMonitor2 >= 9999999)) _
  Then
    lng_Screen2Left = BidBaseFileViewer.ScaleX(rc.Left, vbPixels, vbTwips)
    lng_Screen2Top = BidBaseFileViewer.ScaleX(rc.Top, vbPixels, vbTwips)
    lng_Screen2Width = BidBaseFileViewer.ScaleX(rc.Right, vbPixels, vbTwips) - lng_Screen2Left
    lng_Screen2Height = BidBaseFileViewer.ScaleX(rc.Bottom, vbPixels, vbTwips) - lng_Screen2Top
  End If
  
  If lng_Screen3Width < 1 And _
     ((hMonitor = lng_hMonitor3 And _
       lng_hMonitor3 < 9999999) Or _
      (Abs(hMonitor - lng_hMonitor3) < 100 And _
       lng_hMonitor3 >= 9999999)) _
  Then
    lng_Screen3Left = BidBaseFileViewer.ScaleX(rc.Left, vbPixels, vbTwips)
    lng_Screen3Top = BidBaseFileViewer.ScaleX(rc.Top, vbPixels, vbTwips)
    lng_Screen3Width = BidBaseFileViewer.ScaleX(rc.Right, vbPixels, vbTwips) - lng_Screen2Left
    lng_Screen3Height = BidBaseFileViewer.ScaleX(rc.Bottom, vbPixels, vbTwips) - lng_Screen2Top
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

Sub SetupMultipleMonitors()
  int_NumMonitors = GetSystemMetrics(SM_CMONITORS)
  lng_VirtualScrLeft = BidBaseFileViewer.ScaleX(GetSystemMetrics(SM_XVIRTUALSCREEN), vbPixels, vbTwips)
  lng_VirtualScrTop = BidBaseFileViewer.ScaleX(GetSystemMetrics(SM_YVIRTUALSCREEN), vbPixels, vbTwips)
  lng_VirtualScrHeight = BidBaseFileViewer.ScaleX(GetSystemMetrics(SM_CYVIRTUALSCREEN), vbPixels, vbTwips)
  lng_VirtualScrWidth = BidBaseFileViewer.ScaleX(GetSystemMetrics(SM_CXVIRTUALSCREEN), vbPixels, vbTwips)
End Sub

