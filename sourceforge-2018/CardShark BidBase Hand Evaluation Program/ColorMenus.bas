Attribute VB_Name = "Module1"
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
        
Declare Function CreateHatchBrush Lib "gdi32" _
            (ByVal nIndex As Long, _
             ByVal crColor As Long) _
        As Long

Declare Function OleTranslateColor Lib "olepro32.dll" _
            (ByVal OLE_COLOR As Long, _
             ByVal HPALETTE As Long, _
             pccolorref As Long) _
        As Long

