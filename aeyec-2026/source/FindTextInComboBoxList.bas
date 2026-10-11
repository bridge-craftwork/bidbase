Attribute VB_Name = "Module2"
''''''''''''''''''''''''''''''
' Find text in ComboBox list '
''''''''''''''''''''''''''''''

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

    


