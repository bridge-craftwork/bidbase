VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function TextPixelWidth(txt As String) As Long
    Dim c As String ' character being measured '
    Dim i As Integer
    Dim t As Long ' ?? '
    Dim tt As Long ' total of whatever t is '
    Dim p As Long ' width of a character in pixels '
    Dim tp As Long ' total pixels width '

    ' This sub calculates the screen width of text for MyMsgBox. '
    ' Seems excessive, but what the heck. Neatness counts. '

    txt = Trim$(txt)
    
    For i = 1 To Len(txt)
        c = Mid$(txt, i, 1)
        Select Case c
        Case "f", "i", "í", "j", "l", "t", ".", ",", ";", ":", "|", "[", "]", "I", "'"
            t = t + 0.428
            p = p + 45.7 ' pixels '
        Case "r", "-", "/", "\", "(", ")", " "
            t = t + 0.571
            p = p + 60.71
        Case "{", "}"
            t = t + 0.712
            p = p + 60.71
        Case Chr$(34)
            t = t + 0.716
            p = p + 45.455
        Case "x", "z"
            t = t + 0.857
            p = p + 90.426
        Case "c", "h", "k", "n", "s", "u", "v", "y", "J", "L", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "="
            t = t + 1
            p = p + 106.25
        Case "a", "b", "d", "e", "g", "o", "p", "q", "F", "K", "X", "Z"
            t = t + 1.147
            p = p + 121.434
        Case "A", "B", "C", "E", "P", "S", "T", "V", "Y", "<", ">", "w", "#"
            t = t + 1.289
            p = p + 137.097
        Case "D", "G", "H", "N", "O", "Q", "R", "U"
            t = t + 1.431
            p = p + 151.786
        Case "M", "m"
            t = t + 1.576
            p = p + 166.67
        Case "W"
            t = t + 1.857
            p = p + 197.674
        Case Else
            t = t + 1
            p = p + 100
        End Select
        tt = tt + t
        tp = tp + p
    Next
    Stop
    TextWidth = tp ' or tt? '
End Function

