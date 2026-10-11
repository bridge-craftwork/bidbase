VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form TableSpecs 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table Specifications"
   ClientHeight    =   6720
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8070
   Icon            =   "BidBaseFileViewer_Dialog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   5475
      Left            =   2025
      TabIndex        =   21
      Top             =   270
      Width           =   5805
      ExtentX         =   10239
      ExtentY         =   9657
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame Frame2 
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   165
      TabIndex        =   20
      Top             =   3420
      Width           =   1740
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1035
         Left            =   150
         TabIndex        =   27
         Top             =   315
         Width           =   1515
         Begin VB.CheckBox ck_Background 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   32
            Top             =   690
            Width           =   165
         End
         Begin VB.CheckBox ck_Background 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   31
            Top             =   315
            Value           =   1  'Checked
            Width           =   165
         End
         Begin VB.CommandButton cmdBackgroundPic 
            Caption         =   "&Picture"
            Height          =   360
            Left            =   405
            TabIndex        =   29
            Top             =   645
            Width           =   975
         End
         Begin VB.CommandButton cmbTableColors 
            Caption         =   "&Color"
            Height          =   315
            Index           =   0
            Left            =   420
            TabIndex        =   28
            Top             =   255
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Background"
            Height          =   210
            Index           =   1
            Left            =   30
            TabIndex        =   30
            Top             =   15
            Width           =   1185
         End
      End
      Begin VB.OptionButton optFrameColor 
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   24
         Top             =   2415
         Value           =   -1  'True
         Width           =   180
      End
      Begin VB.OptionButton optFrameColor 
         Caption         =   "Option1"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   22
         Top             =   1725
         Width           =   180
      End
      Begin VB.CommandButton cmbTableColors 
         Caption         =   "&Shadow"
         Height          =   315
         Index           =   3
         Left            =   615
         TabIndex        =   10
         Top             =   2610
         Width           =   855
      End
      Begin VB.CommandButton cmbTableColors 
         Caption         =   "&Highlight"
         Height          =   315
         Index           =   2
         Left            =   630
         TabIndex        =   9
         Top             =   2190
         Width           =   855
      End
      Begin VB.CommandButton cmbTableColors 
         Caption         =   "Solid"
         Height          =   315
         Index           =   1
         Left            =   495
         TabIndex        =   8
         Top             =   1665
         Width           =   1020
      End
      Begin VB.Label Label2 
         Caption         =   "Frame Color:"
         Height          =   210
         Index           =   2
         Left            =   150
         TabIndex        =   26
         Top             =   1425
         Width           =   1185
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   525
         X2              =   525
         Y1              =   2160
         Y2              =   2955
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   630
         X2              =   525
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   630
         X2              =   525
         Y1              =   2145
         Y2              =   2145
      End
      Begin VB.Label Label2 
         Caption         =   "OR"
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   1935
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sizing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   180
      TabIndex        =   13
      Top             =   180
      Width           =   1740
      Begin VB.OptionButton optTableWidth 
         Alignment       =   1  'Right Justify
         Caption         =   " # of pixels"
         Height          =   210
         Index           =   1
         Left            =   225
         TabIndex        =   7
         Top             =   2655
         Width           =   1215
      End
      Begin VB.OptionButton optTableWidth 
         Alignment       =   1  'Right Justify
         Caption         =   "% of window"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   6
         Top             =   2430
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtTableWidth 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1140
         TabIndex        =   5
         Top             =   2085
         Width           =   390
      End
      Begin VB.TextBox txtCellPadding 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "5"
         Top             =   1710
         Width           =   270
      End
      Begin VB.TextBox txtCellSpacing 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "5"
         Top             =   1395
         Width           =   270
      End
      Begin VB.TextBox txtBorderSize 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   "5"
         Top             =   1095
         Width           =   270
      End
      Begin VB.TextBox txtNumColumns 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "3"
         Top             =   780
         Width           =   270
      End
      Begin VB.TextBox txtNumRows 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Text            =   "6"
         Top             =   480
         Width           =   270
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   1590
         X2              =   1485
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   1590
         X2              =   1485
         Y1              =   2835
         Y2              =   2835
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1590
         X2              =   1590
         Y1              =   2055
         Y2              =   2850
      End
      Begin VB.Label Label1 
         Caption         =   "Table Width:"
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   19
         Top             =   2130
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Cell Margins:"
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   18
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cell Frame:"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   17
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Table frame:"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   16
         Top             =   1125
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "# Of Cols.:"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   825
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "# Of Rows:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   510
         Width           =   1020
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6570
      TabIndex        =   12
      Top             =   6030
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   6030
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Cell properties can be set with other HTML menu options."
      Height          =   435
      Left            =   2445
      TabIndex        =   25
      Top             =   6015
      Width           =   2445
   End
End
Attribute VB_Name = "TableSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private str_BackgroundPic As String
Private str_TableColors(3) As String


Private Sub CancelButton_Click()
    Unload TableSpecs
End Sub


Private Sub ck_Background_Click(index As Integer)
    SaveChange
End Sub


Private Sub cmbTableColors_Click(index As Integer)
    ' Get table colors. '
    
    Dim colr As String
    Dim colrRev As String
    Dim hx As String
    
    On Error GoTo TblColorError
    With BidBaseFileViewer.CmDia
        .CancelError = True
        .Flags = cdlCCRGBInit Or cdlCCFullOpen
        colr = Mid$(str_TableColors(index), 2)
        ' HTML colors are RGB, '
        ' and CmDia's are BGR, '
        ' so reverse them:     '
        colrRev = "&H" _
                & Right$(colr, 2) _
                & Mid$(colr, 3, 2) _
                & Left$(colr, 2)
        .Color = colrRev
        .ShowColor
        hx = Hex(.Color)
        str_TableColors(index) = "#" _
                           & Right$(hx, 2) _
                           & Mid$(hx, 3, 2) _
                           & Left$(hx, 2)
    End With
    SaveChange
ExitTblColor:
    On Error GoTo 0
Exit Sub

TblColorError:
    'User pressed Cancel button.
    Resume ExitTblColor
End Sub


Private Sub cmdBackgroundPic_Click()
    ' Get table background picture.'
    
    Dim fi As String
    
    With BidBaseFileViewer.CmDia
        .CancelError = True
        On Error GoTo PicErr
        .Filter = "(*.jpg)(.bmp)|*.jpg *.bmp"
        .FilterIndex = 0
        .ShowOpen
        fi = .FileName
    End With
    str_BackgroundPic = fi
    SaveChange
ExitPic:
    On Error GoTo 0
Exit Sub
      
PicErr:
    Resume ExitPic
End Sub


Private Sub Form_Load()
    Dim a As String
    Dim i As Integer
    
    ' Create a work file: '
    On Local Error GoTo TblSpecsNotFound
    FileCopy App.Path & "\TableSpecs.htm ", App.Path & "\Temp.htm"
    On Local Error GoTo 0
    ' Load the work file: '
    WebBrowser1.Navigate App.Path & "\Temp.htm"
  
    ' Need to load the color numbers because   '
    ' each time the user changes a table spec, '
    ' we rewrite the Temp.htm file so that the '
    ' changes show up. Or use the last values  '
    ' if this form was loaded before.          '
    If str_TableColors(0) <> "" Then Exit Sub
  
    Open App.Path & "\Temp.htm" For Input As #1
        For i = 1 To 4
            Line Input #1, a
        Next
        For i = 0 To 3
            Line Input #1, a
            Input #1, str_TableColors(i)
        Next
    Close
Exit Sub

TblSpecsNotFound:
    MsgBox App.Path & "\notes\TableSpecs.htm not found."
    Exit Sub

End Sub


Private Sub OKButton_Click()
    Dim a As String
    Dim b As String
    Dim i As Integer
    Dim fiNum As Integer
    
    fiNum = FreeFile
    Open App.Path & "\notes\Temp.htm" For Input As #fiNum
        Do
            Line Input #fiNum, a
            b = b + a + vbCrLf
        Loop While Not EOF(fiNum)
    Close fiNum
    
    i = InStr(b, "<T")
    If i > 0 Then b = Mid$(b, i)
    
    i = InStr(100, b, "</body>")
    If i > 0 Then b = Left$(b, i - 1)
   
    BidBaseFileViewer.Text2(1).SelText = b
   
    Unload TableSpecs
End Sub


Private Sub optFrameColor_Click(index As Integer)
    SaveChange
End Sub


Private Sub optTableWidth_Click(index As Integer)
    SaveChange
End Sub


Private Sub txtBorderSize_Change()
    SaveChange
End Sub


Private Sub txtCellPadding_Change()
    SaveChange
End Sub


Private Sub txtCellSpacing_Change()
    SaveChange
End Sub


Private Sub TxtNumColumns_Change()
    SaveChange
End Sub


Private Sub TxtNumRows_Change()
    SaveChange
End Sub


Private Sub SaveChange()
    If Val(txtNumRows) = 0 Or _
       Val(txtNumColumns) = 0 _
    Then
        Exit Sub
    End If
  
    Dim fiNum As Integer
    Dim a As String
    Dim r As Integer
    Dim c As Integer
    
    fiNum = FreeFile
    Open App.Path & "\notes\Temp.htm" For Output As #fiNum
    ' It doesn't really matter how  '
    ' the Temp file is formatted    '
    ' because we don't read it back '
    ' in. Only the browser reads it.'
    Print #fiNum, "<html><head><title>Table</title></head>"
    Print #fiNum, "<body><CENTER>"
    Print #fiNum, "<Table";
  
    If Val(txtBorderSize) > 0 Then 'Table frame
        a = " Border=""" & Trim$(txtBorderSize) & """"
        Print #1, a
    End If
  
    If Val(txtCellPadding) > 0 Then 'Cell frame
        a = " CellPadding=""" & Trim$(txtCellPadding) & """"
        Print #1, a
    End If
  
    If Val(txtCellSpacing) > 0 Then 'Cell margins
        a = " CellSpacing=""" & Trim$(txtCellSpacing) & """"
        Print #1, a
    End If
  
    If Val(txtTableWidth) > 0 Then 'Table width
        a = " Width=" & Chr$(34) & Trim$(txtTableWidth)
        If optTableWidth(0) Then
            a = a & "%" & Chr$(34)
        Else
            a = a & Chr$(34)
        End If
        Print #1, a
    End If
  
    If str_TableColors(0) <> "" And _
       ck_Background(0).Value <> 0 _
    Then
        a = " bgcolor=" & Chr$(34) & str_TableColors(0) & Chr$(34)
        Print #1, a
    End If
  
    If str_TableColors(1) <> "" And _
       optFrameColor(0) _
    Then
        a = " bordercolor=" & Chr$(34) _
           & str_TableColors(1) & Chr$(34)
        Print #1, a
  
    ElseIf optFrameColor(1) Then
        If str_TableColors(2) <> "" Then
            a = " bordercolorlight=" & Chr$(34) _
                & str_TableColors(2) & Chr$(34)
            Print #1, a
        End If
    
        If str_TableColors(3) <> "" Then
            a = " bordercolordark=" & Chr$(34) _
                & str_TableColors(3) & Chr$(34)
            Print #1, a
        End If
    End If
  
    If str_BackgroundPic <> "" And _
       ck_Background(1).Value <> 0 _
    Then
        Print #1, " background=" & Chr$(34) _
                  & str_BackgroundPic & Chr$(34)
    End If
  
    'Print #1, ">"
    
    'lay out columns and rows:
    For r = 0 To Val(txtNumRows)
        Print #1, "<TR>"
        For c = 0 To Val(txtNumColumns)
            Print #1, "<TD>&nbsp;</TD>"
        Next
        Print #1, "</TR>"
    Next
    Print #1, "</TABLE>"

    Print #1, "</body>"
    Print #1, "</html>"
    Close
    WebBrowser1.Refresh
End Sub


Private Sub txtTableWidth_Change()
    If Val(txtTableWidth) > 9 Then SaveChange
End Sub
