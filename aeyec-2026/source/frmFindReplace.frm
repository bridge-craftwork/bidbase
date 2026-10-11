VERSION 5.00
Begin VB.Form frmFindReplace 
   Caption         =   "Find/Replace Text"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   Icon            =   "frmFindReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkIgnoreHTML 
      Caption         =   "&Ignore text in < >"
      Height          =   195
      Left            =   150
      TabIndex        =   13
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2190
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1365
      TabIndex        =   0
      Top             =   135
      Width           =   2775
   End
   Begin VB.TextBox txtReplace 
      Height          =   285
      Left            =   1365
      TabIndex        =   2
      Top             =   495
      Width           =   2775
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "&Find Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4245
      TabIndex        =   1
      Top             =   135
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4245
      TabIndex        =   3
      Top             =   615
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace &All"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   4245
      TabIndex        =   4
      Top             =   1095
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4245
      TabIndex        =   5
      Top             =   1575
      Width           =   1215
   End
   Begin VB.CheckBox chkMatchWholeWord 
      Caption         =   "Match &whole words only"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   1050
      Width           =   2205
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Match &case"
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   1365
      Width           =   2190
   End
   Begin VB.Frame Frame1 
      Caption         =   "Start Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   2445
      TabIndex        =   10
      Top             =   975
      Width           =   1695
      Begin VB.OptionButton optDirection 
         Caption         =   "From &Top"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   315
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "From Cursor"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   645
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fi&nd what:"
      Height          =   195
      Left            =   165
      TabIndex        =   12
      Top             =   135
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Re&place with:"
      Height          =   195
      Left            =   165
      TabIndex        =   11
      Top             =   495
      Width           =   975
   End
End
Attribute VB_Name = "frmFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private str_BeforeReplace As String
Private lng_HighlightStart As Long
Private lng_HighlightLength As Long
Private int_NumFound As Integer
Private int_NumReplaced As Integer
Private bln_Replace As Boolean
Private lng_StartPtr As Long 'starting point of search
Private lng_TxtPtr As Long


Private Sub chkMatchCase_Click()
    int_NumFound = 0
    int_NumReplaced = 0
End Sub


Private Sub chkMatchWholeWord_Click()
    If optDirection(0) Then lng_TxtPtr = 0 'from top
    int_NumFound = 0
    int_NumReplaced = 0
End Sub


Private Sub cmdCancel_Click()
    'Discard changes:
    str_BeforeReplace = BidBaseFileViewer.Text2(1).Text
    Unload Me
End Sub


Private Sub cmdFindNext_Click()
    UnHighLight
    bln_Replace = False
    FindText
End Sub


Private Sub cmdReplace_Click(index As Integer)
    If index = 0 Then
        'Single replace:
        cmdReplace(0).Enabled = False
        If lng_TxtPtr = 0 Then lng_TxtPtr = 1
        With BidBaseFileViewer.Text2(1)
            .SelText = txtReplace
            .SelStart = lng_TxtPtr - 1
            .SelLength = Len(txtReplace)
        End With
        int_NumReplaced = int_NumReplaced + 1
        lng_TxtPtr = lng_TxtPtr + Len(txtReplace)
    
    Else  'Replace all:
        bln_Replace = True
        int_NumReplaced = 0
        FindText
    End If

End Sub


Private Sub Form_Activate()
    If str_BeforeReplace = "" Then
        lng_TxtPtr = BidBaseFileViewer.Text2(1).SelStart
        cmdReplace(0).Enabled = False
        str_BeforeReplace = BidBaseFileViewer.Text2(1).Text
        txtFind.SetFocus
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim ret As Long
    
    UnHighLight
    
    If str_BeforeReplace <> BidBaseFileViewer.Text2(1).Text Then
        ret = MsgBox("Save text replacements?", vbYesNoCancel)
        Select Case ret
            Case vbCancel
                Cancel = True
                Exit Sub
            Case vbNo
                BidBaseFileViewer.Text2(1).Text = _
                      str_BeforeReplace
            Case vbYes
                BidBaseFileViewer.mSave_Click 0
        End Select
    End If
    str_BeforeReplace = Empty
End Sub


Private Sub optDirection_Click(index As Integer)
    If optDirection(0) Then
        lng_TxtPtr = 0
    Else 'keep track of starting point:
        lng_StartPtr = lng_TxtPtr
    End If
End Sub


Private Sub txtFind_Change()
    cmdFindNext.Enabled = txtFind.Text <> ""
    cmdReplace(0).Enabled = cmdFindNext.Enabled
    cmdReplace(1).Enabled = cmdFindNext.Enabled
    If optDirection(0) Then lng_TxtPtr = 0
    int_NumReplaced = 0
    int_NumFound = 0
End Sub


Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub


Private Sub txtReplace_GotFocus()
    txtReplace.SelStart = 0
    txtReplace.SelLength = Len(txtReplace.Text)
End Sub


Private Sub FindText()
    Dim t As String
    Dim tf As String
    Dim notWholeWord As Boolean
    Dim c As String
    Dim alphaNum As String
    Dim s As Long
    Dim rBracket As Long
    Dim lBracket As Long
    Dim ret As Long
    Dim didRBracket As Long
    
With BidBaseFileViewer.Text2(1)

    UnHighLight
    
    If chkMatchCase.Value = 0 Then
        t = LCase$(.Text)
        tf = LCase$(txtFind)
    Else
        t = .Text
        tf = txtFind
    End If
    
    If optDirection(1).Value Then
        lng_TxtPtr = .SelStart + 1
    End If
    
    Do
        lng_TxtPtr = InStr(lng_TxtPtr + 1, t$, tf$)
        notWholeWord = False
        If lng_TxtPtr > 0 Then
            If chkMatchWholeWord.Value <> 0 Then
                alphaNum = "abcdefghijklmnopqrstuvwxyz0123456789"
                'Get letter before word:
                c = LCase$(Mid$(" " & t, lng_TxtPtr, 1))
                If InStr(alphaNum, c) > 0 Then
                    ' If letter before the found text is  '
                    ' alphanumeric (e.g.: 's' in front of '
                    ' 'tint') then the found text is not  '
                    ' a whole word.                       '
                    notWholeWord = True
                    
                Else ' Get letter after the word: '
                    c = LCase$(Mid$(t & " ", lng_TxtPtr + Len(tf), 1))
                    
                    If InStr(alphaNum, c) > 0 Then
                        notWholeWord = True
                    End If
                End If
            End If
           
            If chkIgnoreHTML.Value Then
                s = lng_TxtPtr - 100
                If s < 1 Then s = 1
                lBracket = InStr(s, t, "<")
              
                Do While lBracket > 0 And _
                         lBracket < lng_TxtPtr
                       
                    rBracket = InStr(lBracket + 1, t, ">")
                    If rBracket > lng_TxtPtr Then
                        If rBracket = didRBracket Then
                            Stop
                        End If
                        didRBracket = rBracket
                        FindText
                        Exit Sub
                        ' If "<" comes before the found text and   '
                        ' ">" comes after it, then ignore the find '
                        ' because it is part of an HTML command.   '
                    End If
                    lBracket = InStr(lBracket + 1, t, "<")
                Loop
            End If
        End If
    Loop While notWholeWord And lng_TxtPtr > 0
    
    If lng_TxtPtr > 0 Then
        If Not bln_Replace Then
            .SelStart = lng_TxtPtr + 200
            ' This is so that the found text won't '
            ' be displayed on the bottom line, but '
            ' this is not needed when doing a      '
            ' Replace All.                         '
        End If
        
        Highlight lng_TxtPtr - 1, Len(tf)
        
        If bln_Replace Then 'Replace All
            .SelText = txtReplace
            int_NumReplaced = int_NumReplaced + 1
            int_NumFound = int_NumFound + 1
            lng_TxtPtr = lng_TxtPtr + Len(txtReplace)
            FindText
        Else
            int_NumFound = int_NumFound + 1
            cmdReplace(0).Enabled = True
            Me.SetFocus
        End If
    Else
        If optDirection(1) Then
            ret = MsgBox("Not found." & vbCr & _
                         "Search from top?" _
                         , vbYesNo)
            
            If ret = vbYes Then
                optDirection(0).Value = True
                FindText
                Exit Sub
            End If
        
        Else
            MsgBox "Source has been searched." & vbCr _
                   & Str$(int_NumFound) _
                   & " matches found." & vbCr _
                   & Str$(int_NumReplaced) _
                   & " changes made."
            int_NumFound = 0
            int_NumReplaced = 0
        End If
        If optDirection(0) Then lng_TxtPtr = 0
    End If
End With

End Sub


Private Sub Highlight(ByVal ss As Long, _
                      ByVal sl As Long)
    ' Highlight found text: '
    lng_HighlightStart = ss
    lng_HighlightLength = sl
    With BidBaseFileViewer.Text2(1)
        .SelStart = lng_HighlightStart
        .SelLength = lng_HighlightLength
        .SelColor = vbRed
        '.SelText = ""
        .SelBold = True
    End With
End Sub


Private Sub UnHighLight()
    ' Unhighlight highlighted text, if any: '
    If lng_HighlightLength > 0 Then
        With BidBaseFileViewer.Text2(1)
            .SelStart = lng_HighlightStart
            .SelLength = lng_HighlightLength
            .SelColor = vbBlack
            
            .SelBold = False
        End With
    End If
    lng_HighlightStart = 0
    lng_HighlightLength = 0
End Sub

