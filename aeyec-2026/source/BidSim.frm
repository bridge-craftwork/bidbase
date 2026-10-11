VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BidSim 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CardShark BidBase Bidding Simulator"
   ClientHeight    =   9090
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton b_RunSim 
      Caption         =   "&Abort"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7245
      TabIndex        =   19
      Top             =   8580
      Width           =   1260
   End
   Begin VB.CheckBox ck_Log 
      Caption         =   "Log Hands Counted"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   3045
      TabIndex        =   6
      Top             =   8565
      Width           =   2685
   End
   Begin VB.CheckBox ck_Log 
      Caption         =   "Log Matching Deals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   225
      TabIndex        =   5
      Top             =   8550
      Width           =   2385
   End
   Begin VB.CommandButton b_RunSim 
      Caption         =   "&Run Sim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   7
      Top             =   8580
      Width           =   1260
   End
   Begin MSComDlg.CommonDialog CmDia 
      Left            =   11475
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox t_Inp 
      BackColor       =   &H0000FFFF&
      Height          =   300
      Index           =   1
      Left            =   225
      TabIndex        =   1
      Top             =   2535
      Width           =   1155
   End
   Begin VB.TextBox t_Inp 
      BackColor       =   &H0000FFFF&
      Height          =   300
      Index           =   0
      Left            =   3540
      TabIndex        =   0
      Top             =   570
      Width           =   950
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6030
      Index           =   1
      Left            =   165
      TabIndex        =   12
      Top             =   2175
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   10636
      _Version        =   393216
      Rows            =   22
      Cols            =   12
      FixedCols       =   0
      RowHeightMin    =   280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmb_SimVul 
      Height          =   315
      ItemData        =   "BidSim.frx":0000
      Left            =   1380
      List            =   "BidSim.frx":0002
      TabIndex        =   4
      Top             =   885
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1815
      Index           =   0
      Left            =   2520
      TabIndex        =   10
      Top             =   225
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   5
      Cols            =   10
      RowHeightMin    =   285
   End
   Begin VB.TextBox t_MaxMatches 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1380
      TabIndex        =   3
      Text            =   "1000"
      ToolTipText     =   "Quit dealing at this # of matches."
      Top             =   555
      Width           =   975
   End
   Begin VB.TextBox t_MaxDeals 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1380
      TabIndex        =   2
      Text            =   "2000000"
      ToolTipText     =   "Most # of deals to use."
      Top             =   210
      Width           =   975
   End
   Begin VB.Label L_SimName 
      Height          =   240
      Left            =   5955
      TabIndex        =   20
      Top             =   0
      Width           =   4740
   End
   Begin VB.Label Label1 
      Caption         =   "Data-To-Count Specifications:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   195
      TabIndex        =   18
      Top             =   1950
      Width           =   3555
   End
   Begin VB.Label Label1 
      Caption         =   "Hand/Suit Specifications:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   2520
      TabIndex        =   17
      Top             =   -15
      Width           =   3555
   End
   Begin VB.Label L_Matches 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1365
      TabIndex        =   16
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label L_Deals 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1350
      TabIndex        =   15
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "# of Matches:"
      Height          =   240
      Index           =   4
      Left            =   180
      TabIndex        =   14
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "# of Deals:"
      Height          =   240
      Index           =   3
      Left            =   390
      TabIndex        =   13
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Vulnerability:"
      Height          =   240
      Index           =   2
      Left            =   345
      TabIndex        =   11
      Top             =   915
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Max Matches:"
      Height          =   240
      Index           =   1
      Left            =   195
      TabIndex        =   9
      Top             =   570
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "Max # of Deals:"
      Height          =   240
      Index           =   0
      Left            =   195
      TabIndex        =   8
      Top             =   232
      Width           =   1185
   End
   Begin VB.Menu mF 
      Caption         =   "&File"
      Begin VB.Menu mLoad 
         Caption         =   "&Load Sim"
         Shortcut        =   ^L
      End
      Begin VB.Menu mSave 
         Caption         =   "&Save Sim"
         Shortcut        =   ^S
      End
      Begin VB.Menu z1 
         Caption         =   "-"
      End
      Begin VB.Menu mQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mE 
      Caption         =   "&Edit"
      Begin VB.Menu mClearGrids 
         Caption         =   "Clear top grid"
         Index           =   0
         Shortcut        =   {F11}
      End
      Begin VB.Menu mClearGrids 
         Caption         =   "Clear bottom grid"
         Index           =   1
         Shortcut        =   {F12}
      End
      Begin VB.Menu mClearGrids 
         Caption         =   "Clear both grids"
         Index           =   2
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mH 
      Caption         =   "&Help"
      Begin VB.Menu mHelp 
         Caption         =   "Sim &Help File"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "BidSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CardShark BidBase Bidding Simulator '
' Purpose: Given the bidding and/or known '
' cards held by some or all players, the  '
' program randomly deals cards meeting    '
' those specifications to see how many    '
' time specified cards or held by each    '
' player and/or how often they make a     '
' specified bid.

' Project:                       '
'  - References:                 '
'       OLE Automation           '
'  -Components:                  '
'       MS Common Dialog Control '
'       MS Flexgrid Control      '

Option Explicit

Private Const Version$ = "2018-09-15"
    
Public NewBid As BidBaseDB.Bidder ' This is Bidder.cls in BidBaseDB.vbp '

Private bln_Continue As Boolean
Private lng_DealNum As Long
Private int_FiNum(2) As Integer
Private str_FiPath As String
Private str_HandsCountedFile As String
Private HCP(4, 4) As Long
Private str_MatchingDealsFile As String
Private lng_MaxDeals As Long
Private lng_MaxMatches As Long
Private int_MinHCPs(4, 4) As Integer
Private int_MaxHCPs(4, 4) As Integer
Private int_MinSuitLen(4, 4) As Integer 'player #: 1=S 4=E, suit #: 1=C, 4=S
Private int_MaxSuitLen(4, 4) As Integer
Private int_NumCounts As Integer
Private lng_NumMatches As Long
Private int_OrigCol(1) As Integer
Private int_OrigRow(1) As Integer
Private int_NumRandomHands As Integer
Private bln_RandomHand(4) As Boolean
Private str_SimBids(8) As String
Private str_SimFile As String
Private str_SimSp(25, 9) As String
Private lng_StartingDealNum As Long
Private bln_Suit(4, 4) As Boolean
Private int_SuitLen(4, 4) As Integer
Private SuitQty(4, 4) As Long
Private int_Turn As Integer
Private str_vul As String
Private str_Call(30) As String


Private Sub FixGrids()
    Grid1(0).Row = 0
    Grid1(0).Col = 0
    Grid1(0) = " Player"
    Grid1(0).Col = 1
    Grid1(0) = "   Hands (SHDC)"
    Grid1(0).Col = 2
    Grid1(0) = " # Spds"
    Grid1(0).Col = 3
    Grid1(0) = " # Hrts"
    Grid1(0).Col = 4
    Grid1(0) = " # Dias"
    Grid1(0).Col = 5
    Grid1(0) = " # Clbs"
    Grid1(0).Col = 6
    Grid1(0) = "S HCP"
    Grid1(0).Col = 7
    Grid1(0) = "H HCP"
    Grid1(0).Col = 8
    Grid1(0) = "D HCP"
    Grid1(0).Col = 9
    Grid1(0) = "C HCP"
    Grid1(0).Col = 10
    Grid1(0) = "  HCPs"
    Grid1(0).Col = 11
    Grid1(0) = "  Bids-1"
    Grid1(0).Col = 12
    Grid1(0) = " Bids-2"
    Grid1(0).Col = 0
    
    Dim i As Integer
    Dim c As Integer
    
    For i = 1 To 4
        Grid1(0).Row = i
        Grid1(0) = "" & i & "       "
    Next
  
    For c = 11 To 12
        Grid1(0).Col = c
        For i = 1 To 4
            Grid1(0).Row = i
            Grid1(0).CellAlignment = 6 'right justify
        Next
    Next
  
    Grid1(1).Col = 0
    Grid1(1).Row = 0
    Grid1(1) = "    Count:"
  
    For i = 1 To 9
        Grid1(1).Col = i
        Grid1(1) = "     And:"
    Next
    Grid1(1).Col = 11
    Grid1(1) = "  # Found:"
    Grid1(1).Col = 12
    Grid1(1) = "  Percent:"
End Sub


Private Sub b_RunSim_Click(Index As Integer)
    ' BidSim has two buttons, b_RunSim(0) and (1).'
    ' The Index=0 button is the "Run Sim" button. '
    ' Index=1 is "Abort" when a sim is running or '
    ' "Continue" after "Abort" has been clicked.  '
    
    Dim dl As Long
            
    If b_RunSim(1).Caption = "&Continue" Then
        
        If Index = 1 Then
            ' Continue was clicked. '
            If L_Matches = t_MaxMatches Then
                MsgBox "# of matches already equals Max Matches"
                t_MaxMatches.SetFocus
                Exit Sub
            End If
            bln_Continue = True
    
        ElseIf Index = 0 Then
            ' Run Sim was clicked. Disable it and '
            ' change the Continue button to Abort.'
            b_RunSim(0).Enabled = False
            b_RunSim(1).Caption = "&Abort"
        End If
    
    ElseIf b_RunSim(1).Enabled Then
        ' When the sim is running, '
        ' the button causes Abort: '
        bln_AbortSim = True
        Exit Sub
  
    Else 'The Run Sim button
        bln_Continue = False
        b_RunSim(0).Enabled = False
        b_RunSim(1).Enabled = True
    End If
    
    DoLogFiles
  
    Screen.MousePointer = vbHourglass
    
    bln_AbortSim = False
    
    Erase str_Hands()
    Erase int_MinHCPs()
    Erase int_MaxHCPs()
    Erase int_MinSuitLen()
    Erase int_MaxSuitLen()
    Erase str_SimBids()
    Erase str_SimSp()
    Erase bln_Suit()
    
    lng_MaxDeals = Val(t_MaxDeals)
    lng_MaxMatches = Val(t_MaxMatches)
  
    If bln_Continue Then
        b_RunSim(0).Enabled = False
        b_RunSim(1).Enabled = True
        b_RunSim(1).Caption = "&Abort"
        lng_DealNum = Val(L_Deals) + 1
        lng_NumMatches = Val(L_Matches)
  
    Else
        If L_Deals <> "" Then
            lng_DealNum = Val(L_Deals)
        End If
        dl = lng_DealNum + 1
        dl = InputBox("Enter a starting deal number:", "BidSim", dl)
        If dl = 0 Then
            Exit Sub
        End If
        lng_DealNum = dl
        lng_StartingDealNum = lng_DealNum
    End If
  
    PutSpecsIntoVariables
      
    bln_Continue = False
    
    Do
        ShuffleCards
        CountMatchingSpecs
        GetMatchingDeal
        lng_DealNum = lng_DealNum + 1
    Loop While lng_DealNum - lng_StartingDealNum < lng_MaxDeals
      
    WrapUpSim

    DisplayLogs
      
    Screen.MousePointer = 0
End Sub


Private Sub WrapUpSim()
    Dim i As Integer
    Dim avg As Single

    On Local Error Resume Next
    For i = 1 To int_NumCounts
        Grid1(1).Row = i
        Grid1(1).Col = 11
        If lng_NumMatches > 1 Then
            avg = Fix(Val(Grid1(1)) / lng_NumMatches * 100 + 0.5)
            Grid1(1).Col = 12
            Grid1(1) = "" & avg
        End If
    Next
    On Local Error GoTo 0
  
    b_RunSim(0).Enabled = True
    b_RunSim(1).Caption = "&Continue"
  
    int_NumRandomHands = 0
    bln_Continue = False
  
    ' I don't know why this is done: '
    lng_DealNum = lng_StartingDealNum
    For i = 0 To 1
        Grid1(i).Col = int_OrigCol(i)
        Grid1(i).Row = int_OrigRow(i)
    Next
    
    Close #int_FiNum(1)
    Close #int_FiNum(2)
  
    b_RunSim(1).Enabled = True
End Sub


Private Sub DisplayLogs()
    Dim ret As Long
    
    If ck_Log(0).Value <> 0 And _
       ck_Log(1).Value <> 0 _
    Then
        ret = Shell("BidBase FileViewer.exe " _
                        & App.Path & "\Sims\" & str_MatchingDealsFile & " " _
                        & App.Path & "\Sims\" & str_HandsCountedFile, _
                        vbNormalFocus)
  
    ElseIf ck_Log(0).Value <> 0 Then
        ret = Shell("BidBase FileViewer.exe " _
                    & App.Path & "\Sims\" & str_MatchingDealsFile, _
                    vbNormalFocus)
    
    ElseIf ck_Log(1).Value <> 0 Then
    
        ret = Shell("BidBase FileViewer.exe " _
                    & App.Path & "\Sims\" & str_HandsCountedFile, _
                    vbNormalFocus)
    End If
End Sub


Private Sub CountMatchingSpecs()
    ' Bid specs in grid match the deal, '
    ' so now count the data specified   '
    ' eg:  1:hq=5
    '      123456
    
    Dim r As Integer
    Dim c As Integer
    Dim playerNum As Integer
    Dim suitNum As Integer
    Dim e1 As String
    Dim spec As String
    Dim specType As String
    Dim e2n As Integer
    Dim compSign As String
    Dim cn2 As Integer
    Dim specQty As Integer
    Dim n2s As String
    Dim specsMatch As Boolean
    Dim bd As Integer
    Dim pt As Integer
    Dim dn As Integer
    
    If ck_Log(0).Value <> 0 Then
        PrintLog0
    End If
    
    For r = 1 To int_NumCounts
        c = 0
        specsMatch = True
        Do While str_SimSp(r, c) <> "" _
             And specsMatch
            
            CheckCountSpecs str_SimSp(r, c), _
                            playerNum, _
                            suitNum, _
                            specType, _
                            compSign, _
                            n2s, _
                            specQty

            If bln_AbortSim Then Exit Sub
          
            specsMatch = CountSpecsMatch( _
                            playerNum, _
                            suitNum, _
                            specType, _
                            compSign, _
                            n2s, _
                            specQty)
            c = c + 1
        Loop
      
        ' Increment count: '
        Grid1(1).Col = 11
        Grid1(1).Row = r
        Grid1(1) = "" & (Val(Grid1(1)) + 1)
      
        ' Print to log: '
        If ck_Log(1).Value > 0 Then
            Print #int_FiNum(2), _
                playerNum; " "; _
                Right(" " + str_Call(1), 2); "-"; _
                Right(" " + str_Call(2), 2); "-"; _
                Right(" " + str_Call(3), 2); "-"; _
                Right(" " + str_Call(4), 2); ","; _
                Right(" " + str_Call(5), 2); "-"; _
                Right(" " + str_Call(6), 2); "-"; _
                Right(" " + str_Call(7), 2); "    "; _
                "HCPs:"; Right$("  " & HCP(playerNum, 0), 3); "  "; _
                "SHDC:"; _
                    SuitQty(playerNum, Spd); _
                    SuitQty(playerNum, Hrt); _
                    SuitQty(playerNum, Dia); _
                    SuitQty(playerNum, Clb); "    "; _
                    "Deal #"; lng_DealNum
            
            Dim dis As Long
            Print #int_FiNum(2), _
                playerNum; " "; bd; pt; _
                dis; str_Hands(playerNum); dn
            ' The above makes no sense. "bd, pt, dis and dn"  '
            ' are not used/valued anywhere, but I DIMMED them '
            ' in this Sub so that it would compile. '
        End If
        
        DoEvents
        If bln_AbortSim Then
            bln_AbortSim = False
            Exit Sub
        End If
    Next

End Sub


Private Function CountSpecsMatch( _
                    playerNum, _
                    suitNum, _
                    specType, _
                    compSign, _
                    n2s, _
                    specQty) _
                 As Boolean
                 
    Dim bn As Long
    Dim cd As Long
    Dim suit(4, 4, 14) As Integer
                 
    ' Compare Count specs to hand: '
    Select Case specType
        
    Case "q"
        If compSign = "<" And _
            SuitQty(playerNum, suitNum) >= specQty _
        Then
            Exit Function
                
        ElseIf compSign = "=" And _
            SuitQty(playerNum, suitNum) <> specQty _
        Then
            Exit Function
                
        ElseIf compSign = ">" And _
            SuitQty(playerNum, suitNum) <= specQty _
        Then
            Exit Function
        End If
            
    Case "p"
        If compSign = "<" And _
            HCP(playerNum, suitNum) >= specQty _
        Then
            Exit Function
                
        ElseIf compSign = "=" And _
                HCP(playerNum, suitNum) <> specQty _
        Then
            Exit Function
                
        ElseIf compSign = ">" And _
                HCP(playerNum, suitNum) <= specQty _
        Then
            Exit Function
        End If
          
    Case "1", "2"
        ' Bid number of the 8 bids in the grid,
        ' e.g.: player 2's bid numbers are 2 and 6.
        bn = playerNum + 4 * Val(specType - 1)
        If InStr("Pass Rdbl Dbl", str_Call(bn)) > 0 Then
            If compSign = "=" And _
                n2s <> Left$(str_Call(bn), 1) _
            Then
                Exit Function
                    
            ElseIf compSign = "<" Then
                ' If the spec (n2s) is 3:b1<D or <R '
                ' or <P there really isn't any bid  '
                ' that can meet that spec, so go on:'
                Exit Function
                    
            ElseIf compSign = ">" And _
                    Val(str_Call(bn)) = 0 _
            Then
                ' If the spec is 3:b2>P or >D or >R, '
                ' then any number bid will do, from  '
                ' 1C to 7N. Only a bid not starting  '
                ' with a number, such as P, D, R,    '
                ' will fail.                         '
                If Right$(n2s, 1) = "N" And _
                    Val(str_Call(bn)) > 0 _
                Then
                    If str_Call(bn) <= Val(n2s) Then
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            End If
                
        ElseIf InStr("PDR", n2s) > 0 Then
            If compSign = "=" And _
                n2s <> Left$(str_Call(bn), 1) _
            Then
                Exit Function
                    
            ElseIf compSign = ">" And _
                    Val(str_Call(bn)) = 0 _
            Then
                Exit Function
                    
            ElseIf compSign = "<" Then
                Exit Function
            End If
            
        ElseIf compSign = "=" And _
                n2s <> str_Call(bn) _
        Then
            Exit Function
            
        ElseIf compSign = "<" And _
                str_Call(bn) >= n2s _
        Then
            Exit Function
            
        ElseIf compSign = ">" Then
            If Right$(n2s, 1) = "N" And _
                Val(str_Call(bn)) > 0 _
            Then
                If Val(str_Call(bn)) <= Val(n2s) Then
                    Exit Function
                End If
              
            ElseIf str_Call(bn) <= n2s Then
                Exit Function
            End If
        End If
            
    Case "c"
        If compSign = "=" Then
            If n2s = "" Then
                cd = 0
            Else
                cd = InStr("A23456789TJQK", n2s)
            End If
                    
            If suit(playerNum, suitNum, cd) = 0 Then
                ' Can't find where this is assigned a value. '
                Exit Function
            End If
        End If
    End Select
    
    CountSpecsMatch = True
                 
End Function


Private Sub CheckCountSpecs(ByVal spec As String, _
                playerNum As Integer, _
                suitNum As Integer, _
                specType As String, _
                compSign As String, _
                n2s As String, _
                specQty As Integer)
                
    Dim e1 As String
                
    ' Validate spec:  '
    ' eg:     1:hq=5  '
    ' or:     3:tq>10 '
    ' or:     2:b1=2S '
    ' or:     2:hc=A  '
    ' char# = 123456  '
            
    ' 1st character: '
    playerNum = Val(spec) 'player number
    If playerNum < 1 Or playerNum > 4 Then
        MsgBox "The player number '" & playerNum & _
                "' in '" & spec & _
                "' is not valid." & vbCr & _
                "Simulation will terminate."
        bln_AbortSim = True
        Exit Sub
    End If
          
    ' 2nd char is ":", so skip to '
    ' 3rd character:              '
    e1 = Mid$(spec, 3, 1)
    suitNum = InStr("tcdhsb", e1)
    If suitNum = 0 Then
        MsgBox "The character '" & e1 & _
                "' in '" & spec & _
                "' is not valid." & vbCr & _
                "Valid characters are " & _
                "t, c, d, h, s, b" & _
                "Simulation will terminate."
        bln_AbortSim = True
        Exit Sub
    Else
        suitNum = suitNum - 1 'so t=0, c=1, d=2, h=3, s=4, b=5
    End If
          
    ' 4th character: '
    specType = Mid$(spec, 4, 1)
    ' p (HPCs) or q (Qty) or '
    ' c (specific card) or   '
    ' 1 or 2 (bid round)     '
          
    If InStr("pqc12", specType) = 0 Then
        MsgBox "The character '" & specType & _
                "' in '" & spec & _
                "' is not valid." & vbCr & _
                "Acceptable characters are " & _
                "p, q, c, 1 and 2." & vbCr & _
                "Simulation will terminate."
        bln_AbortSim = True
        Exit Sub
    End If

    ' 5th character: '
    compSign = Mid$(spec, 5, 1) ' <, =, >
    If InStr("<=>", compSign) = 0 Then
        MsgBox "The character '" & compSign & _
                "' in '" & spec & _
                "' is not valid." & vbCr & _
                "Acceptable characters are " & _
                "<, =, and >." & vbCr & _
                "Simulation will terminate."
        bln_AbortSim = True
        Exit Sub
    End If
          
    ' 6th character: '
    n2s = Trim$(Mid$(spec & " ", 6, 2)) '# pts or cards
    specQty = Val(n2s)
        
End Sub


Private Sub PrintLog0()
    Dim i As Integer
    
        Print #int_FiNum(1), "Deal number: "; lng_DealNum
        For i = 1 To 4
            Print #int_FiNum(1), "  "; i; _
                str_Hands(i); _
                "  Bid:" & Right$("  " & str_Call(i), 2); _
                ", " & Right$("  " & str_Call(i + 4), 2); _
                "  HCPs:" & Right$("  " & HCP(i, 0), 3); _
                "  SHDC:"; _
                 SuitQty(i, Spd); _
                 SuitQty(i, Hrt); _
                 SuitQty(i, Dia); _
                 SuitQty(hn, Clb)
        Next
End Sub


Private Sub MakeBid(h As Integer)
    Dim b As String
    Dim c As Integer
    Dim d As String
    Dim BON As String
    Dim ID As Long

    'Submit specs to DLL:
    With NewBid
        .DBPathAndName = App.Path & "\BidBase.accdb"
        .PlayersHand = str_Hands(h)
        .PriorCalls = str_AllCalls
        b = .Bid
        c = .ErrCode
        d = .Disclosure
        BON = .BON
        ID = .ID
    End With
    
    str_Call(int_CallNum) = b
    
End Sub


Private Sub GetMatchingDeal()
        Dim p As Integer
        Dim s As Integer
        
        ScrambleHands
        
        DoEvents
        If bln_AbortPlay Then
            GoTo ExitAnalyze
        End If
        
        'DealerSpecified = 1
        'SetPlayerNumbers
        
        ' Filter HCP and suit length specs '
        ' before going to bidding:         '
        For p = 1 To 4
            If bln_RandomHand(p) Then
              For s = 0 To 4
                If HCP(p, s) < int_MinHCPs(p, s) Then GoTo GetNextDeal
                If HCP(p, s) > int_MaxHCPs(p, s) Then GoTo GetNextDeal
                If s > 0 Then
                    If SuitQty(p, s) < int_MinSuitLen(p, s) Then GoTo GetNextDeal
                    If SuitQty(p, s) > int_MaxSuitLen(p, s) Then GoTo GetNextDeal
                End If
              Next
            End If
        Next
    
        Dim str_AllCalls As String
        'ReDim str_Call(50) As String
        Dim simNum As Integer
        Dim firstBidNum As Integer
        Dim sel As String
        Dim int_CallNum As Long
        
        str_AllCalls = ""
    
        int_CallNum = 1
        int_Turn = 1
        simNum = 1
        firstBidNum = 0
    
        ' For each specified bid (from the last 2 columns of the   '
        ' top grid, for a total of 8 bids), make sure the bidding  '
        ' for this deal matches. When a bid is found, it is put in '
        ' the str_Call() array using the int_CallNum counter. If a '
        ' bid is not found, str_Call(int_CallNum) comes back empty,'
        ' which is the same as a pass, except that opening passes  '
        ' do not go into str_Call() array. So the int_CallNum counter'
        ' is only incremented starting with the first non-passing  '
        ' bid. Likewise, the str_AllCalls() array does not contain '
        ' opening passes. The simNum counter is incremented each   '
        ' time thru the loop to point to the specified bids. The   '
        ' int_Turn counter is incremented each time thru the loop: '
        ' 1 thru 4 then back to 1 to indicate the hand to bid.     '
    Do While str_SimBids(simNum) <> ""
      If bln_RandomHand(int_Turn) Or str_SimBids(simNum) = "?" Then
        sel = "SELECT * FROM " + Bid1FileName$ + " " _
            & "WHERE Prior_Bids='" + str_AllCalls + "' " _
            & "ORDER BY Bid_Order_Number;"
        Set BidTable = Bid1DB.OpenRecordset(sel$, dbOpenForwardOnly)
        
        int_CallNum = simNum
        LookUpBid
        
        If str_Call(simNum) = "" Or str_Call(simNum) = "Pass" Then
          str_Call(simNum) = "P"
        ElseIf Left$(str_Call(simNum), 1) = "D" Then
          str_Call(simNum) = "D"
        ElseIf Left$(str_Call(simNum), 1) = "R" Then
          str_Call(simNum) = "R"
        End If
        
        If str_Call(int_CallNum) <> str_SimBids(simNum) And _
           str_SimBids(simNum) <> "?" And Not _
           (str_SimBids(simNum) = "P" And _
            str_Call(int_CallNum) = "") _
        Then
            GoTo GetNextDeal
        End If
      Else
        int_CallNum = int_CallNum + 1
        str_Call(simNum) = str_SimBids(simNum)
        'Debug.Print lng_dealNum & ", simn: " & simNum & ", turn: " & int_Turn & ", bidn: " & int_CallNum & ", bid = " & str_Call(simNum) & ", sim = " & str_simBids(simNum) & ", " & str_Hands(int_Turn)
      End If
      If str_AllCalls = "" Then
        If str_Call(simNum) <> "P" And str_Call(simNum) <> "" Then
          str_AllCalls = str_Call(simNum)
          firstBidNum = simNum
        End If
      Else
        str_AllCalls = str_AllCalls + "-" + str_Call(simNum)
      End If
      int_Turn = int_Turn + 1
      If int_Turn = 5 Then int_Turn = 1
      simNum = simNum + 1
      If simNum = 9 Then Exit Do
    Loop
    
    lng_NumMatches = lng_NumMatches + 1
    L_Matches = "" & lng_NumMatches
        L_Matches.Refresh
    
End Sub

Private Sub DoLogFiles()
    
    If ck_Log(0).Value <> 0 Then
        ' Open log file: '
        Close #1  'int_FiNum(1)
        str_MatchingDealsFile = _
            Left$(str_SimFile, _
                  Len(str_SimFile) - 4) _
            & "_matching_deals.txt"
            
        If bln_Continue Then
            Open App.Path & "\Sims\" _
                & str_MatchingDealsFile _
            For Append As #1  'int_FiNum(1)
    
        Else
            Open App.Path + "\Sims\" _
                & str_MatchingDealsFile _
            For Output As #1  'int_FiNum(1)
        
        End If
    End If
  
    If ck_Log(1).Value <> 0 Then
        Close #2  'int_FiNum(2)
        str_HandsCountedFile = _
            Left$(str_SimFile, _
                  Len(str_SimFile) - 4) _
            & "_hands_counted.txt"
    
        If bln_Continue Then
            Open App.Path & "\Sims\" _
                & str_HandsCountedFile _
            For Append As #2  'int_FiNum(2)
    
        Else
            Open App.Path & "\Sims\" _
                & str_HandsCountedFile _
            For Output As #2  'int_FiNum(2)
        
        End If
    End If
    
End Sub


Private Sub PutSpecsIntoVariables()
    'Specs are read into variables because when doing upwards '
    'of 20 million comparisons, comparing variables is much   '
    'faster than manipulating and comparing grid cells.       '
  
    Dim c As Integer
    Dim i As Integer
    Dim s As Integer
    Dim p As Integer
    Dim r As Integer
    Dim x As String
  
    ' Put Hands specified into variables: '
    Grid1(0).Col = 1
    For i = 1 To 4
        Grid1(0).Row = i
        str_Hands(i) = Trim$(Grid1(0))
        
        If str_Hands(i) = "" Then
            bln_RandomHand(i) = True
            int_NumRandomHands = int_NumRandomHands + 1
        Else
            bln_RandomHand(i) = False
            Stop ' access DLL to get array?
            'Bid1Entry.ConvertHandToArray str_Hands(i), i
        End If
    Next
  
    ' Put SUIT LEN specs into variables: '
    For p = 1 To 4
        Grid1(0).Row = p
        For s = 1 To 4
            Grid1(0).Col = 6 - s
            x = Grid1(0)
            If Len(x) > 0 Then
                If Len(x) = 1 Then  ' e.g.: 5 '
                    int_MinSuitLen(p, s) = Val(x)  ' 5 '
                    int_MaxSuitLen(p, s) = int_MinSuitLen(p, s)  ' 5 '
        
                ElseIf Left$(x, 1) = "<" Then  ' <5 '
                    int_MinSuitLen(p, s) = 0
                    int_MaxSuitLen(p, s) = Val(Mid$(x, 2)) - 1  ' 4 '
        
                ElseIf Right$(x, 1) = "+" Then ' 5+ '
                    int_MinSuitLen(p, s) = Val(x)  ' 5  '
                    int_MaxSuitLen(p, s) = 13
        
                ElseIf InStr(x, "-") > 0 Then  ' 5-7 '
                    int_MinSuitLen(p, s) = Val(x)  '  5  '
                    int_MaxSuitLen(p, s) = Val(Mid$(x, InStr(x, "-") + 1))  ' 7 '
        
                Else
                    int_MinSuitLen(p, s) = Val(x)
                    int_MaxSuitLen(p, s) = int_MinSuitLen(p, s)
                End If
            Else
                int_MinSuitLen(p, s) = 0
                int_MaxSuitLen(p, s) = 13
            End If
        Next
    Next
  
    ' Put HCP specs into variables: '
    For s = 0 To 4
        Grid1(0).Col = 10 - s
        
        For p = 1 To 4
            Grid1(0).Row = p
            x = Trim$(Grid1(0)) ' e.g.: 10+ '
            If Len(x) > 0 Then
                If Left$(x, 1) = "<" Then ' <10 '
                    int_MinHCPs(p, s) = 0
                    int_MaxHCPs(p, s) = Val(Mid$(x, 2)) - 1 ' 9 '
                ElseIf Right$(x, 1) = "+" Then  ' 4+ '
                    int_MinHCPs(p, s) = Val(x)  ' 4  '
                    int_MaxHCPs(p, s) = 40
                ElseIf InStr(x, "-") > 0 Then  ' 6-9 '
                    int_MinHCPs(p, s) = Val(x) '  6  '
                    int_MaxHCPs(p, s) = Val(Mid$(x, InStr(x, "-") + 1)) ' 9 '
                Else ' 11 '
                    int_MinHCPs(p, s) = Val(x) ' 11 (no range, just "11") '
                    int_MaxHCPs(p, s) = int_MinHCPs(p, 0) ' 11 '
                End If
            Else
                int_MinHCPs(p, s) = 0
                int_MaxHCPs(p, s) = 40
            End If
        Next
    Next
  
    ' VUL specs: '
    str_vul = cmb_SimVul
  
    ' BIDS specs: '
    For c = 11 To 12
        Grid1(0).Col = c
        For p = 1 To 4
            Grid1(0).Row = p
            i = p + (c - 11) * 4
            str_SimBids(i) = Trim$(Grid1(0))
        Next
    Next
  
    ' COUNTs specified: '
    Grid1(1).Row = 1
    Grid1(1).Col = 0
    Do While Grid1(1) <> ""
        Do
            str_SimSp(Grid1(1).Row, Grid1(1).Col) = Grid1(1)
            If Grid1(1).Col < 9 Then
                Grid1(1).Col = Grid1(1).Col + 1
            Else
                Exit Do
            End If
        Loop While Grid1(1) <> ""
        
        If Grid1(1).Row < Grid1(1).Rows - 1 Then
            Grid1(1).Row = Grid1(1).Row + 1
        Else
            Exit Do
        End If
        Grid1(1).Col = 0
    Loop
  
    Grid1(1).Col = 0
    If Grid1(1) = "" Then
        int_NumCounts = Grid1(1).Row - 1
    Else
        int_NumCounts = Grid1(1).Row
    End If
  
    If bln_Continue Then
        bln_Continue = False
    Else
        For r = 1 To int_NumCounts
            Grid1(1).Row = r
            Grid1(1).Col = 11
            Grid1(1) = ""
            Grid1(1).Col = 12
            Grid1(1) = ""
        Next
        L_Deals = ""
        L_Matches = ""
    End If
    
End Sub


Private Sub ck_Log_Click(Index As Integer)
    If ck_Log(Index).Value = 0 Then
        Close #int_FiNum(Index)
    End If
End Sub


Private Sub Form_Load()
    Dim i As Integer
    Dim c As Integer
    
    Me.Move 0, 0
    Me.Width = 12250
    Me.Height = 9800
  
    Grid1(0).Height = 1550
    Grid1(0).Width = 9360
    Grid1(0).Cols = 13
    For i = 0 To 12
        If i = 1 Then
            Grid1(0).ColWidth(i) = 1520
        Else
            Grid1(0).ColWidth(i) = 640
        End If
    Next
    Grid1(0).RowHeight(4) = 310
    
    Me.Caption = Me.Caption & "     ver." & Version$
  
    Grid1(1).Move Label1(0).Left, Grid1(0).Top + Grid1(0).Height + 300
    Label1(6).Top = Grid1(1).Top - 230
    Grid1(1).Cols = 13
    Grid1(1).Width = 11800
    Grid1(1).Height = 6390
    Grid1(1).Rows = 22
    Grid1(1).ColWidth(0) = 1100
    For c = 1 To 12
        Grid1(1).ColWidth(c) = 955
    Next
    Grid1(1).ColWidth(10) = 50
    Grid1(1).RowHeight(21) = 310
  
    FixGrids
  
    int_OrigRow(0) = 1
    int_OrigCol(0) = 1
    Grid1(0).Row = 1
    Grid1(0).Col = 1
  
    int_OrigRow(1) = 1
    int_OrigCol(1) = 0
    Grid1(1).Row = 1
    Grid1(1).Col = 0
  
    t_Inp(0) = Grid1(0)
    t_Inp(0).Width = Grid1(0).ColWidth(1)
    t_Inp(0).Move 3200, 540
    t_Inp(0) = ""
  
    t_Inp(1) = Grid1(1)
    t_Inp(1).Width = Grid1(1).ColWidth(0)
    t_Inp(1).Move 240, 2400
    t_Inp(1) = ""
  
    b_RunSim(0).Enabled = True
    b_RunSim(1).Caption = "&Abort"
    b_RunSim(1).Enabled = False
    BidSim.Show
    t_Inp(0).SetFocus
    
    Set NewBid = New BidBaseDB.Bidder
    Randomize
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    Set NewBid = Nothing

End Sub

Private Sub Grid1_Click(Index As Integer)
    If Index = 1 And Grid1(Index).Col > 9 Then
        Grid1(Index).Col = int_OrigCol(1)
        Grid1(Index).Row = int_OrigRow(1)
    End If
  
    ' When the user clicks on the grid,  '
    ' the input box moves to the cell on '
    ' which the user clicked:            '
  
    Dim n As Integer
    Dim i As Integer
    Dim t As Integer
    Dim adj As Integer
    Dim c As Integer
    Dim r As Integer
  
With Grid1(Index)
    'Left coordinate of input box:
    n = .Left + 30
    For i = 0 To .Col - 1
        n = n + .ColWidth(i)
    Next
  
    'Top coordinate:
    t = .Top + 315
    t = t _
      + .RowHeight(1) _
        * (.Row - 1)
  
    If .Col < .Cols - 1 Then
        adj = 30
    Else
        adj = 0
    End If
    
    t_Inp(Index).Width = .ColWidth(.Col) + adj
    t_Inp(Index).Move n, t
  
    ' Put data from input box into prior cell; '
    ' put data from new cell into input box:   '
    c = .Col
    r = .Row
    .Col = int_OrigCol(Index)
    .Row = int_OrigRow(Index)
    .Text = t_Inp(Index)
    .Col = c
    .Row = r
    int_OrigCol(Index) = c
    int_OrigRow(Index) = r
    t_Inp(Index) = .Text
    .SetFocus 'Just to get focus from t_Inp for the next line: '
    t_Inp(Index).SetFocus
End With

End Sub

Private Sub mClearGrids_Click(Index As Integer)
    If Index < 2 Then 'Clear one
        Grid1(Index).Clear
        t_Inp(Index) = ""
  
    Else 'Clear both
        Grid1(0).Clear
        Grid1(1).Clear
        t_Inp(0) = ""
        t_Inp(1) = ""
    End If
  
    FixGrids
    L_Deals = ""
    L_Matches = ""
    b_RunSim(0).Enabled = True
    b_RunSim(1).Enabled = False
    b_RunSim(1).Caption = "&Abort"
    L_SimName.Caption = ""
    bln_Continue = False
End Sub


Private Sub mHelp_Click()
    Dim f As String
    Dim ret As Long
    
    f = App.Path & "\Notes\CardShark_BidBase_Bid_Simulator.htm"
    ret = Shell("BidBase File Viewer.exe " & f, vbNormalFocus)
End Sub


Private Sub mLoad_Click()
    ' Load a simulation spec file: '
    
    Dim g As Integer
    Dim fiNum As Integer
    Dim x As String
    Dim i As Integer
    Dim p As Integer
    Dim s As Integer
    Dim b As Integer
    Dim rslt1 As Long
    Dim rslt2 As Long
    
    With CmDia
        .DialogTitle = "Select simulation spec file."
        .FileName = App.Path & "\Sims\*.sim"
        .ShowOpen
        str_SimFile = .FileTitle
        If str_SimFile = "" Then Exit Sub
    End With
    
    mClearGrids_Click 2
    
    For g = 0 To 1
        Grid1(g).Col = 1 - g
        Grid1(g).Row = 1
        Grid1_Click g
    Next
  
    'On Local Error GoTo SimLoadError
  
    fiNum = FreeFile
    Open App.Path & "\Sims\" & str_SimFile For Input As #fiNum
    Input #fiNum, x 'MAX DEALS:
    Input #fiNum, x: t_MaxDeals = x
    Input #fiNum, x 'MAX MATCHES:
    Input #fiNum, x: t_MaxMatches = x
    Input #fiNum, x 'VUL:
    Input #fiNum, x: cmb_SimVul = x
  
    Input #fiNum, x 'HANDS:
    
With Grid1(0)
    .Col = 1
    For i = 1 To 4
        Input #fiNum, x
        .Row = i
        .Text = Trim$(Mid$(x + " ", 3))
        If i = 1 Then t_Inp(0) = Grid1(0)
    Next
  
    Input #fiNum, x 'SUIT LEN:
    
    For p = 1 To 4
        .Row = p
        For s = 1 To 4
            .Col = s + 1
            Input #fiNum, x
            .Text = Trim$(Mid$(x + " ", 2))
        Next
    Next
  
    Input #fiNum, x 'HCPs:
    For p = 1 To 4
        .Row = p
        For s = 0 To 4
            .Col = s + 6
            Input #fiNum, x
            .Text = Trim$(Mid$(x & " ", 2))
        Next
    Next
  
    Input #fiNum, x 'BIDS:
    .Col = 11
    For p = 1 To 4
        Input #fiNum, x
        .Row = p
        .Text = Trim$(Mid$(x & " ", 3))
    Next
    
    .Col = 12
    For p = 1 To 4
        Input #fiNum, x
        .Row = p
        .Text = Trim$(Mid$(x & " ", 3))
    Next
End With

With Grid1(1)
    Input #fiNum, x 'COUNT:
    .Col = 0
    .Row = 0
    Do While Not EOF(1) _
             And x <> "RESULTS:" _
             And .Row < .Rows - 1
        
        .Row = .Row + 1
        Line Input #fiNum, x
        x = Trim$(x)
    
        Do While x <> "" And x <> "RESULTS:"
            If Left$(x, 3) = "and" Then
                x = Trim$(Mid$(x, 5))
            End If
      
            b = InStr(x, " ")
            If b = 0 Then b = Len(x) + 1
            .Text = Mid$(x, 1, b - 1)
            
            If .Col = 0 And _
               .Row = 1 _
            Then
                t_Inp(1) = .Text
            End If
            
            .Col = .Col + 1
            x = Trim$(Mid$(x + " ", b + 1))
        Loop
        .Col = 0
    Loop
  
    If Not EOF(1) Then
        If x <> "RESULTS:" Then
            Input #fiNum, x 'RESULTS:
        End If
    
        Input #fiNum, x ' Deals
        L_Deals = Trim$(Mid$(x, 7))
    
        Input #fiNum, x ' Matches
        L_Matches = Trim$(Mid$(x, 9))
        .Row = 0
    
        Do While Not EOF(1) _
                 And Grid1(1).Row < Grid1(1).Rows - 1
            
            .Row = .Row + 1
            Input #fiNum, rslt1, rslt2
            .Col = 11
            .Text = rslt1
            .Col = 12
            .Text = rslt2
        Loop
    
        b_RunSim(1).Enabled = True
        b_RunSim(1).Caption = "&Continue"
    End If
End With

ExitLoadSim:
    On Local Error GoTo 0
    Close #fiNum
    Grid1(0).Col = 1
    Grid1(0).Row = 1
    Grid1(1).Col = 0
    Grid1(1).Row = 1
    L_SimName.Caption = str_SimFile
Exit Sub
  
SimLoadError:
    MsgBox "Error in Sim File format."
    Resume ExitLoadSim
End Sub


Private Sub mQuit_Click()
    Unload Me
End Sub


Private Sub mSave_Click()
    ' Save to grid any changes made in input box: '
    Dim g As Integer
    Dim ret As Long
    Dim fiNum As Integer
    Dim p As Integer
    Dim s As Integer
    Dim r As Integer
    Dim lastRow As Integer
    Dim rslt1 As Double
    Dim rslt2 As Double
    
    For g = 0 To 1
        Grid1(g).Row = int_OrigRow(g)
        Grid1(g).Col = int_OrigCol(g)
        Grid1(g) = t_Inp(g)
    Next
  
    Do
        CmDia.DialogTitle = "Enter file name."
        If str_SimFile <> "" Then
            CmDia.FileName = App.Path + "\Sims\" + str_SimFile
        Else
            CmDia.FileName = App.Path + "\Sims\*.sim"
        End If
        CmDia.ShowSave
        If CmDia.FileTitle = "" Then Exit Sub
        str_SimFile = CmDia.FileTitle
    
        If Dir$(str_SimFile) <> "" Then
            
            ret = MsgBox("File already exists." & vbCr _
                         & "Overwrite it?", _
                        vbYesNoCancel)
        Else
            ret = vbYes
        End If
    Loop While ret = vbNo
    
    fiNum = FreeFile
    Open App.Path + "\Sims\" & str_SimFile For Output As #fiNum
        Print #fiNum, "MAX DEALS:"
        Print #fiNum, t_MaxDeals
        Print #fiNum, "MAX MATCHES:"
        Print #fiNum, t_MaxMatches
        Print #fiNum, "VUL:"
        Print #fiNum, cmb_SimVul
        Print #fiNum, "HANDS:"
        Grid1(0).Col = 1
        For p = 1 To 4
            Grid1(0).Row = p
            Print #fiNum, "" & p & "=" & Grid1(0)
        Next
        Print #fiNum, "SUIT LEN:"
        For p = 1 To 4
            Grid1(0).Row = p
            For s = 1 To 4
                Grid1(0).Col = s + 1
                Print #fiNum, Mid$("cdhs", s, 1); Grid1(0);
                If s < 4 Then
                    Print #fiNum, ", ";
                Else
                    Print #fiNum, ""
                End If
            Next
        Next
        Print #fiNum, "HCPs:"
        For p = 1 To 4
            Grid1(0).Row = p
            For s = 0 To 4
                Grid1(0).Col = s + 6
                Print #fiNum, Mid$("tcdhs", s + 1, 1); Grid1(0);
                If s < 4 Then
                    Print #fiNum, ", ";
                Else
                    Print #fiNum, ""
                End If
            Next
        Next
        Print #fiNum, "BIDS:"
        For g = 11 To 12
            Grid1(0).Col = g
            For r = 1 To 4
                Grid1(0).Row = r
                Print #fiNum, "" & r + 4 * (g - 11); "="; Grid1(0)
            Next
        Next
        Print #fiNum, "COUNT:"
        Grid1(1).Row = 1
        Grid1(1).Col = 0
        Do
            Do While Grid1(1) <> ""
                If Grid1(1).Col > 0 Then
                    Print #fiNum, " and "; Grid1(1);
                Else
                    Print #fiNum, Grid1(1);
                End If
            
                If Grid1(1).Col < 9 Then
                    Grid1(1).Col = Grid1(1).Col + 1
                Else
                    Exit Do
                End If
            Loop
            Print #fiNum, ""
            If Grid1(1).Row < Grid1(1).Rows - 1 Then
                Grid1(1).Row = Grid1(1).Row + 1
            Else
                Exit Do
            End If
            Grid1(1).Col = 0
        Loop While Grid1(1) <> ""
  
        lastRow = Grid1(1).Row - 1
  
        Print #fiNum, "RESULTS:"
        Print #fiNum, "Deals: "; L_Deals
        Print #fiNum, "Matches: "; L_Matches
        If L_Matches <> "" Then
            For r = 1 To lastRow
                Grid1(1).Row = r
                Grid1(1).Col = 11
                rslt1 = Val(Grid1(1))
                Grid1(1).Col = 12
                rslt2 = Val(Grid1(1))
                Print #fiNum, rslt1, rslt2
            Next
        End If
    Close #fiNum
    
    L_SimName.Caption = str_SimFile
    
    MsgBox str_SimFile & " has been saved"
End Sub


Private Sub t_Inp_GotFocus(Index As Integer)
    t_Inp(Index).SelStart = 0
    t_Inp(Index).SelLength = 32768
End Sub


Private Sub t_Inp_KeyDown( _
                Index As Integer, _
                KeyCode As Integer, _
                Shift As Integer)
  
With Grid1(Index)

    If Shift > 0 Then Exit Sub
        
    If KeyCode = 37 Then 'left arrow
        If .Col > 1 + (Index = 1) And _
           t_Inp(Index).SelStart = 0 And _
           t_Inp(Index).SelLength = 0 _
        Then
            .Col = .Col - 1
            Grid1_Click Index
        End If
  
    ElseIf KeyCode = 39 Then 'right arrow
        If .Col < .Cols - 1 + 3 * (Index = 1) And _
           t_Inp(Index).SelStart = Len(t_Inp(Index)) _
        Then
            .Col = .Col + 1
            Grid1_Click Index
        End If
  
    ElseIf KeyCode = 38 Then 'up arrow
        If .Row > 1 Then
            .Row = .Row - 1
            Grid1_Click Index
        End If
  
    ElseIf KeyCode = 40 Then 'down arrow
        If .Row < .Rows - 1 Then
            .Row = .Row + 1
            Grid1_Click Index
        End If
    End If
End With

End Sub


Private Sub t_Inp_KeyPress(Index As Integer, _
                           KeyAscii As Integer)
    ' This input box overlays the grid   '
    ' and accepts the actual input from  '
    ' the user and passes it to the grid.'
    ' It also gets data already in a cell'
    ' and displays it as the default.    '
    '                                    '
    ' This routine processes each key-   '
    ' press to make sure it is valid and '
    ' to do the other things shown below.'
    ' After this routine, the program    '
    ' continues to process the input, so '
    ' we can change the key pressed to   '
    ' something else by changing KeyAscii'
    ' or we can cancel the keypress by   '
    ' changing KeyAscii to zero.         '
    
    Dim ky As String
    Dim i As Integer
    Dim j As Integer
    Dim ss As Integer
        
    If KeyAscii = vbKeyBack Then
        ' Let Backspace thru: '
        Exit Sub
    End If
    
With t_Inp(Index)

    If KeyAscii = 27 Then
        ' Esc: restore original text: '
        .Text = Grid1(Index).Text
        Exit Sub
    End If
    
    ky = Chr$(KeyAscii)
            
    If Index = 0 Then 'Top grid:
        If Grid1(0).Col = 1 Then
            ' Maximum length = 16 = 13 cards + 3 dashes '
            If Len(.Text) = 16 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            ' Convert a space to a dash to make '
            ' entering hands and bids easier.   '
            If ky = " " Then
                ky = "-"
                KeyAscii = 45
            End If
            
            ' If card letter entered is in '
            ' lowercase, convert to upper: '
            If InStr("akqjt", ky) > 0 Then
                KeyAscii = KeyAscii - 32
                ky = UCase$(ky)
            End If
            
            ' Valid keypresses are AKQJT, 2-9, '
            ' If not one of these, cancel/exit.'
            If InStr("AKQJT", ky) = 0 And _
               (ky < "2" Or ky > "9") And _
               ky <> "-" _
            Then
                KeyAscii = 0
                Exit Sub
            End If
            
            i = InStr("AKQJT98765432", ky)
            ' Validate card entered: '
            If .SelStart > 0 And i > 0 Then
                j = InStr("AKQJT98765432", _
                          Mid$(.Text, .SelStart, 1))
                If j >= i Then 'card entered is out of rank order
                    ss = .SelStart
                    .Text = Left$(.Text, .SelStart) + ky + Mid$(.Text, _
                                  .SelStart + 1)
                    .SelStart = ss + 1
                    Beep
                    MsgBox "Card spec is out of order."
                    KeyAscii = vbKeyBack
                    Exit Sub
                End If
            End If
        End If

    Else 'Bottom ("Count") grid
  
        ' It's important to control input into the Count grid    '
        ' since the sim won't work if the exact format and type  '
        ' of input is not used. A keypress which does not follow '
        ' the required format is cancelled by setting KeyAscii=0.'
        ' If a keypress does not require testing, just exit this '
        ' sub and the key input will continue normally, such as: '
  
        Select Case .SelStart
                  ' .SelStart is the character '
                  ' position minus one.        '
        
            Case 0  'player number
                If InStr("1234", ky) = 0 Then
                    ' Not a valid character. '
                    ' Cancel:                '
                    KeyAscii = 0
                End If
  
            Case 1 'colon
                KeyAscii = Asc(":")
  
            Case 2 'suit, t for total, or b for bid '
                If InStr("CDHSTB", ky) > 0 Then
                    ' Change to lowercase: '
                    KeyAscii = KeyAscii + 32
    
                ElseIf InStr("cdhstb", ky) = 0 Then
                    ' Not a valid character. '
                    ' Cancel:                '
                    KeyAscii = 0
                End If
  
            Case 3 ' Type or bid round '
                If InStr("QPC", ky) > 0 Then
                    ' Change to lowercase: '
                    KeyAscii = KeyAscii + 32
    
                ElseIf InStr("qpc12", ky) = 0 Then
                    ' Not a valid character. '
                    ' Cancel:                '
                    KeyAscii = 0
                End If
  
            Case 4 ' Math symbol: '
                If InStr("<=>", ky) = 0 Then
                    ' Not a valid character. '
                    ' Cancel:                '
                    KeyAscii = 0
                End If
  
            Case Is > 4 'number or bid
                If InStr("akqjtcdhsndprx", ky) > 0 Then
                    ' Change to uppercase: '
                    KeyAscii = KeyAscii - 32
    
                ElseIf InStr("1234567890AKQJTCDHSNPDRX", ky) = 0 Then
                    ' Not a valid character. '
                    ' Cancel:                '
                    KeyAscii = 0
                End If
        
        End Select
    
    End If

End With

End Sub

Private Sub t_Inp_LostFocus(Index As Integer)
    ' Save to grid any changes made in input box: '
    Dim c As Integer
    Dim r As Integer
    
    With Grid1(Index)
        c = .Col
        r = .Row
        .Row = int_OrigRow(Index)
        .Col = int_OrigCol(Index)
        .Text = t_Inp(Index)
        .Col = c
        .Row = r
    End With
End Sub

