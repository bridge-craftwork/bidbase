Attribute VB_Name = "BBEval"
'Client input variables:
Private str_PlayersHand As String
Private str_Vul As String
Private str_OrigData As String
'Card naming constants:
Global Const Ace = 1
Global Const King = 13
Global Const Queen = 12
Global Const Jack = 11
Global Const Ten = 10
Global Const Nine = 9
Global Const Eight = 8
'Suit naming constants:
Global Const Clbs = 1 ' Clb set
Global Const Dias = 2
Global Const Hrts = 3
Global Const Spds = 4
' Player's holdings: '
Private bln_AlreadyHaveOPsData ' Flag for other player's data retrieval
    ' Fits: '
Private bln_Fit_Established As Boolean ' Fit is known to bidder.
Private bln_Fit_Confirmed As Boolean  ' Fit has been shown by bidder.
Private int_Fit_Suit(3) As Integer
Private int_Fit_Len(3) As Integer
Private int_Opps_FitSuit(2) As Integer
    ' My hand: '
Private bln_I_Bid_NT As Boolean
Private int_My_CueBid As Integer
Private int_My_Fit(2) As Integer ' Suit(s) in which I raised partner
Private int_My_HelpSuit As Integer
Private int_My_Pos As Integer
Private int_My_RebidSuit(2) As Integer
Private int_My_SS(3) As Integer
Private int_My_Suit(4) As Integer
    ' Pard: '
Private bln_Pard_Bid_NT As Boolean
Private int_Pards_CueBid As Integer ' Pard has shown A/K in suit 1-4 (C-S)
Private int_Pards_Fit(2) As Integer
Private int_Pards_HelpSuit As Integer
Private int_Pards_Pos As Integer ' Partner's position
Private int_Pards_RebidSuit(2) As Integer
Private int_Pards_SS As Integer ' Pard's short suit number
Private int_Pards_Suit(4) As Integer ' C=1 ... S=4
Private int_Pards_SuitLen(4) As Integer  ' The expected suit len.        '
Private str_Pards_Strength As String
    ' LHO: '
Private bln_LHO_Bid_NT As Boolean
Private int_LHOs_CueBid As Integer
Private int_LHOs_Fit(2) As Integer
Private int_LHOs_HelpSuit As Integer
Private int_LHOs_Pos As Integer
Private int_LHOs_RebidSuit(2) As Integer
Private int_LHOs_SS As Integer
Private int_LHOs_Suit(2) As Integer
Private int_LHOs_SuitLen(4) As Integer
Private str_LHOs_Strength As String
    ' RHO: '
Private bln_RHO_Bid_NT As Boolean
Private int_RHOs_CueBid As Integer
Private int_RHOs_Fit(2) As Integer
Private int_RHOs_HelpSuit As Integer
Private int_RHOs_Pos As Integer
Private int_RHOs_RebidSuit(2) As Integer
Private int_RHOs_SS As Integer ' short suit
Private int_RHOs_Suit(2) As Integer
Private int_RHOs_SuitLen(4) As Integer
Private str_RHOs_Strength As String
' Hand analysis variables: '
' (In alphabetical order,  '
'  ignoring prefixes.)     '
Private int_BidNum As Integer   ' Current bid #, starting with first non-pass
Private int_BidRound As Integer ' Bidding round, starting with 1st seat
Private str_Bid(50) As String
Private str_Call(50) As String
Private int_CallNum As Integer  ' Current call #, starting with 1st seat
Private str_Disclosure As String  ' Give client all possible bid meanings
Private int_DistribPts(4) As Integer  ' Distribution points by suit/hand
Private int_HCP(4) As Integer   ' High Card Points by suit/hand
Private lng_ID As Long  'ID# of entry matching hand's specs
Private int_NumLosers(4) As Integer ' Losing Trick Count by suit/hand
Private sng_NumRunners(4) As Single ' # of running winners
Private sng_NumStops(4) As Single ' # of stoppers by suit, # of suits stopped
Private sng_NumWinners(4) As Single ' # of winners by suit(max=4)/hand
Private str_Position As String  ' Calculated position codes
Private int_PosNum As Integer   ' Position # not incl. opening passes.
Private str_PriorBids As String ' Prior bids/calls from client
Private sng_QuickTricks(4) As Single ' # of Quick Tricks by suit/hand
Private str_Shape As String     ' Shape code for this hand
Private str_Shortness As String ' Code if singleton, void, or both
Private str_Suit(4) As String   ' Card string by suit (e.g.: AJ732)
Private bln_Suit(4, 14) As Boolean ' Cards held (T/F) array for hand
Private int_SuitLen(4) As Integer  ' # of cards in each suit
Private str_SuitNames(4) As String ' Clubs, Diamonds, Hearts, Spades
' BidBase Hand Valuation variables:                   '
    ' Specs: '
Private str_CardsSpec As String
Private str_CardsSpecCombo As String
Private str_CardsSpecExtracted As String
Private str_CardsSpecOperator As String
Private str_FitSpec As String
Private str_HandShape As String
Private str_HandSpec As String
Private int_LenSpec As Integer
Private str_LenSpec As String
Private str_OPsSpec As String
Private str_PosSpec As String
Private int_SuitSpec As Integer
Private str_SuitSpec As String
Private sng_ValSpec As Single
Private str_ValSpec As String
    ' Flags: '
Private bln_CountFlag As Boolean
Private str_Flag(20) As String
Private int_Flag(20) As Integer
Private int_FlagNum As Integer
Private str_FlagSpec As String
Private int_NumFlags As Integer
    ' Hand Value: '
Private sng_HandVal As Single
Private sng_HandNT_Val As Single
Private sng_HandSuit_Val As Single
Private int_HCPn(4) As Integer
Private int_Intermediates(4) As Integer
Private int_Shpn(4) As Integer
Private int_SuitNum As Integer
Private bln_ValAdded As Boolean
Private str_ValSystem(20) As String ' Names of hand valuation systems.'
Private sng_ValPoints(20, 2) As Single ' 1=suit, 2=NT '
' Grid Columns: '
Private int_Col_VP As Integer
Private int_Col_Len As Integer
Private int_Col_Cards As Integer
Private int_Col_Suit As Integer
Private int_Col_Fit As Integer
Private int_Col_Hands As Integer
Private int_Col_Pos As Integer
Private int_Col_Flag As Integer
Private int_Col_Value As Integer
Private int_Col_NTSuit As Integer

Private Function LowWord(dwValue As Long) As Integer
   If dwValue < &H8000& Then
      LowWord = CInt(dwValue)
   Else
      LowWord = CInt(dwValue - &H10000)
   End If
End Function

Public Sub ComputeBBVal()
    Dim s As String
    Dim i As Integer
    Dim n As Single
    Dim a As String

    AnalyzeHand
    Eval_A_Hand "BidBase"
    If Len("" & sng_HandVal) > 5 Then
        sng_HandVal = CInt(sng_HandVal * 100) / 100
    End If
    MsgBox "Stop ComputeBBVal 1"
    '(sng_HandVal + sng_HandSuit_Val)
    '(sng_HandVal + sng_HandNT_Val)
End Sub


'*****************************************************************
' Name:  Eval_A_Hand                                             '
' Purpose: Calculate custom hand Valuation Points based on specs '
'          in the HandVal table in BidBase.                      '
' See the programmer docs for details and explanations.          '
' All Subs used by this Sub start with "Eval_" to distinguish    '
' them from Subs used by the LookUpBid sub.                      '
'*****************************************************************

Public Sub Eval_A_Hand(sysName As String)
    Dim n As Single
    Dim m As String
    Dim i As Integer
    Dim j As Integer
    Dim row As Integer
    Dim col As Integer
    Dim startVal As Single
    Dim startNT As Single
    Dim lNewVal As Long
    Dim recNum As Long

    Eval_Initialize sysName

    Do While DBGrid1.Text <> ""
        ' See programmer docs. '

        Eval_SetUpFieldVariables
        recNum = recNum + 1
        
        If Eval_ValSpecOk Then
            If Eval_FlagSpecOk Then
                If Eval_PosSpecOk Then
                    If Eval_HandSpecOk Then
                        If Eval_FitSpecOk Then
                            If Eval_SuitSpecOk Then
                                Eval_CardsSpecs
                            End If
                        End If
                    End If
                End If
            End If
        End If

        int_SuitNum = 0

        ' Display entry's results, if any, in grid: '
        With DBGrid1
        '          '
        If sng_HandVal + sng_HandSuit_Val <> startVal Then
            n = sng_HandVal + sng_HandSuit_Val
            m = "" & (n - startVal)
            If Len(m) > 5 Then
                m = Left$(m, 5)
            End If
            Select Case Len(m)
            Case Is = 1, 2
                m = "   " & m
            Case Is = 3
                m = "  " & m
            Case Is = 4
                m = " " & m
            End Select
            .Columns(int_Col_VP).Text = m
            startVal = n

        ElseIf sng_HandNT_Val <> startNT Then
            m = "" & (sng_HandNT_Val - startNT)
            If Len(m) > 5 Then
                m = Left$(m, 5)
            End If
            Select Case Len(m)
            Case Is = 1, 2
                m = "   " & m
            Case Is = 3
                m = "  " & m
            Case Is = 4
                m = " " & m
            End Select
            .Columns(int_Col_VP).Text = m
            startNT = sng_HandNT_Val
        End If

        If .row = 29 Then
            j = Data1.Recordset.RecordCount - recNum
            If j = 0 Then Exit Do
            MsgBox "Stop Eval_A_Hand 1" ' At first I needed the below, then it worked without it. '
            If j > 29 Then
                .Scroll 0, 29
            ElseIf j > 0 Then
                .Scroll 0, j
            Else
                Exit Do
            End If
            .row = 0
        ElseIf .row = Data1.Recordset.RecordCount - 1 Then
            Exit Do
        Else
            .row = .row + 1
        End If
        
        If .Text = "" Then Exit Do

        End With
        '       '
        DoEvents
    Loop

    j = 1
    Do While str_ValSystem(j) <> sysName And j < 20
        j = j + 1
    Loop
    If str_ValSystem(j) <> sysName Then
        j = 0
    End If
    sng_ValPoints(j, 1) = sng_HandVal + sng_HandSuit_Val
    sng_ValPoints(j, 2) = sng_HandVal + sng_HandNT_Val
    For i = 1 To 2
        n = Fix(sng_ValPoints(j, i))
        If sng_ValPoints(j, i) - n >= 0.5 Then
            sng_ValPoints(j, i) = n + 0.5
        Else
            sng_ValPoints(j, i) = n
        End If
    Next
End Sub


Private Sub Eval_Initialize(sysName As String)
    sng_HandVal = 0
    sng_HandNT_Val = 0
    sng_HandSuit_Val = 0
    Erase int_Shpn()
    Erase int_HCPn()
    str_HandShape = ""
    Erase str_Flag()
    Erase int_Flag()
    int_FlagNum = 0
    int_NumFlags = 0

    'DBGrid1.row = 0
    'DBGrid1.col = 0
    Eval_CalcHandShape
    Eval_SortSuitsByHCP
    GetOtherPlayersData
End Sub


Private Sub Eval_CalcHandShape()
    ' Sort suits by length where s1(4) = longest: '
    ' (Doesn't matter here, but 1=C, 4=S.)        '
    Dim n As Integer

    If int_Shpn(1) > 0 Then ' already done '
        Exit Sub
    End If

    For n = 1 To 4
        int_Shpn(n) = n
    Next

    Do
        If int_SuitLen(int_Shpn(2)) < int_SuitLen(int_Shpn(1)) Then
            int_Shpn(0) = int_Shpn(2)
            int_Shpn(2) = int_Shpn(1)
            int_Shpn(1) = int_Shpn(0)
        ElseIf int_SuitLen(int_Shpn(3)) < int_SuitLen(int_Shpn(2)) Then
            int_Shpn(0) = int_Shpn(2)
            int_Shpn(2) = int_Shpn(3)
            int_Shpn(3) = int_Shpn(0)
        ElseIf int_SuitLen(int_Shpn(4)) < int_SuitLen(int_Shpn(3)) Then
            int_Shpn(0) = int_Shpn(4)
            int_Shpn(4) = int_Shpn(3)
            int_Shpn(3) = int_Shpn(0)
        Else
            Exit Do
        End If
    Loop
    str_HandShape = int_SuitLen(int_Shpn(4)) & _
                    int_SuitLen(int_Shpn(3)) & _
                    int_SuitLen(int_Shpn(2)) & _
                    int_SuitLen(int_Shpn(1))
End Sub


Private Sub Eval_SortSuitsByHCP()
    ' Sort suits by HCP where int_HCPn(4) = highest: '
    Dim n As Integer

    If int_HCPn(1) > 0 Then
        Exit Sub ' Alread done. '
    End If

    For n = 1 To 4
        int_HCPn(n) = n
    Next

    Do
        If int_HCP(int_HCPn(2)) < int_HCP(int_HCPn(1)) Then
            int_HCPn(0) = int_HCPn(2)
            int_HCPn(2) = int_HCPn(1)
            int_HCPn(1) = int_HCPn(0)
        ElseIf int_HCP(int_HCPn(3)) < int_HCP(int_HCPn(2)) Then
            int_HCPn(0) = int_HCPn(2)
            int_HCPn(2) = int_HCPn(3)
            int_HCPn(3) = int_HCPn(0)
        ElseIf int_HCP(int_HCPn(4)) < int_HCP(int_HCPn(3)) Then
            int_HCPn(0) = int_HCPn(4)
            int_HCPn(4) = int_HCPn(3)
            int_HCPn(3) = int_HCPn(0)
        Else
            Exit Do
        End If
    Loop
End Sub


Private Sub Eval_SetUpFieldVariables()
    With DBGrid1
        str_ValSpec = Trim$(.Columns(int_Col_Value))
        sng_ValSpec = Val(str_ValSpec)
        str_FlagSpec = Trim$(.Columns(int_Col_Flag))
        str_PosSpec = Trim$(.Columns(int_Col_Pos))
        str_HandSpec = Trim$(.Columns(int_Col_Hands))
        str_FitSpec = Trim$(.Columns(int_Col_Fit))
        str_SuitSpec = Trim$(.Columns(int_Col_Suit))
        str_CardsSpec = Trim$(.Columns(int_Col_Cards))
        str_LenSpec = Trim$(.Columns(int_Col_Len))
    End With
        
    If Left$(str_LenSpec & " ", 1) = "<" Then
        int_LenSpec = Val(Mid$(str_LenSpec, 2))
    Else
        int_LenSpec = Val(str_LenSpec)
    End If
    int_SuitNum = 0
    int_SuitSpec = 0
    str_CardsSpecCombo = ""
    bln_ValAdded = False
    DBGrid1.col = 0
End Sub


Private Function Eval_FlagSpecOk() As Boolean
    Dim f As String
    Dim fCompare As String
    Dim i As Integer
    Dim n As Integer
    Dim op As String

    ' When an entry has a Hand spec with "flag" in it,  '
    ' we first check to see if the flag's ID is already '
    ' in the str_Flag() array. If it is then set        '
    ' int_FlagNum to its array index, then Eval_AddVal  '
    ' increments the int_Flag() array counter if the    '
    ' entry matches the current hand. If the flag ID is '
    ' *not* in the array, then we add it to the array   '
    ' and set int_FlagNum to the new number. If the flag'
    ' spec has "=" in it, then we compare the number    '
    ' specified after the "=" to the number in int_Flag()

    bln_CountFlag = False

    If str_FlagSpec = "" Then
        int_FlagNum = 0
        Eval_FlagSpecOk = True
        Exit Function
    End If

    f = str_FlagSpec

    ' See if this is a flag test: '
    n = InStr(f, "=") ' e.g.: B1=0
    If n = 0 Then n = InStr(f, ">")
    If n = 0 Then n = InStr(f, "<")
    If n > 1 Then
        fCompare = Mid$(f, n + 1) ' e.g.: "0"
        op = Mid$(f, n, 1) ' e.g.: "="
        f = Left$(f, n - 1) ' e.g.: "B1"
    End If

    ' Look for flag in str_Flag() array: '
    For i = 1 To int_NumFlags
        If str_Flag(i) = f Then
            Exit For
        End If
    Next

    If str_Flag(i) = f Then
        ' Flag spec matches existing flag.'
        If fCompare <> "" Then
            If op = "=" And int_Flag(i) = Val(fCompare) Or _
               op = ">" And int_Flag(i) > Val(fCompare) Or _
               op = "<" And int_Flag(i) < Val(fCompare) _
            Then
                If str_CardsSpec = "" And _
                   str_LenSpec = "" And _
                   str_HandSpec = "" And _
                   str_PosSpec = "" And _
                   str_FitSpec = "" And _
                   str_SuitSpec = "" _
                Then ' No other specs, so add value: '
                    Eval_AddVal
                    Exit Function
                End If
            Else
                Eval_FlagSpecOk = False
                Exit Function
            End If
        End If

    Else
        ' Flag spec doesn't match existing flag.'
        If fCompare <> "" Or f = "stop" Then
            MsgBox "No previous flags were found for" & vbCr & _
                   str_FlagSpec
            Exit Function

        Else ' Set up new flag:
            str_Flag(i) = f
            int_NumFlags = int_NumFlags + 1
        End If
    End If
    If op = "" Then
        bln_CountFlag = True
    End If
    int_FlagNum = i
    Eval_FlagSpecOk = True
End Function


Private Function Eval_HandsSpecCalcOk() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim p As String
    Dim r As String
    Dim s As String
    Dim suitNum As Integer
    Dim neg As Integer
    Dim LLL(50) As String

    s = str_HandSpec
    If Left$(s, 1) = "-" Then
        neg = -1
        s = Mid$(s, 2)
    End If

    If s = "" Then
        Eval_HandsSpecCalcOk = True
        Exit Function
    End If

    If Len(s) = 1 Then
        ' This spec just means that the specified  '
        ' player has shown (bid) at least one suit.'
        Select Case s
        Case Is = "L"
            Eval_HandsSpecCalcOk = (int_LHOs_Suit(1) > 0) Or neg
        Case Is = "R"
            Eval_HandsSpecCalcOk = (int_RHOs_Suit(1) > 0) Or neg
        Case Is = "P"
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) > 0) Or neg
        Case Is = "M"
            Eval_HandsSpecCalcOk = (int_My_Suit(1) > 0) Or neg
        Case Is = "E"
            Eval_HandsSpecCalcOk = (int_LHOs_Suit(1) > 0 Or _
                                    int_RHOs_Suit(1) > 0) Or neg
        Case Else
            MsgBox "Unrecognized Hands spec: " & str_HandSpec
        End Select
        Exit Function
    End If
    
    If Len(s) = 3 Or (Len(s) = 4 And Mid$(s, 2, 2) = "<>") Then
        i = InStr("LPR", Left$(s, 1))
        k = InStr("BWIOPSVX", Right$(s, 1))
        If Len(s) = 4 And k > 0 Then
            ' No idea what LLL(I) is supposed to be. '
            Eval_HandsSpecCalcOk = (LLL(i) <> Right$(s, 1)) Or LLL(i - 1) = ""
        Else
            j = InStr("=<>", Mid$(s, 2, 1))
            If i > 0 And j > 0 And k > 0 Then
                i = i - 1
                Select Case Mid$(s, 2, 1)
                Case Is = "="
                    Eval_HandsSpecCalcOk = (LLL(i) = Right$(s, 1))
                Case Is = "-"
                    Eval_HandsSpecCalcOk = (LLL(i) <> Right$(s, 1)) Or LLL(i) = ""
                Case Is = "<"
                    Eval_HandsSpecCalcOk = (InStr("BWIOPSVX", LLL(i)) < _
                                            InStr("BWIOPSVX", Right$(s, 1))) Or LLL(i) = ""
                Case Is = ">"
                    Eval_HandsSpecCalcOk = (InStr("BWIOPSVX", LLL(i)) > _
                                            InStr("BWIOPSVX", Right$(s, 1)))
                End Select
            End If
            Exit Function
        End If
    End If
    
    If Left$(s & "  ", 3) = "HCP" Then
        k = Val(Mid$(s & "  ", 5)) ' e.g.: HCP=15+ '
        If Mid$(s & " ", 4, 1) = "=" Then
            If Right$(s, 1) = "+" Then
                Eval_HandsSpecCalcOk = (int_HCP(0) >= k)
            ElseIf InStr(s, "-") > 0 Then
                j = Val(Mid$(s, InStr(s, "-") + 1))
                Eval_HandsSpecCalcOk = (int_HCP(0) >= k And _
                                        int_HCP(0) <= j)
            Else
                Eval_HandsSpecCalcOk = (int_HCP(0) = k)
            End If
        ElseIf InStr(s & " ", 4, 1) = "<" Then
            Eval_HandsSpecCalcOk = (int_HCP(0) < k)
        End If
        Exit Function
    ElseIf Left$(s & "  ", 3) = "VPS" Then
        k = Val(Mid$(s & "  ", 5)) ' e.g.: VPS=15+ '
        If Mid$(s & " ", 4, 1) = "=" Then
            If Right$(s, 1) = "+" Then
                Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandSuit_Val) >= k)
            Else
                Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandSuit_Val) = k)
            End If
        ElseIf InStr(s & " ", 4, 1) = "<" Then
            Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandSuit_Val) < k)
        End If
        Exit Function
    ElseIf Left$(s & "  ", 3) = "VPN" Then
        k = Val(Mid$(s & "  ", 5)) ' e.g.: VPN=15+ '
        If Mid$(s & " ", 4, 1) = "=" Then
            If Right$(s, 1) = "+" Then
                Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandNT_Val) >= k)
            Else
                Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandNT_Val) = k)
            End If
        ElseIf InStr(s & " ", 4, 1) = "<" Then
            Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandNT_Val) < k)
        End If
        Exit Function
    End If

    p = Left$(s, 1)                 ' eg: L, LS, L=Spds, L=15, P, PC, PC<>L
    r = Trim$(Mid$(s & " ", 2))  ' eg:  ,  S,  =Spds,  =15,  ,  C,  C<>L

    Select Case p
    '            '
    Case Is = "L"
    ' LHO Specs: '
        If Left$(r, 1) = "=" Then
            ' e.g.: L=15 = LHO has 15+ VPs (entered by user) '
            Eval_HandsSpecCalcOk = (Val(str_LHOs_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: L<20 = LHO has <15 VPs '
            Eval_HandsSpecCalcOk = (Val(str_LHOs_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If

        If Len(s) = 2 Then ' LN, LS, LC, LR, LF -- alone. '
            Select Case r
            Case Is = "C"
                Eval_HandsSpecCalcOk = (int_LHOs_CueBid > 0) Or neg
            Case Is = "F"
                Eval_HandsSpecCalcOk = (int_LHOs_Fit(1) > 0) Or neg
            Case Is = "N"
                Eval_HandsSpecCalcOk = bln_LHO_Bid_NT Or neg
            Case Is = "R"
                Eval_HandsSpecCalcOk = (int_LHOs_RebidSuit(1) > 0) Or neg
            Case Is = "S"
                Eval_HandsSpecCalcOk = (int_LHOs_SS > 0) Or neg
            End Select
            Exit Function
        End If

        suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 3))
        If suitNum > 3 Then
            ' L=Clbs, etc. '
            suitNum = suitNum / 4
            Eval_HandsSpecCalcOk = (int_LHOs_Suit(1) = j Or _
                                    int_LHOs_Suit(2) = j)
        Else
            suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 4))
            If suitNum > 3 Then
                ' LS=Clbs, etc. '
                suitNum = suitNum / 4
                Select Case Left$(p, 2)
                Case Is = "LC"
                    Eval_HandsSpecCalcOk = (int_LHOs_CueBid = j)
                Case Is = "LF"
                    Eval_HandsSpecCalcOk = (int_LHOs_Fit(1) = j Or _
                                            int_LHOs_Fit(2) = j)
                Case Is = "LH"
                    Eval_HandsSpecCalcOk = (int_LHOs_HelpSuit = j)
                Case Is = "LR"
                    Eval_HandsSpecCalcOk = (int_LHOs_RebidSuit(1) = j Or _
                                            int_LHOs_RebidSuit(1) = j)
                Case Is = "LS"
                    Eval_HandsSpecCalcOk = (int_LHOs_SS = j)
                Case Else
                    MsgBox "Hand spec not recognized: " & s
                End Select
            End If
        End If
        
        Exit Function
                
    Case Is = "R"
    ' RHO Specs: '
        If Left$(r, 1) = "=" Then
            ' e.g.: R=15 = RHO has 15+ VPs (entered by user) '
            Eval_HandsSpecCalcOk = (Val(str_RHOs_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: R<20 = RHO has <15 VPs '
            Eval_HandsSpecCalcOk = (Val(str_RHOs_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If

        If Len(s) = 2 Then ' RN, RS, RC, RR, RF -- alone. '
            Select Case r
            Case Is = "C"
                Eval_HandsSpecCalcOk = (int_RHOs_CueBid > 0) Or neg
            Case Is = "F"
                Eval_HandsSpecCalcOk = (int_RHOs_Fit(1) > 0) Or neg
            Case Is = "N"
                Eval_HandsSpecCalcOk = bln_RHO_Bid_NT Or neg
            Case Is = "R"
                Eval_HandsSpecCalcOk = (int_RHOs_RebidSuit(1) > 0) Or neg
            Case Is = "S"
                Eval_HandsSpecCalcOk = (int_RHOs_SS > 0) Or neg
            End Select
            Exit Function
        End If

        suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 3))
        If suitNum > 3 Then
            ' R=Clbs, etc. '
            suitNum = suitNum / 4
            Eval_HandsSpecCalcOk = (int_RHOs_Suit(1) = j Or _
                                    int_RHOs_Suit(2) = j)
        Else
            suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 4))
            If suitNum > 3 Then
                ' RS=Clbs, etc. '
                suitNum = suitNum / 4
                Select Case Left$(p, 2)
                Case Is = "RC"
                    Eval_HandsSpecCalcOk = (int_RHOs_CueBid = j)
                Case Is = "RF"
                    Eval_HandsSpecCalcOk = (int_RHOs_Fit(1) = j Or _
                                            int_RHOs_Fit(2) = j)
                Case Is = "RH"
                    Eval_HandsSpecCalcOk = (int_RHOs_HelpSuit = j)
                Case Is = "LR"
                    Eval_HandsSpecCalcOk = (int_RHOs_RebidSuit(1) = j Or _
                                            int_RHOs_RebidSuit(1) = j)
                Case Is = "RS"
                    Eval_HandsSpecCalcOk = (int_RHOs_SS = j)
                Case Else
                    MsgBox "Hand spec not recognized: " & s
                End Select
            End If
        End If
        
        Exit Function
                
    Case Is = "P"
    ' Pard's specs: '
        If Left$(r, 1) = "=" Then
            ' e.g.: P=15 = Pard has 15+ VPs (entered by user) '
            Eval_HandsSpecCalcOk = (Val(str_Pards_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: P<20 = Pard has <15 VPs '
            Eval_HandsSpecCalcOk = (Val(str_Pards_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If
        
        If Len(r) = 1 Then
            Select Case r  ' PC, PF, PN, PR, PS -- alone '
            Case Is = "C"
                Eval_HandsSpecCalcOk = (int_Pards_CueBid > 0) Or neg
            Case Is = "F"
                Eval_HandsSpecCalcOk = (int_Pards_Fit(1) > 0) Or neg
            Case Is = "N"
                Eval_HandsSpecCalcOk = bln_Pard_Bid_NT Or neg
            Case Is = "R"
                Eval_HandsSpecCalcOk = (int_Pards_RebidSuit(1) > 0) Or neg
            Case Is = "S"
                Eval_HandsSpecCalcOk = (int_Pards_SS > 0) Or neg
            End Select
            Exit Function
        End If

        suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 3))
        If suitNum > 3 Then
            ' P=Clbs, etc. '
            suitNum = suitNum / 4
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = j Or _
                                    int_Pards_Suit(2) = j)
            Exit Function
        Else
            suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 4))
            If suitNum > 3 Then
                ' PS=Clbs, etc. '
                suitNum = suitNum / 4
                Select Case Left$(p, 2)
                Case Is = "PC"
                    Eval_HandsSpecCalcOk = (int_Pards_CueBid = j)
                Case Is = "PF"
                    Eval_HandsSpecCalcOk = (int_Pards_Fit(1) = j Or _
                                            int_Pards_Fit(2) = j)
                Case Is = "PH"
                    Eval_HandsSpecCalcOk = (int_Pards_HelpSuit = j)
                Case Is = "PR"
                    Eval_HandsSpecCalcOk = (int_Pards_RebidSuit(1) = j Or _
                                            int_Pards_RebidSuit(1) = j)
                Case Is = "PS"
                    Eval_HandsSpecCalcOk = (int_Pards_SS = j)
                Case Else
                    MsgBox "Hand spec not recognized: " & s
                End Select
                Exit Function
            End If
        End If

        ' Pard or Me comes first comparing to LHO/RHO '
        ' so LHO and RHO don't have these sections:   '

        Select Case s
        ' not equal: '
        Case Is = "PS<>L"
            Eval_HandsSpecCalcOk = (int_Pards_SS = 0 Or _
                            (int_Pards_SS <> int_LHOs_Suit(1) And _
                             int_Pards_SS <> int_LHOs_Suit(2)))
        Case Is = "PS<>R"
            Eval_HandsSpecCalcOk = (int_Pards_SS = 0 Or _
                            (int_Pards_SS <> int_RHOs_Suit(1) And _
                             int_Pards_SS <> int_RHOs_Suit(2)))
        Case Is = "PC<>L"
            Eval_HandsSpecCalcOk = (int_Pards_CueBid = 0 Or _
                            (int_Pards_CueBid <> int_LHOs_Suit(1) And _
                             int_Pards_CueBid <> int_LHOs_Suit(2)))
        Case Is = "PC<>R"
            Eval_HandsSpecCalcOk = (int_Pards_CueBid = 0 Or _
                            (int_Pards_CueBid <> int_RHOs_Suit(1) And _
                             int_Pards_CueBid <> int_RHOs_Suit(2)))
        Case Is = "P<>LS"
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = 0 Or _
                            (int_Pards_Suit(1) <> int_LHOs_SS And _
                             int_Pards_Suit(2) <> int_LHOs_SS))
        Case Is = "P<>RS"
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = 0 Or _
                            (int_Pards_Suit(1) <> int_RHOs_SS And _
                             int_Pards_Suit(2) <> int_RHOs_SS))
        Case Is = "P<>LC"
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = 0 Or _
                            (int_Pards_Suit(1) <> int_LHOs_CueBid And _
                             int_Pards_Suit(2) <> int_LHOs_CueBid))
        Case Is = "P<>RC"
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = 0 Or _
                            (int_Pards_Suit(1) <> int_RHOs_CueBid And _
                             int_Pards_Suit(2) <> int_RHOs_CueBid))

        ' equal: '
        Case Is = "PS=L"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_SS = int_LHOs_Suit(1) And _
                             int_Pards_SS = int_LHOs_Suit(2))
        Case Is = "PS=R"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_SS <> int_RHOs_Suit(1) And _
                             int_Pards_SS = int_RHOs_Suit(2))
        Case Is = "PC=L"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_CueBid = int_LHOs_Suit(1) And _
                             int_Pards_CueBid = int_LHOs_Suit(2))
        Case Is = "PC=R"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_CueBid = int_RHOs_Suit(1) And _
                             int_Pards_CueBid = int_RHOs_Suit(2))
        Case Is = "P=LS"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_Suit(1) = int_LHOs_SS And _
                             int_Pards_Suit(2) = int_LHOs_SS)
        Case Is = "P=RS"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_Suit(1) = int_RHOs_SS And _
                             int_Pards_Suit(2) = int_RHOs_SS)
        Case Is = "P=LC"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_Suit(1) = int_LHOs_CueBid And _
                             int_Pards_Suit(2) = int_LHOs_CueBid)
        Case Is = "P=RC"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_Suit(1) = int_RHOs_CueBid And _
                             int_Pards_Suit(2) = int_RHOs_CueBid)
        End Select
        Exit Function

    Case Is = "M"
        ' "My hand" (the hand entered by user) '
        ' Remember that just "M" has already   '
        ' been handled at the top, so from here'
        ' down is "M" and something else.      '
        ' Variable 'r' is the rest of the spec '
        ' following the "M".                   '

        If Len(r) = 1 Then ' MN, MS, MC, MR, MF -- alone. '
            Select Case r
            Case Is = "C"
                Eval_HandsSpecCalcOk = (int_My_CueBid > 0) Or neg
            Case Is = "F"
                Eval_HandsSpecCalcOk = (int_My_Fit(1) > 0) Or neg
            Case Is = "N"
                Eval_HandsSpecCalcOk = bln_I_Bid_NT Or neg
            Case Is = "R"
                Eval_HandsSpecCalcOk = (int_My_RebidSuit(1) > 0) Or neg
            Case Is = "S"
                Eval_HandsSpecCalcOk = (int_SuitLen(int_My_SS(1)) < 2) Or neg
            End Select
            Exit Function
        End If
        
        If Left$(s, 3) = "MS=" Then
            Select Case Right$(s, 1)
            Case Is = "0"
                If int_My_SS(1) > 0 Then
                    Eval_HandsSpecCalcOk = (int_SuitLen(int_My_SS(1)) = 0)
                End If
            Case Is = "1"
                For i = 1 To 3
                    If int_My_SS(i) > 0 Then
                        If int_SuitLen(int_My_SS(1)) = 1 Then
                            Eval_HandsSpecCalcOk = True
                            Exit For
                        End If
                    End If
                Next
            Case Is = "2"
                For i = 1 To 4
                    If int_SuitLen(i) = 2 Then
                        Eval_HandsSpecCalcOk = True
                        Exit For
                    End If
                Next
            Case Else
                MsgBox "Unrecognized Hand spec: " & s
            End Select
            Exit Function
        End If

        suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 3))
        If suitNum > 3 Then
            ' M=Clbs, etc. '
            suitNum = suitNum / 4
            Eval_HandsSpecCalcOk = (int_My_Suit(1) = j Or _
                                    int_My_Suit(2) = j)
            Exit Function
        Else
            suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 4))
            If suitNum > 3 Then
                ' MS=Clbs, etc. '
                suitNum = suitNum / 4
                Select Case Left$(p, 2)
                Case Is = "MC"
                    Eval_HandsSpecCalcOk = (int_My_CueBid = j)
                Case Is = "MF"
                    Eval_HandsSpecCalcOk = (int_My_Fit(1) = j Or _
                                            int_My_Fit(2) = j)
                Case Is = "MH"
                    Eval_HandsSpecCalcOk = (int_My_HelpSuit = j)
                Case Is = "MR"
                    Eval_HandsSpecCalcOk = (int_My_RebidSuit(1) = j Or _
                                            int_My_RebidSuit(1) = j)
                Case Is = "MS"
                    For i = 1 To 3
                        If int_My_SS(i) = j Then
                            Eval_HandsSpecCalcOk = True
                            Exit For
                        End If
                    Next
                Case Else
                    MsgBox "Hand spec not recognized: " & s
                End Select
                Exit Function
            End If
        End If

        ' not equal: '
        Select Case s
        Case Is = "MS<>L"
            Eval_HandsSpecCalcOk = (int_My_SS(1) = 0 Or _
               (int_My_SS(1) <> int_LHOs_Suit(1) And _
                int_My_SS(1) <> int_LHOs_Suit(2) And _
                int_My_SS(2) <> int_LHOs_Suit(1) And _
                int_My_SS(2) <> int_LHOs_Suit(2) And _
                int_My_SS(3) <> int_LHOs_Suit(1) And _
                int_My_SS(3) <> int_LHOs_Suit(2)))
            
        Case Is = "MS<>R"
            Eval_HandsSpecCalcOk = (int_My_SS(1) = 0 Or _
               (int_My_SS(1) <> int_RHOs_Suit(1) And _
                int_My_SS(1) <> int_RHOs_Suit(2) And _
                int_My_SS(2) <> int_RHOs_Suit(1) And _
                int_My_SS(2) <> int_RHOs_Suit(2) And _
                int_My_SS(3) <> int_RHOs_Suit(1) And _
                int_My_SS(3) <> int_RHOs_Suit(2)))
        
        Case Is = "MC<>L"
            Eval_HandsSpecCalcOk = (int_My_CueBid = 0 Or _
                            (int_My_CueBid <> int_LHOs_Suit(1) And _
                             int_My_CueBid <> int_LHOs_Suit(2)))
        Case Is = "MC<>R"
            Eval_HandsSpecCalcOk = (int_My_CueBid = 0 Or _
                            (int_My_CueBid <> int_RHOs_Suit(1) And _
                             int_My_CueBid <> int_RHOs_Suit(2)))
        Case Is = "M<>LS"
            Eval_HandsSpecCalcOk = (int_My_Suit(1) = 0 Or _
                            (int_My_Suit(1) <> int_LHOs_SS And _
                             int_My_Suit(2) <> int_LHOs_SS))
        Case Is = "M<>RS"
            Eval_HandsSpecCalcOk = (int_My_Suit(1) = 0 Or _
                            (int_My_Suit(1) <> int_RHOs_SS And _
                             int_My_Suit(2) <> int_RHOs_SS))
        Case Is = "M<>LC"
            Eval_HandsSpecCalcOk = (int_My_Suit(1) = 0 Or _
                            (int_My_Suit(1) <> int_LHOs_CueBid And _
                             int_My_Suit(2) <> int_LHOs_CueBid))
        Case Is = "P<>RC"
            Eval_HandsSpecCalcOk = (int_My_Suit(1) = 0 Or _
                            (int_My_Suit(1) <> int_RHOs_CueBid And _
                             int_My_Suit(2) <> int_RHOs_CueBid))

        ' equal: '
        Case Is = "MS=L"
            Eval_HandsSpecCalcOk = _
               (int_My_SS(1) = int_LHOs_Suit(1) Or _
                int_My_SS(1) = int_LHOs_Suit(2) Or _
                int_My_SS(2) = int_LHOs_Suit(1) Or _
                int_My_SS(2) = int_LHOs_Suit(2) Or _
                int_My_SS(3) = int_LHOs_Suit(1) Or _
                int_My_SS(3) = int_LHOs_Suit(2))
        
        Case Is = "MS=R"
            Eval_HandsSpecCalcOk = _
               (int_My_SS(1) = int_RHOs_Suit(1) Or _
                int_My_SS(1) = int_RHOs_Suit(2) Or _
                int_My_SS(2) = int_RHOs_Suit(1) Or _
                int_My_SS(2) = int_RHOs_Suit(2) Or _
                int_My_SS(3) = int_RHOs_Suit(1) Or _
                int_My_SS(3) = int_RHOs_Suit(2))
        
        Case Is = "MC=L"
            Eval_HandsSpecCalcOk = _
                            (int_My_CueBid = int_LHOs_Suit(1) And _
                             int_My_CueBid = int_LHOs_Suit(2))
        Case Is = "MC=R"
            Eval_HandsSpecCalcOk = _
                            (int_My_CueBid = int_RHOs_Suit(1) And _
                             int_My_CueBid = int_RHOs_Suit(2))
        Case Is = "M=LS"
            Eval_HandsSpecCalcOk = _
                            (int_My_Suit(1) = int_LHOs_SS And _
                             int_My_Suit(2) = int_LHOs_SS)
        Case Is = "M=RS"
            Eval_HandsSpecCalcOk = _
                            (int_My_Suit(1) = int_RHOs_SS And _
                             int_My_Suit(2) = int_RHOs_SS)
        Case Is = "M=LC"
            Eval_HandsSpecCalcOk = _
                            (int_My_Suit(1) = int_LHOs_CueBid And _
                             int_My_Suit(2) = int_LHOs_CueBid)
        Case Is = "M=RC"
            Eval_HandsSpecCalcOk = _
                            (int_My_Suit(1) = int_RHOs_CueBid And _
                             int_My_Suit(2) = int_RHOs_CueBid)
        End Select
    End Select
End Function


Private Function Eval_SuitSpecOk() As Boolean
    Dim s As String
    Dim s1 As String
    Dim op As String
    Dim i As Integer
    Dim ok As Boolean
    Dim a As String
    Dim n As Boolean
    Dim r As Boolean
    Dim l As Boolean
    Dim l0 As Boolean
    Dim r0 As Boolean
    Dim LLL(50) As String

    s = str_SuitSpec
    If s = "" Then
        Eval_SuitSpecOk = True
        Exit Function
    End If

    If Left$(s, 1) = "-" And Len(s) > 1 Then
        op = "-"
        s1 = Mid$(s, 2)
    Else
        s1 = s
    End If

    int_SuitSpec = 0

    If s1 = "R" Then ' RHO's suit. '
        int_SuitSpec = int_RHOs_Suit(1)

    ElseIf s1 = "L" Then ' LHO's suit. '
        int_SuitSpec = int_LHOs_Suit(1)

    ElseIf s = "E" Then
        If int_RHOs_Suit(1) > 0 Or int_LHOs_Suit(1) > 0 Then
            Eval_SuitSpecOk = True
        End If
        Exit Function

    ElseIf s = "ER" Then
        If int_RHOs_RebidSuit(1) > 0 Or int_LHOs_RebidSuit(1) > 0 Then
            Eval_SuitSpecOk = True
        End If
        Exit Function

    ElseIf s = "-M" Or s = "-P" Or s = "-F" Or s = "-E" Or s = "-ER" Then
        ' These are checked in Eval_CardsSpecOk. '
        Eval_SuitSpecOk = True
        Exit Function

    ElseIf s1 = "F" Then ' Fit suit. '
        int_SuitSpec = int_Fit_Suit(1)

    ElseIf s1 = "M" Then ' My suit. '
        int_SuitSpec = int_My_Suit(1)

    ElseIf s1 = "P" Then ' Pard's suit. '
        a = Right$(" " & LLL(4 + int_Pards_Suit(1)), 1)
        If a <> "F" Then
            ' If a="F", then pard was showing a Fit '
            ' for My suit; otherwise, it's his suit.'
            int_SuitSpec = int_Pards_Suit(1)
        End If

    ElseIf Left$(s & " ", 2) = "PC" Then
        ' Pard's cue A/K control. '
        If int_Pards_CueBid = 0 Then
            Exit Function
        End If

        If s = "PC" Then
            int_SuitSpec = int_Pards_CueBid
        ElseIf s = "PCE" Then
            a = LLL(4 + int_Pards_Suit(1))
        Else
            Exit Function
        End If

    ElseIf Left$(s & " ", 2) = "PS" Or Mid$(s & "  ", 2) = "PS" Then
        ' Set int_SuitSpec to a suit matching  '
        ' one of the following suits:          '
        ' PS = Partner's Short suit specified. '
        ' -PS = not Pard's SS                  '
        ' PSL = Pard's SS = LHO's main suit.   '
        ' PSR = likewise for RHO               '
        ' PSE = Pard's SS = either opp's suit. '
        ' -PSL = PSS <> LHO's suit.            '
        ' -PSR = likewise for RHO              '
        ' -PSE = PSS <> either opp's suit.     '
        If int_Pards_SS = 0 And Left$(s, 1) <> "-" Then
            Exit Function
        End If
        ' See if Pard's short suit LHO's suit: '
        l = (int_Pards_SS = int_LHOs_Suit(1))
        l = l Or (int_Pards_SS = int_LHOs_Suit(2))
        ' l is true if PSS = either of LHO's  '
        ' long suits, if any.                 '
        l0 = (int_LHOs_Suit(1) = 0)
        ' l0 is true if LHO has not bid a suit.

        ' See if Pard's short suit RHO's suit: '
        r = (int_Pards_SS = int_RHOs_Suit(1))
        r = r Or (int_Pards_SS = int_RHOs_Suit(2))
        ' r is true if PSS = either of RHO's  '
        ' long suits, if any.                 '
        r0 = (int_RHOs_Suit(1) = 0)
        ' r0 is true if RHO has not bid a suit.

        n = l Or r
        ' n is true if PSS = L or R           '

        ' DO THE COMPARISONS: '

        If Not (s = "PS<>L" And (l = False Or l0)) Then
            ' Spec is PSS <> LHO's long suit. '
            ' If L is true, get out:          '
            Exit Function

        ElseIf Not (s = "PS=L" And l = True) Then
            ' Spec is that PSS must = LHO. '
            Exit Function

        ElseIf Not (s = "PS<>R" And (r = False Or r0)) Then
            Exit Function

        ElseIf Not (s = "PS=R" And r = True) Then
            Exit Function

        ElseIf Not (s = "PS<>E" And (n = False Or (l0 And r0))) Then
            ' Either Pard's suit matches the opp's '
            ' or opp's don't have a suit.          '
            Exit Function

        ElseIf Not (s = "PS=E" And n = True) Then
            Exit Function

        End If

        int_SuitSpec = int_Pards_SS

    ElseIf s = "PH" Then ' Pard's Help suit. '
        int_SuitSpec = int_Pards_HelpSuit

    ElseIf s = "PN" Then ' Pard bid NT. '
        If int_BidNum Mod 4 = 3 Then
            ' We are 2 above opener = pard opened '
            If Left$(str_PriorBids, 2) = "1N" Then
                Eval_SuitSpecOk = True
                Exit Function
            End If
        End If

    ElseIf Left$(s, 1) = "A" Then
        If s = "A3" Then
            ' All HCPs are in 3 (or fewer) suits: '
            ok = (int_HCP(int_HCPn(1)) = 0)

        ElseIf s = "A2" Then
            ' All HCPs are in 2 (or fewer) suits: '
            ok = (int_HCP(int_HCPn(2)) = 0)

        ElseIf s = "A1" Then
            ' All HCPs are in 1 suit: '
            ok = (int_HCP(int_HCPn(4)) = int_HCP(0))

        Else
            MsgBox "Unrecognized Suit spec:" & vbCr & s
            Exit Function
        End If

        If ok Then
            If str_CardsSpec = "" And _
               str_LenSpec = "" _
            Then ' No other specs, so add value: '
                Eval_AddVal
            Else
                Eval_SuitSpecOk = True
            End If
        End If
        Exit Function

    ElseIf s = "ZSS" Then
        ' Short Suit has 0 HPCs: '
        Eval_SuitSpecOk = (int_HCP(int_Shpn(1)) = 0)
        Exit Function

    ElseIf s = "Clbs" Or s = "Dias" Or s = "Hrts" Or s = "Spds" Then
        int_SuitSpec = InStr("   ClbsDiasHrtsSpds", s) / 4
        Eval_SuitSpecOk = True
        Exit Function

    Else
        MsgBox "Suit spec not recognized:" & vbCr & s
    End If

    If int_SuitSpec > 0 Then
        Eval_SuitSpecOk = True

        If Left$(s, 1) = "-" Then
            int_SuitSpec = -int_SuitSpec
        End If
    End If
End Function


Private Function Eval_FitSpecOk() As Boolean
    Dim ok As Boolean
    Dim f As String
    Dim i As Integer
    Dim LLL(50) As String

    If int_Pards_Suit(1) = 0 Or str_FitSpec = "" Then
        Eval_FitSpecOk = (str_FitSpec = "" Or str_FitSpec = "N")
        Exit Function
    End If

    If Left$(str_FitSpec & " ", 1) = "Y" Then
        f = Right$(str_FitSpec, 1)
        If f = "P" Then ' Must be partner's suit:'
            If int_Pards_SuitLen(int_Pards_Suit(1)) > 3 Then
                ok = True
            Else
                Exit Function
            End If
        Else
            i = InStr(" CDHS", f) - 1
            If i > 0 Then
                ' Suit specified.      '
                ' See if it was raised '
                ' ("F"= showed a Fit.) '
                If Right$(" " & LLL(4 + i), 1) = "F" Then
                    ok = True
                End If

            ElseIf int_Fit_Suit(1) > 0 Then
                ok = True
            End If
        End If

    ElseIf str_FitSpec = "N" Then
        If int_Fit_Suit(1) = 0 Then
            ok = True
        End If

    Else
        MsgBox "Fit spec not recognized:" & vbCr & str_FitSpec
    End If

    If ok Then
        If str_CardsSpec = "" And _
           str_LenSpec = "" And _
           str_HandSpec = "" And _
           str_PosSpec = "" _
        Then ' No other specs, so add value: '
            Eval_AddVal
        Else ' Keep checking other specs: '
            Eval_FitSpecOk = True
        End If
    End If
End Function


Private Function Eval_PosSpecOk() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim p As String
    Dim s As String

    If str_PosSpec = "" Then
        Eval_PosSpecOk = True
        Exit Function
    End If

    p = str_PosSpec
    i = Val(p)
    k = Len(p)

    If k = 1 Then
        ' Just a seat number if len(p)=1, '
        ' not a range of seat numbers nor '
        ' opener specified, etc. See docs.'
        If int_PosNum = i Then
            GoTo PosAddVal
        End If
        Exit Function
    End If

    If k > 2 Then
        j = Val(Mid$(p, 3, 1))
        If int_CallNum < i Or _
           int_CallNum > j _
        Then
            Exit Function
        End If
    End If

    ' We've established that position # is correct.'
    ' Now see if other specs, if any, match:       '

    If k > 3 Then
        ' e.g.: 2-4R means we are in chair 2-4 and '
        '       RHO opened, 2-4P means pard opened,'
        '       2-4L means LHO opened.             '
        s = Mid$(p, 4, 1)
        i = InStr(" RPL", s)
        If int_BidNum <> i Then
            ' e.g.: If p="2-4R", then the bid  '
            ' number must match "i" which is 2.'
            ' If int_BidNum is not 2, then RHO '
            ' didn't open, so exit:            '
            Exit Function
        ElseIf i < 2 Then
            MsgBox "Position spec not valid:" & vbCr & _
                   str_PosSpec & vbCr & _
                   "See the help file."
            Exit Function
        End If
    End If

    ' We've established that the specified player opened.
    If k > 4 Then
        ' e.g.: 2-4R-S means RHO opened, but not in spades.'
        s = Right$(p, 1)
        If InStr(" CDHS", s) < 2 Then
            MsgBox "Position spec suit not valid:" & vbCr & _
                   str_PosSpec & vbCr & _
                   "See the help file."
            Exit Function
        End If

        If Mid$(p, 5, 1) = "-" Then
            If s = Right$(str_Bid(1), 1) Then
                ' The minus sign means that the '
                ' specified suit must NOT have  '
                ' been bid. Since it WAS bid... '
                Exit Function
            End If
       Else ' Specified suit must be opened: '
            If s <> Right$(str_Bid(1), 1) Then
                ' Specified suit not opened. '
                Exit Function
            End If
        End If
    End If

PosAddVal:
    If str_CardsSpec = "" And _
       str_LenSpec = "" And _
       str_HandSpec = "" _
    Then ' No other specs, so add value: '
        Eval_AddVal
    Else ' Keep checking other specs: '
        Eval_PosSpecOk = True
    End If

End Function


Private Function Eval_HandSpecOk() As Boolean
    If str_HandSpec = "" Then
        Eval_HandSpecOk = True
    ElseIf Eval_HandsSpecCalcOk Then
        If str_CardsSpec = "" And _
           str_LenSpec = "" And str_SuitSpec = "" _
        Then ' No other specs, so add value: '
            If int_SuitSpec <> 0 Then
                If int_SuitNum <> int_SuitSpec Then
                    Exit Function
                End If
            Else
                int_SuitNum = int_SuitSpec
            End If
            Eval_AddVal
        Else ' Keep checking other specs: '
            Eval_HandSpecOk = True
        End If
    End If
End Function


Private Sub Eval_CardsSpecs()
    Dim j As Integer
    Dim k As Integer
    Dim n As Integer
    Dim oper As String
    Dim s As String
    Dim b As Boolean

    If str_CardsSpec = "" And str_LenSpec = "" Then
        Exit Sub
    End If

    For int_SuitNum = 1 To 4

        bln_ValAdded = False
        str_CardsSpecCombo = ""
        str_CardsSpec = DBGrid1.Columns(int_Col_Cards) 'Trim$("" & rs_HandValue!Cards)

        If int_SuitSpec > 0 Then
            ' Check only the specified suit: '
            int_SuitNum = int_SuitSpec

        ElseIf str_SuitSpec = "P" Then
            ' Check only partner's suit: '
            Do While int_SuitNum <> int_Pards_Suit(1)
                int_SuitNum = int_SuitNum + 1
                If int_SuitNum = 5 Then
                    int_SuitNum = 0
                    Exit Sub
                End If
            Loop

        ElseIf str_SuitSpec = "-P" Then ' NOT pard's suit '
            If int_SuitNum = int_Pards_Suit(1) Then
                GoTo EvalCards_GetNext
            End If

        ElseIf str_SuitSpec = "-M" Then ' Not my suit '
            If int_SuitNum = int_My_Suit(1) Then
                GoTo EvalCards_GetNext
            End If

        ElseIf str_SuitSpec = "E" Then ' Either opp's suit '
            For k = 1 To 2
                b = (int_LHOs_Suit(k) = int_SuitNum)
                b = b Or (int_RHOs_Suit(k) = int_SuitNum)
                If b Then Exit For
            Next
            If Not b Then
                GoTo EvalCards_GetNext
            End If

        ElseIf str_SuitSpec = "-E" Then ' NOT either opp's suit. '
            b = (int_LHOs_Suit(1) = int_SuitNum)
            b = b Or (int_RHOs_Suit(1) = int_SuitNum)
            If b Then
                GoTo EvalCards_GetNext
            End If

        ElseIf str_SuitSpec = "ER" Then
            For k = 1 To 2
                b = (int_LHOs_RebidSuit(k) = int_SuitNum)
                b = b Or (int_RHOs_RebidSuit(k) = int_SuitNum)
                If b Then Exit For
            Next
            If Not b Then
                GoTo EvalCards_GetNext
            End If

        ElseIf str_SuitSpec = "-F" Then
            For k = 1 To 3
                If int_Fit_Suit(k) = int_SuitNum Then
                    GoTo EvalCards_GetNext
                End If
            Next

        ElseIf Left$(str_CardsSpec & "    ", 4) = "HCP:" Then
            ' Suit HCPs specified. '
            j = Val(Mid$(str_SuitSpec & " ", 5))
            If int_HCP(int_SuitNum) < j Then
                'Eval_SuitLenSpec
                'Eval_AddVal
                GoTo EvalCards_GetNext
            End If

        End If

        ' When cards are specified: '
        Do While str_CardsSpec <> ""
            ' When more than one spec, such as '
            ' "??9, T9, 98", extract one spec  '
            ' at a time and when one passes the'
            ' test, skip the rest.             '

            Eval_ExtractCardSpecs
            Eval_SuitLenSpec

            If bln_ValAdded Then
                If str_CardsSpecCombo = "" Then
                    Exit Do
                Else
                    ' crds1 started off a something like  '
                    ' -Q & -J and "-J" was split to Combo '
                    ' and Extracted was set to "-Q", which'
                    ' has passed its test (though the value
                    ' was not really added). Next will see'
                    ' the next spec will also passes.     '
                    str_CardsSpec = str_CardsSpecCombo
                    str_CardsSpecCombo = ""
                End If
            End If
        Loop

        If str_LenSpec <> "" And _
           Trim$(DBGrid1.Columns(int_Col_Cards)) = "" _
        Then
            If Right$("  " & str_ValSpec, 2) = "*n" Then
                n = int_SuitLen(int_SuitNum)
                If DBGrid1.Columns(int_Col_Fit) = "Y" Or _
                   DBGrid1.Columns(int_Col_Suit) = "F" _
                Then ' Fit col.
                    n = n + int_Pards_SuitLen(int_Pards_Suit(1))
                End If
                Select Case Mid$(str_LenSpec & " ", 2, 1)
                Case Is = ">"
                    ' Only get value for # of cards '
                    ' greater than specified number.'
                    ' Example: n>4                  '
                    j = Val(Mid$(str_LenSpec & " ", 3))
                    n = n - j
                    If n < 0 Then n = 0

                Case Is = "+"
                    If Mid$(str_LenSpec & "  ", 3, 2) = "T>" Then
                        j = int_Pards_SuitLen(int_Pards_Suit(1))
                        n = n + j
                        j = Val(Mid$(str_LenSpec & " ", 5))
                        n = n - j
                        If n < 0 Then n = 0
                    End If
                End Select

                sng_ValSpec = Val(str_ValSpec) * n
                If sng_ValSpec > 0 Then
                    Eval_AddVal
                End If
            ElseIf Right$(" " & str_LenSpec, 1) = "+" Then
                If int_SuitLen(int_SuitNum) >= int_LenSpec Then
                    Eval_AddVal
                End If

            ElseIf int_SuitLen(int_SuitNum) = int_LenSpec Then
                Eval_AddVal
            End If
        End If

        If int_SuitSpec <> 0 Then
           ' Only check the specified suit, if any.'
           Exit For
        End If
EvalCards_GetNext:
    Next
    int_SuitNum = 0
End Sub


Private Sub Eval_ExtractCardSpecs()
    ' When more than one spec, such as '
    ' "??9, T9, 98", extract one spec  '
    ' at a time and when one passes the'
    ' test, skip the rest.             '
    Dim n As Integer

    n = InStr(str_CardsSpec, ",")
    If n > 1 Then
        str_CardsSpecExtracted = Trim$(Left$(str_CardsSpec, n - 1))
        str_CardsSpec = Trim$(Mid$(str_CardsSpec & " ", n + 1))
    Else
        str_CardsSpecExtracted = Trim$(str_CardsSpec)
        str_CardsSpec = ""
    End If

    n = InStr(str_CardsSpecExtracted, "&")
    If n > 1 Then
        str_CardsSpecCombo = _
            Trim$(Mid$(str_CardsSpecExtracted & " ", n + 1))

        str_CardsSpecExtracted = _
            Trim$(Left$(str_CardsSpecExtracted, n - 1))
    End If

    str_CardsSpecOperator = _
        Trim$(Left$(str_CardsSpecExtracted & " ", 1))

    If str_CardsSpecOperator <> "" Then
        If InStr(" AKQJT98", str_CardsSpecOperator) < 2 Then
            str_CardsSpecExtracted = Mid$(str_CardsSpecExtracted, 2)
        Else
            str_CardsSpecOperator = ""
        End If
    End If

End Sub


Private Sub Eval_SuitLenSpec()
    Dim n As Integer
    Dim j As Integer
    Dim notS As Boolean
    Dim X As String

    ' Below: Exit if the following is *NOT* true:      '
    '        length of current suit = specified length '
    '        -OR-                                      '
    '        length spec ends with "+" (e.g.: 4+) AND  '
    '        length of cur.suit < specified length     '
    '        -OR-                                      '
    '        length spec starts with < (e.g.: <9) AND  '
    '        length of cur.suit >= specified length    '
    '        -OR-                                      '
    '        specified length = "" (nothing specified) '
    '        -OR-                                      '
    '        specified length = "n", which is the suit '
    '        length.                                   '
    If Left$(str_LenSpec, 1) = "<" Then
        If int_SuitLen(int_SuitNum) >= int_LenSpec Then
            Exit Sub
        End If
    ElseIf Right$(str_LenSpec, 1) = "+" Then
        If int_SuitLen(int_SuitNum) < int_LenSpec Then
            Exit Sub
        End If
    ElseIf (int_SuitLen(int_SuitNum) <> int_LenSpec And _
            str_LenSpec <> "" And _
            str_LenSpec <> "n") _
    Then
        Exit Sub
    End If

    If Right$("  " & str_ValSpec, 2) = "*n" Then
        n = int_SuitLen(int_SuitNum)

        Select Case Mid$(str_LenSpec & " ", 2, 1)
        '                                       '
        Case Is = ">"
            ' Only get value for # of cards '
            ' greater than specified number.'
            ' Example: n>4                  '
            j = Val(Mid$(str_LenSpec & " ", 3))
            n = n - j
            If n < 0 Then n = 0

        Case Is = "<"
            ' Suit len must be less than the '
            ' specified amount (e.g.:  <4 ). '
            j = Val(Mid$(str_LenSpec & " ", 2))
            If n < j Then n = 1 ' multiply Value by 1

        Case Is = "+"
            If Mid$(str_LenSpec & "  ", 3, 2) = "P>" Then
                j = int_Pards_SuitLen(int_SuitNum)
                n = n + j
                j = Val(Mid$(str_LenSpec & " ", 5))
                n = n - j
                If n < 0 Then n = 0
            End If

        End Select
        '                                         '
        sng_ValSpec = Val(str_ValSpec) * n

    End If

    If str_CardsSpecOperator = "-" Then
        X = Left$(str_CardsSpecExtracted, 1)
        If X = "?" Then
            str_CardsSpecOperator = X
            X = "-"
        End If
    Else
        X = ""
    End If

    Select Case str_CardsSpecOperator
    Case Is = "*"
        If InStr(str_Suit(int_SuitNum), _
                 str_CardsSpecExtracted) > 0 _
        Then
            Eval_AddVal
        End If

    Case Is = "?"
        j = 1
        ' Can be ?K, ??Q, ???J, so find last ?. '
        ' (Remember that one ? has already been '
        ' stripped off to Operator.)            '
        Do While Mid$(str_CardsSpecExtracted, j, 1) = "?"
            j = j + 1
        Loop

        If j > Len(str_CardsSpecExtracted) Then
            ' spec is all ?s '
            X = Mid$(str_Suit(int_SuitNum), j, 1)
            If InStr(" AKQJT", X) > 1 Then
                Eval_AddVal
            End If

        ElseIf Mid$(str_Suit(int_SuitNum), _
                j + 1, _
                Len(str_CardsSpecExtracted) - j + 1) _
           = Mid$(str_CardsSpecExtracted, j) _
        Then
            If X <> "-" Then
                Eval_AddVal
            End If

        ElseIf X = "-" Then
            Eval_AddVal
        End If

    Case Is = "-"
        If InStr(str_Suit(int_SuitNum), _
                 str_CardsSpecExtracted) = 0 _
        Then
            Eval_AddVal
        End If

    Case Is = ""
        If Left$(str_Suit(int_SuitNum), _
                 Len(str_CardsSpecExtracted)) _
            = str_CardsSpecExtracted _
        Then
            If int_SuitLen(int_SuitNum) = int_LenSpec Or _
               str_LenSpec = "" Or InStr(str_LenSpec, "n") _
            Then
                Eval_AddVal

            ElseIf Right$(str_LenSpec, 1) = "+" Then
                If int_SuitLen(int_SuitNum) >= int_LenSpec Then
                    Eval_AddVal
                End If

            ElseIf Left$(str_LenSpec, 1) = "<" Then
                If int_SuitLen(int_SuitNum) < int_LenSpec Then
                    Eval_AddVal
                End If
            End If
        End If

    Case Else
        MsgBox "Invalid suit length spec:" & vbCr & str_LenSpec

    End Select

End Sub


Private Function Eval_ValSpecOk() As Boolean
    Dim v As Single
    
    bln_ValAdded = False
    
    If str_ValSpec = "L-S" Then
        v = int_SuitLen(int_Shpn(4)) - _
            int_SuitLen(int_Shpn(1))

    ElseIf str_ValSpec = "L1+L2" Then
        v = int_SuitLen(int_Shpn(4)) + _
            int_SuitLen(int_Shpn(3))

    ElseIf str_ValSpec = "T-S" Then
        ' Trump suit minus shortest suit '
        int_SuitNum = int_Pards_Suit(1)
        v = int_SuitLen(int_SuitNum) - _
            int_SuitLen(int_Shpn(1))
    End If

    If v <> 0 Then
        sng_ValSpec = v
        If str_PosSpec = "" And _
           str_HandSpec = "" And _
           str_FitSpec = "" And _
           str_SuitSpec = "" And _
           str_CardsSpec = "" And _
           str_LenSpec = "" _
        Then
            Eval_AddVal
        Else
            Eval_ValSpecOk = True
        End If
    Else
        Eval_ValSpecOk = True
    End If
End Function


Private Sub Eval_AddVal()
    Dim X As Boolean
    Dim h As Single

    'If ("" & rs_HandValue!Value) = "" And int_FlagNum > 0 Then
    If DBGrid1.Columns(int_Col_Value) = "" And int_FlagNum > 0 Then
        ' Counting cards or distribution '
        ' but not adding to hand value.  '
        int_Flag(int_FlagNum) = int_Flag(int_FlagNum) + 1
        'bln_ValAdded = True
        Exit Sub
    ElseIf sng_ValSpec <> 0 And _
       (int_SuitSpec = int_SuitNum Or _
        int_SuitSpec = 0) _
    Then
        If str_CardsSpecCombo = "" Then
            If InStr(str_ValSpec, "*HCP") Then
                h = int_HCP(int_SuitNum)
            Else
                h = 1
            End If
            Select Case DBGrid1.Columns(int_Col_NTSuit) ' rs_HandValue!NT_or_Suit
            Case Is = "S"
                sng_HandSuit_Val = sng_HandSuit_Val + _
                                   sng_ValSpec * h
            Case Is = "N"
                sng_HandNT_Val = sng_HandNT_Val + _
                                 sng_ValSpec * h

            Case Else
                sng_HandVal = sng_HandVal + _
                              sng_ValSpec * h
            End Select
        End If
        bln_ValAdded = True
        If bln_CountFlag Then
            int_Flag(int_FlagNum) = int_Flag(int_FlagNum) + 1
        End If
    Else
        bln_ValAdded = False
    End If

End Sub


Private Sub GetOtherPlayersData()
    Dim pb As String
    Dim b As String
    Dim m As Integer
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim s As String
    Dim c As String
    Dim q(4) As Integer
    Dim z(4) As Integer
    Dim LLL(50) As String ' Don't know what this is. This sub is not being used. '
    Dim NNN(50) As Boolean ' ditto '

    Erase int_LHOs_Suit()
    Erase int_RHOs_Suit()
    Erase int_Pards_Suit()
    Erase int_LHOs_SuitLen()
    Erase int_RHOs_SuitLen()
    Erase int_Pards_SuitLen()
    Erase int_LHOs_Fit()
    Erase int_RHOs_Fit()
    Erase int_Pards_Fit()
    Erase int_RHOs_RebidSuit()
    Erase int_LHOs_RebidSuit()
    Erase int_Pards_RebidSuit()
    Erase int_My_Suit()
    Erase int_My_Fit()
    Erase int_My_SS()
    Erase int_Fit_Suit()
    Erase int_Fit_Len()
    str_LHOs_Strength = ""
    str_RHOs_Strength = ""
    str_Pards_Strength = ""
    bln_Fit_Established = False
    bln_Fit_Confirmed = False
    int_RHOs_CueBid = 0
    int_LHOs_CueBid = 0
    int_Pards_CueBid = 0
    int_RHOs_HelpSuit = 0
    int_LHOs_HelpSuit = 0
    int_Pards_HelpSuit = 0
    int_LHOs_SS = 0
    int_RHOs_SS = 0
    int_Pards_SS = 0

    ' Get LHO's suits: '
    For i = 1 To 4
        q(i) = Val(LLL(i))                '  <--- LLL = t_Len '
        int_LHOs_SuitLen(i) = q(i)
        If LLL(i) = "S" Then
            int_LHOs_SS = i
        ElseIf LLL(i) = "C" Then
            int_LHOs_CueBid = i
        ElseIf LLL(i) = "H" Then
            int_LHOs_HelpSuit = i
        ElseIf Right$(" " & LLL(i), 1) = "F" Then
            ' Raised RHO's suit. '
            If int_Opps_FitSuit(1) = 0 Then
                int_Opps_FitSuit(1) = i
                int_LHOs_Fit(1) = i
            Else
                int_Opps_FitSuit(2) = i
                int_LHOs_Fit(2) = i
            End If
        ElseIf LLL(i) = "R" Then
            If int_LHOs_RebidSuit(1) = 0 Then
                int_LHOs_RebidSuit(1) = i
            Else
                int_LHOs_RebidSuit(2) = i
            End If
        End If
    Next
    GoSub SortSuitLengths
    ' z(1) is the index number of the longest suit. '
    If q(z(1)) > 0 Then
        If q(z(1)) > 5 Then
            int_LHOs_RebidSuit(1) = z(1)
        ElseIf q(z(1)) > 3 Then
            int_LHOs_Suit(1) = z(1)
        End If
        If q(z(2)) > 3 Then
            int_LHOs_Suit(2) = z(2)
        End If
    End If
    str_LHOs_Strength = LLL(0)
    bln_LHO_Bid_NT = NNN(0)

    ' Get RHO's suits: '
    For i = 1 To 4
        j = 8 + i
        q(i) = Val(LLL(j))
        int_RHOs_SuitLen(i) = q(i)
        If LLL(j) = "S" Then
            int_RHOs_SS = i
        ElseIf LLL(j) = "C" Then
            int_RHOs_CueBid = i
        ElseIf LLL(j) = "H" Then
            int_RHOs_HelpSuit = i
        ElseIf Right$(" " & LLL(j), 1) = "F" Then
            ' Raised LHO's suit. '
            If int_Opps_FitSuit(1) = 0 Then
                int_Opps_FitSuit(1) = i
                int_RHOs_Fit(1) = i
            Else
                int_Opps_FitSuit(2) = i
                int_RHOs_Fit(2) = i
            End If
        ElseIf LLL(j) = "R" Then
            If int_RHOs_RebidSuit(1) = 0 Then
                int_RHOs_RebidSuit(1) = i
            Else
                int_RHOs_RebidSuit(2) = i
            End If
        End If
    Next
    GoSub SortSuitLengths
    If q(z(1)) > 0 Then
        If q(z(1)) > 5 Then
            int_RHOs_RebidSuit(1) = z(1)
        ElseIf q(z(1)) > 3 Then
            int_RHOs_Suit(1) = z(1)
        End If
        If q(z(2)) > 3 Then
            int_RHOs_Suit(2) = z(2)
        End If
    End If
    'str_RHOs_Strength = t_Len(2)            ' <--- LLL = t_Len '
    bln_RHO_Bid_NT = NNN(2)                 ' <--- NNN = ck_NT '

    ' Get Pard's suits: '
    For i = 1 To 4
        j = 4 + i
        q(i) = Val(LLL(j))
        int_Pards_SuitLen(i) = q(i)
        If LLL(j) = "S" Then
            int_Pards_SS = i
        ElseIf LLL(j) = "C" Then
            int_Pards_CueBid = i
        ElseIf LLL(j) = "H" Then
            int_Pards_HelpSuit = i
        ElseIf Right$(" " & LLL(j), 1) = "F" Then
            ' Showed Fit for (raised) my suit. '
            For k = 1 To 3
                If int_Fit_Suit(k) = 0 Then
                    int_Fit_Suit(k) = i
                    int_Pards_Fit(k) = i
                    Exit For
                End If
            Next
            int_My_Suit(1) = i
        ElseIf LLL(j) = "R" Then
            If int_Pards_RebidSuit(1) = 0 Then
                int_Pards_RebidSuit(1) = i
            Else
                int_Pards_RebidSuit(2) = 1
            End If
        End If
    Next
    GoSub SortSuitLengths
    If q(z(1)) > 0 Then
        If q(z(1)) > 5 Then
            int_Pards_RebidSuit(1) = z(1)
            int_Pards_Suit(1) = z(1)
        ElseIf q(z(1)) > 3 Then
            int_Pards_Suit(1) = z(1)
        End If
        If int_SuitLen(z(1)) + q(z(1)) > 7 Then
            int_Fit_Suit(1) = z(1)
        End If

        If q(z(2)) > 3 Then
            int_Pards_Suit(2) = z(2)
            If int_SuitLen(z(2)) + q(z(2)) > 7 Then
                int_Fit_Suit(2) = z(2)
            End If
        End If
    End If
    str_Pards_Strength = LLL(1)
    bln_Pard_Bid_NT = NNN(1)
    
    ' My data: '
    For i = 1 To 4
        q(i) = int_SuitLen(i)
        If LLL(12 + i) <> "" Then
            Select Case LLL(12 + i)
            Case "S"
                int_My_SS(1) = i
            Case "R"
                int_My_RebidSuit(1) = i
            Case "H"
                int_My_HelpSuit = i
            Case "F"
                int_My_Fit(1) = i
            Case Else
                If Val(LLL(12 + i)) > 3 Then
                    int_My_Suit(1) = i
                End If
            End Select
        End If
    Next
    GoSub SortSuitLengths
    If int_My_SS(1) > 0 Then
        j = 2
    Else
        j = 1
    End If
    If int_SuitLen(z(4)) < 2 Then
        int_My_SS(j) = z(4)
        If int_SuitLen(z(3)) < 2 Then
            int_My_SS(j + 1) = z(3)
            If int_SuitLen(z(2)) < 2 Then
                int_My_SS(j + 2) = z(2)
            End If
        End If
    End If
    bln_I_Bid_NT = NNN(3)
    bln_AlreadyHaveOPsData = True

Exit Sub


SortSuitLengths:
    For i = 1 To 4
        z(i) = i
    Next
    Do
        If q(z(2)) > q(z(1)) Then
            z(0) = z(2)
            z(2) = z(1)
            z(1) = z(0)
        ElseIf q(z(3)) > q(z(2)) Then
            z(0) = z(2)
            z(2) = z(3)
            z(3) = z(0)
        ElseIf q(z(4)) > q(z(3)) Then
            z(0) = z(4)
            z(4) = z(3)
            z(3) = z(0)
        Else
            Exit Do
        End If
    Loop

    ' Put higher-ranked suit first '
    ' when the same length suits:  '
    If z(2) > z(1) And q(z(2)) = q(z(1)) Then
        z(0) = z(2)
        z(2) = z(1)
        z(1) = z(0)
    End If
Return

End Sub


'                        Start hand analysis routines                          '


Private Sub AnalyzeHand()
    Calculate_BidNum
    ConvertHandStringToArrays
    ' TotalHCPs are calculated in the Sub  '
    '     ConvertHandStringToArrays, above.'
    ' DistribPts calc below in Calculate_Shape.'
    Calculate_Winners
    Calculate_LosingTrickCount
    Calculate_QuickTricks
    Calculate_Stoppers
    Calculate_Shape
    ' Intermed is calculated in the sub:   '
    '     ConvertHandStringToArray, above. '
    ' Vul. is specified by client.         '
    ' Suits: values used in Suits_Ok are   '
    ' calculated in other routines above.  '
End Sub


Private Sub Calculate_BidNum()
    Dim pb As String
    Dim pc As String
    Dim i As Integer
    Dim str_PriorCalls As String

    str_PriorCalls = ""

    If str_PriorCalls = "" Then
        int_BidNum = 1
        int_CallNum = 1
        int_BidRound = 1
        int_PosNum = 1
        Exit Sub
    End If
    'Now there must be a Prior Call but not
    'necessarily a Prior Bid (e.g.: "P").

    str_PriorBids = ""
    int_BidNum = 0
    int_CallNum = 0
    int_BidRound = 0
    int_PosNum = 0

    pc = str_PriorCalls & "     "

    ' Count opening passes for number of '
    ' Prior Calls, but set Prior Bids to '
    ' ignore opening passes:             '
    If Left$(pc, 5) = "P-P-P" Then
        str_Call(1) = "P"
        str_Call(2) = "P"
        str_Call(2) = "P"
        int_CallNum = 3
        If Mid$(pc, 6, 1) = "-" Then
            str_PriorBids = Trim$(Mid$(pc, 7))
        End If

    ElseIf Left$(pc, 3) = "P-P" Then
        str_Call(1) = "P"
        str_Call(2) = "P"
        int_CallNum = 2
        If Mid$(pc, 4, 1) = "-" Then
            str_PriorBids = Trim$(Mid$(pc, 5))
        End If

    ElseIf Left$(pc, 1) = "P" Then
        str_Call(1) = "P"
        int_CallNum = 1
        If Mid$(pc, 2, 1) = "-" Then
            str_PriorBids = Trim$(Mid$(pc, 3))
        End If

    Else
        str_PriorBids = str_PriorCalls
    End If

    pb = str_PriorBids

    'Count bids after opening passes:
    i = InStr(pb, "-")
    Do While i > 0
        int_CallNum = int_CallNum + 1
        str_Call(int_CallNum) = Left$(pb, i - 1)
        int_BidNum = int_BidNum + 1
        str_Bid(int_BidNum) = str_Call(int_CallNum)
        pb = Mid$(pb, i + 1)
        i = InStr(pb, "-")
    Loop
    If pb <> "" Then
        'Always one more bid than dashes:
        int_CallNum = int_CallNum + 1
        str_Call(int_CallNum) = Trim$(pb)
        int_BidNum = int_BidNum + 1
        str_Bid(int_BidNum) = str_Call(int_CallNum)
    End If
    ' int_CallNum & BidNum point to last '
    ' Call/Bid. Want them to point to    '
    ' the current call/bid, so add 1 more.
    int_CallNum = int_CallNum + 1
    int_BidNum = int_BidNum + 1


    ' Compute bidding round #: '
    int_BidRound = int_BidNum \ 4 'e.g.: 3\4=0, 4\4=1, 5\4=1
    If int_BidNum / 4 <> int_BidNum \ 4 Then
        ' Not divisible evenly by 4, so add 1: '
        int_BidRound = int_BidRound + 1
    End If

    ' Compute seat/position #: '
    int_PosNum = int_BidNum Mod 4
    If int_PosNum = 0 Then int_PosNum = 4

    ' Compute LHO's, Pard's and RHO's positions: '
    Select Case int_PosNum
    Case 1
        int_LHOs_Pos = 2
        int_Pards_Pos = 3
        int_RHOs_Pos = 4
    Case 2
        int_LHOs_Pos = 3
        int_Pards_Pos = 4
        int_RHOs_Pos = 1
    Case 3
        int_LHOs_Pos = 4
        int_Pards_Pos = 1
        int_RHOs_Pos = 2
    Case 4
        int_LHOs_Pos = 1
        int_Pards_Pos = 2
        int_RHOs_Pos = 3
    End Select
End Sub


Private Sub ConvertHandStringToArrays()
    '*************************************************
    ' 1. Converts hand (e.g.: AK432-Q74-K74-32) into *
    '      1a. suit array "bln_Suit(suit, card)" and *
    '      1b. suit strings "str_Suit(suit)".        *
    ' 2. Verifies accuracy of hand's card values     *
    '      (e.g.:  AJ52 not AJ25, 13 cards, etc.).   *
    ' 3. Counts High Card Points for the hand and    *
    '      for each suit, where A=4, K=3, Q=2, J=1.  *
    ' 4. Counts Intermediates (8-10).                *
    '*************************************************

    Dim strCard As String
    Dim s As Integer
    Dim intCardPtr As Integer
    Dim intCardNum As Integer
    Dim intLastCard As Integer
    Dim str_SuitName(4) As String
    'Outputs:
    Erase str_Suit() ' Substring of hand string by suit '
    Erase bln_Suit() ' T/F for each card held, by suit. '
    Erase int_HCP()  ' High card points.                '
    Erase int_SuitLen()  ' Suit length.                 '
    Erase int_Intermediates() ' # Intermed. by suit/hand'


    str_SuitName(1) = "Clubs"
    str_SuitName(2) = "Diamonds"
    str_SuitName(3) = "Hearts"
    str_SuitName(4) = "Spades"

    s = Spds  'Hands from Client start with Spades.

    intLastCard = 99 'Each card value must be < the last,
                     'so start with a high number.

    For intCardPtr = 1 To 16 'Each card (and dash) in the hand.

        Do While Mid$(str_PlayersHand, intCardPtr, 1) = "-"
            ' Suit is finished when "-" reached.          '
            ' If void in Spades, hand may start with "-". '
            ' If void in H/D, two dashes are together,    '
            ' such as AK432--K853-QJT2, so Do-Loop is used.
            If intCardPtr = 16 Then
                'If hand ends with "-", it is void in Clubs
                Exit For
            Else
                'Set variables for next suit:
                s = s - 1
                intLastCard = 99
                intCardPtr = intCardPtr + 1 'Get card after dash
            End If
        Loop

        strCard = Mid$(str_PlayersHand, intCardPtr, 1)
        intCardNum = InStr(" 23456789TJQKA", strCard)
        ' Check if card is valid: '
        If intCardNum < 2 Then
            'Character in the hand is not 2-A.
            'Return error code and message:
            MsgBox strCard & " in " & _
                    str_SuitName(s) & _
                    " in " & str_PlayersHand & _
                    " is not a valid card."
            Exit Sub

        ElseIf intCardNum >= intLastCard Then
            'If card number is greater than the
            'last card,  it is out of sequence.
            'Return error code and message:
            MsgBox _
                "The " & strCard & " of " _
                & str_SuitName(s) _
                & " is not in sequence in " _
                & Chr$(34) & str_Suit(s) _
                & strCard & Chr$(34)
            Exit Sub

        Else
            'Good card, so add to HCP if an honor,
            'add to Intermediates if 8-10,
            'and add to Suit array:
            int_HCP(s) = int_HCP(s) + _
                         InStr("JQKA", strCard)

            ' Count intermediates: '
            If InStr("9T", strCard) Then
                int_Intermediates(s) = _
                    int_Intermediates(s) + 1
            ElseIf strCard = "8" And _
                   (bln_Suit(s, 9) Or _
                    bln_Suit(s, 10)) _
            Then ' Only count 8 if also 9/10.'
                int_Intermediates(s) = _
                    int_Intermediates(s) + 1
            End If

            ' Check off card in suit array: '
            bln_Suit(s, intCardNum) = True
            ' Add card to suit string:      '
            str_Suit(s) = str_Suit(s) & strCard
            'Update LastCard number: '
            intLastCard = intCardNum
        End If
    Next

    'Total up HCPs and  '
    '  Intermediates    '
    '  for whole hand:  '
    For s = 1 To 4
        int_HCP(0) = int_HCP(0) + _
                     int_HCP(s)

        int_SuitLen(s) = Len(str_Suit(s))

        int_Intermediates(0) _
            = int_Intermediates(0) _
            + int_Intermediates(s)
    Next

End Sub


Private Sub Calculate_Winners()
    '**************************************************
    ' 1. Count number of winners in this hand.        *
    '                                                 *
    ' The number of winners in a hand is based on the *
    ' first four cards in a suit, plus one winner     *
    ' counted for each card more than four in a suit. *
    '   AKQJ=4, AKQT=3.5, AKJT=3.5, AQJT=3.5, AKQ=3   *
    '   KQJT=3, AQJ=2.5, AK=2, KQJ=2, QJTx=2, AQ=1.5  *
    '   A=1, KQ=1, QJT=1, Kx=.5                       *
    '                                                 *
    ' 2. Count number of running winners.             *
    '                                                 *
    ' Running Winners are cards which can be cashed   *
    ' without giving up the lead. AK=2, AQ=1, Kx=0    *
    ' See notes below for counting long running suits.*
    '**************************************************

    Dim s As Integer

    sng_NumWinners(0) = 0
    sng_NumRunners(0) = 0

    For s = 1 To 4
        sng_NumWinners(s) = 0
        sng_NumRunners(s) = 0

        If Left$(str_Suit(s), 4) = "AKQJ" Then
            'Additional top cards are not counted as winners
            'because they are counted in length, below.
            sng_NumWinners(s) = 4

            ' May need to specify length > 5 '
            sng_NumRunners(s) = int_SuitLen(s)

        ElseIf Left$(str_Suit(s), 3) = "AKQ" Then
            sng_NumWinners(s) = 3

            If int_SuitLen(s) > 6 Then
                ' Runs if 6 out are not 4-1-1 or '
                ' worse and then only if pard is '
                ' not the one with the 4+ and if '
                ' the Jack is in the long hand.  '
                sng_NumRunners(s) = int_SuitLen(s)
            Else
                sng_NumRunners(s) = sng_NumRunners(s) + 3
            End If

        ElseIf Left$(str_Suit(s), 2) = "AK" Then

            If int_SuitLen(s) > 7 Then ' 5 out not 3-1-1, etc.
                sng_NumRunners(s) = int_SuitLen(s)
            Else
                sng_NumRunners(s) = 2
            End If

            If Left$(str_Suit(s), 4) = "AKJT" Or _
               Left$(str_Suit(s), 4) = "AKQT" _
            Then
                sng_NumWinners(s) = sng_NumWinners(s) + 3.5
                ' Actually 3.67, since LHO will have
                ' the Q only 1/3 of the time.

            ElseIf Left$(str_Suit(s), 3) = "AKJ" Then
                sng_NumWinners(s) = sng_NumWinners(s) + 2.5

            Else
                sng_NumWinners(s) = sng_NumWinners(s) + 2

            End If

        ElseIf Left$(str_Suit(s), 1) = "A" Then

            If int_SuitLen(s) > 10 Then ' 2 out not 2-0-0
                sng_NumRunners(s) = int_SuitLen(s)
            Else
                sng_NumRunners(s) = 1
            End If

            If Left$(str_Suit(s), 4) = "AQJT" Then
                sng_NumWinners(s) = sng_NumWinners(s) + 3.5

            ElseIf Left$(str_Suit(s), 3) = "AQJ" Then
                sng_NumWinners(s) = sng_NumWinners(s) + 2.5

            ElseIf Left$(str_Suit(s), 2) = "AQ" Then
                sng_NumWinners(s) = sng_NumWinners(s) + 1.5

            ElseIf Left$(str_Suit(s), 3) = "AJT" Then
                sng_NumWinners(s) = sng_NumWinners(s) + 2
                '75% chance one finesse will win.

            Else
                sng_NumWinners(s) = sng_NumWinners(s) + 1

            End If

        ElseIf Left$(str_Suit(s), 4) = "KQJT" Then
            sng_NumWinners(s) = 3

        ElseIf Left$(str_Suit(s), 3) = "KQJ" Then
            sng_NumWinners(s) = 2

        ElseIf Left$(str_Suit(s), 2) = "KQ" Then
            sng_NumWinners(s) = 1

        ElseIf Left$(str_Suit(s), 1) = "K" And _
               int_SuitLen(s) > 1 Then
            sng_NumWinners(s) = 0.5

        ElseIf Left$(str_Suit(s), 3) = "QJT" Then
            If int_SuitLen(s) > 3 Then
                sng_NumWinners(s) = 2
            Else
                sng_NumWinners(s) = 1
            End If
        End If

        'give credit for long suits:
        If int_SuitLen(s) > 4 Then
            sng_NumWinners(s) _
                = sng_NumWinners(s) _
                + int_SuitLen(s) _
                - 4
        End If

        sng_NumWinners(0) _
            = sng_NumWinners(0) _
            + sng_NumWinners(s)

        sng_NumRunners(0) _
            = sng_NumRunners(0) _
            + sng_NumRunners(s)

    Next

End Sub


Private Sub Calculate_LosingTrickCount()
    '***************************************************
    ' Calculate "Losing Trick Count" for the hand:     *
    '     Only top 3 cards in a suit are considered,   *
    '     so a suit can have no more than 3 losers.    *
    '     In top 3 cards, cards below Q are losers:    *
    '       AKx = 1, AQx = 1, AJx = 2, KQx=1, Qxx=2    *
    '     In a doubleton, cards below K are losers:    *
    '       AQ=1, Ax=1, KQ=1, Kx=1, QJ=2, Qx=2         *
    '     In a singleton, any card but A is a loser:   *
    '       K=1, Q=1, x=1                              *
    '***************************************************
    '   The above is per www.PhillipAdlerBridge.com    *
    '***************************************************
    ' Note that "LTC" does not always consider that    *
    ' more tricks may be won with a finesse, so AQT is *
    ' considered 2 losers just like A32. Yet Kx is not *
    ' considered 2 losers. This does not seem logical, *
    ' but we stick with the standard anyway.           *
    '***************************************************

    Dim s As Integer
    Erase int_NumLosers()

    For s = 1 To 4
        If Left$(str_Suit(s), 3) = "AKQ" Then
            'No losers in top 3 cards.

        ElseIf Left$(str_Suit(s), 2) = "AK" Then
            If int_SuitLen(s) > 2 Then 'AKx.
                int_NumLosers(s) = int_NumLosers(s) + 1
            End If

        ElseIf Left$(str_Suit(s), 2) = "AQ" Then
            int_NumLosers(s) = int_NumLosers(s) + 1 'AQ only or A32-AQJ.

        ElseIf Left$(str_Suit(s), 1) = "A" Then
            If int_SuitLen(s) > 2 Then 'A32 to AJT.
                int_NumLosers(s) = int_NumLosers(s) + 2

            ElseIf int_SuitLen(s) = 2 Then 'A2 to AQ.
                int_NumLosers(s) = int_NumLosers(s) + 1

            End If

        ElseIf Left$(str_Suit(s), 3) = "KQJ" Then
            int_NumLosers(s) = int_NumLosers(s) + 1 'KQJ only.

        ElseIf Left$(str_Suit(s), 2) = "KQ" Then
            If int_SuitLen(s) > 2 Then 'KQx to KQT.
                int_NumLosers(s) = int_NumLosers(s) + 2

            ElseIf int_SuitLen(s) = 2 Then 'KQ only.
                int_NumLosers(s) = int_NumLosers(s) + 1

            End If

        ElseIf Left$(str_Suit(s), 1) = "K" Then
            If int_SuitLen(s) > 2 Then 'K32 to KJT.
                int_NumLosers(s) = int_NumLosers(s) + 2

            ElseIf int_SuitLen(s) = 2 Then 'K2 to KJ.
                int_NumLosers(s) = int_NumLosers(s) + 1

            End If

        ElseIf Left$(str_Suit(s), 1) = "Q" Then
            If int_SuitLen(s) > 2 Then 'Q32 to QJT.
                int_NumLosers(s) = int_NumLosers(s) + 2
            Else                   'Q only = 1,  Q2 to QJ = 2.
                int_NumLosers(s) = int_NumLosers(s) + _
                                  int_SuitLen(s)
            End If

        Else 'none of the top 3 honors
            If int_SuitLen(s) > 3 Then '432 to JT9
                int_NumLosers(s) = int_NumLosers(s) + 3
            Else           '2 only to K only = 1,  32 to QJ = 2.
                int_NumLosers(s) = int_NumLosers(s) + _
                                  int_SuitLen(s)
            End If
        End If

        int_NumLosers(0) = int_NumLosers(0) + _
                           int_NumLosers(s)

    Next

End Sub


Private Sub Calculate_QuickTricks()
    '********************************************
    ' Count number of Quick Tricks in the hand  *
    '      ("quick" defensive winners).         *
    '                                           *
    '   AK = 2 tricks                           *
    '   AQ = 1 1/2 tricks                       *
    '   A  = 1 trick                            *
    '   KQ = 1 trick                            *
    '   Kx = 1/2 trick                          *
    '                                           *
    ' The phrase "quick tricks" is sometimes    *
    ' used to refer to tricks which can be      *
    ' cashed without giving up the lead, which  *
    ' is clearly not the case with the above.   *
    ' What they are referring to is what we     *
    ' call "running winners" to avoid confusion *
    ' with the phrase "quick tricks" as it is   *
    ' properly used to refer to hand valuation  *
    ' that derived from methods used eons ago   *
    ' by Ely Culbertson before 4-3-2-1 counting.*
    '********************************************

    Dim s As Integer
    Dim adj As Single
    ' If a suit has 6+ cards, or if partner has '
    ' bid a suit in which this hand has length, '
    ' then it is possible that the bidder has a '
    ' singleton (or even a void) in the suit,   '
    ' so an AQ is probably worth only 1 Quick   '
    ' and AK, maybe 1.5. In fact, if your side  '
    ' is known to have a 10-card fit, these     '
    ' holdings may not even be worth 1 QT, but  '
    ' we have not addressed that yet.           '

    sng_QuickTricks(0) = 0

    For s = Clbs To Spds

        If int_SuitLen(s) > 5 Then
            adj = 0.5
        Else
            adj = 0
        End If

        If Left$(str_Suit(s), 2) = "AK" Then
            sng_QuickTricks(s) = 2 - adj

        ElseIf Left$(str_Suit(s), 2) = "AQ" Then
            sng_QuickTricks(s) = 1.5 - adj

        ElseIf Left$(str_Suit(s), 1) = "A" Then
            sng_QuickTricks(s) = 1

        ElseIf Left$(str_Suit(s), 2) = "KQ" Then
            sng_QuickTricks(s) = 1 - adj

        ElseIf Left$(str_Suit(s), 1) = "K" And _
               int_SuitLen(s) > 1 _
        Then
            sng_QuickTricks(s) = 0.5

        End If

        sng_QuickTricks(0) = sng_QuickTricks(0) + _
                             sng_QuickTricks(s)

    Next
End Sub


Private Sub Calculate_Stoppers()
    '******************************************************'
    ' The Bridge World web site defines a "Stopper" as --  '
    '   "a holding that will (or is likely to, or might)   '
    '    prevent the opponents from immediately running    '
    '    a large number of tricks in a suit at notrump."   '
    '                                                      '
    ' For example, K2 in declarer's hand is always a       '
    ' stopper, but may not be in dummy's hand. If his RHO  '
    ' bid the suit, then it is probably a stopper, but if  '
    ' dummy's LHO bid the suit, then the opening lead thru '
    ' the K2 will probably make the K fall to LHO's Ace.   '
    '                                                      '
    ' In this routine, Kx is always considered a stopper,  '
    ' no matter who bid what. (In reality, LHO may lead a  '
    ' different suit to partner who leads back through the '
    ' Kx, in which case it is NOT a stopper, but for the   '
    ' purposes of calculating stoppers, that is ignored.)  '
    '                                                      '
    ' To count as 2 stoppers, a holding in declarer's hand '
    ' must withstand the opening lead from LHO plus a      '
    ' subsequent lead by RHO. Example:  AQ2 = 2 stoppers   '
    ' because the Q will take the first trick (or the A,   '
    ' if RHO plays the K). KQ2 is not 2 stoppers because a '
    ' low card can force the Q, then RHO leads thru the K. '
    '                                                      '
    ' QT9 = no stoppers. If LHO has AJ8 and RHO has Kx and '
    ' leads thru the QT, but QT9x is a stopper after LHO   '
    ' gets the lead back. QT87 is a stopper according to ML.
    '                                                      '
    ' Q98x is 1 stopper.  If LHO leads low and RHO takes K '
    ' leads back low, you can cover with the 8, then after '
    ' LHO takes with the T, your Q9 holds off the AJ.      '
    '******************************************************'

    Dim i As Integer
    Dim s As Integer
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim pard As Integer

    sng_NumStops(0) = 0 ' NOTE: '
    ' sng_NumStops() for each suit is the total number of  '
    ' stoppers in that suit, whereas sng_NumStops(0) is    '
    ' the total number of SUITS stopped, excluding pard's. '

    ' Get partner's suit: '
    If int_BidNum > 2 Then
        If int_Pards_Suit(1) = 99 Then
            GetOtherPlayersData
            pard = int_Pards_Suit(1)
        End If
       If 1 = 2 Then
        i = int_PosNum
        If i < 3 Then i = i + 4
        s = i - 2 ' pard's position #
        If s < 1 Then s = s + 3
        If str_Call(s) = "P" Then s = s + 4

        int_Pards_Pos = s

        If Val(str_Call(s)) > 0 Then
            ' Must have bid some number of something. '
            s = InStr("CDHS", Mid$(str_Call(s), 2))
            pard = s
        End If
       End If
    End If



    For s = 1 To 4
        sng_NumStops(s) = 0

        c2 = Left$(str_Suit(s) & "  ", 2)
        c3 = Left$(str_Suit(s) & "   ", 3)
        c4 = Left$(str_Suit(s) & "    ", 4)

        If bln_Suit(s, Ace) Then
            If c2 = "AK" Or c2 = "AQ" Then
                'AK or AQ
                sng_NumStops(s) = 2 'Max=2

            ElseIf (c3 = "AJT" Or c3 = "AJ9") And _
                   int_SuitLen(s) > 3 _
            Then
                'AJTx
                sng_NumStops(s) = 2

            ElseIf c4 = "AT98" And _
                   int_SuitLen(s) > 4 _
            Then
                'AT98x
                sng_NumStops(s) = 2

            Else 'A
                sng_NumStops(s) = 1
            End If

        ElseIf bln_Suit(s, King) And _
               int_SuitLen(s) > 1 _
        Then
            If c2 = "KQ" Then
                ' KQ... '
                If c3 = "KQJ" Or c3 = "KQT" Then
                    ' KQJ... or KQT...'
                    sng_NumStops(s) = 2

                ElseIf c3 = "KQ9" And _
                       int_SuitLen(s) > 3 _
                Then
                    ' KQ9x -- J-Q, x-9-T, Ax vs Kx = 2
                    sng_NumStops(s) = 2

                Else ' KQx or less
                    If int_LHOs_Suit(1) = s Then
                        sng_NumStops(s) = 1
                    Else
                        sng_NumStops(s) = 2
                    End If
                End If

            ElseIf bln_Suit(s, Jack) Then
                ' KJ... '
                If int_SuitLen(s) > 3 Then
                    If bln_Suit(s, Ten) Then
                        'KJTx -- LHO leads to RHO's and h's cards,  '
                        '             then RHO leads (now or later) '
                        '                   LHO:   h:               '
                        '        A-x, x-T-Q, 9x vs KJ = 2 or        '
                        '        Q-K, x-T-A, 9x vs Jx = 2 or        '
                        '        x-T, x-J-Q, Ax vs Kx = 2           '
                        sng_NumStops(s) = 2

                    ElseIf bln_Suit(s, Nine) Or bln_Suit(s, 8) Then
                        ' KJ9x -- Q-K, x-9-T, Ax vs Jx               '
                        sng_NumStops(s) = 2

                    Else ' KJxx+
                        sng_NumStops(s) = 1
                    End If

                Else ' KJx or KJ
                    If int_LHOs_Suit(1) = s Or _
                       int_SuitLen(s) < 3 _
                    Then
                        sng_NumStops(s) = 1
                    Else
                        sng_NumStops(s) = 2
                    End If
                End If

            ElseIf bln_Suit(s, Ten) And int_SuitLen(s) > 3 Then
                ' KT9... or KT8...                   '
                ' LHO leads low. Say RHO plays Q and '
                ' we play K, leaving AJ9 over T8x.   '
                ' If RHO gets in and leads low again,'
                ' we play 8 and LHO plays 9. AJ takes'
                ' our last 2. Only 1 stop, but for   '
                ' this to be a problem, LHO must have'
                ' most of the high cards and length  '
                ' and RHO must get in twice. The odds'
                ' of this are sufficiently low that  '
                ' we are going to count KT9x and KT8x'
                ' as 2 stoppers. I would prefer it as'
                ' 1.5 stops, but that's not allowed. '
                If bln_Suit(s, 9) Or bln_Suit(s, 8) Then
                    sng_NumStops(s) = 2

                Else 'KTxx+
                    sng_NumStops(s) = 1

                End If

            Else 'Kx+
                sng_NumStops(s) = 1
            End If

        ElseIf bln_Suit(s, Queen) And _
               int_SuitLen(s) > 2 _
        Then
            ' Q??... '
            If bln_Suit(s, Jack) Then
                ' QJx...
                sng_NumStops(s) = 1

            ElseIf bln_Suit(s, Ten) And int_SuitLen(s) > 3 Then
                ' QTxx '
                If bln_Suit(s, 9) Then
                    ' QT9x '
                    If int_LHOs_Suit(1) = s Then
                        sng_NumStops(s) = 1
                    Else
                        sng_NumStops(s) = 2
                    End If
                End If

                If bln_Suit(s, 8) And bln_Suit(s, 7) Then
                    ' QT87 '
                    If int_LHOs_Suit(1) = s Then
                        sng_NumStops(s) = 1
                    Else
                        sng_NumStops(s) = 2
                    End If
                End If

                ' In his book Overcalls, Mike Lawrence says (on '
                ' page 102, hand 3) that QT7 is a stopper, but  '
                ' if LHO has AJ8xx and RHO has Kx, QT7 is not a '
                ' stopper, so I'm going against ML this time.   '

                ' Later: In one of his columns, Bobby Wolff says'
                ' that QT75 is a stopper. I give up:            '

                If bln_Suit(s, 7) Then
                    ' QT7x '
                    sng_NumStops(s) = 1
                End If


            ElseIf bln_Suit(s, Nine) And _
                   int_SuitLen(s) > 3 _
            Then
                ' Q9xx...
                If bln_Suit(s, Ten) Then
                    'QT9x RHO takes A, leads x-9-J, Kx vs Qx
                    sng_NumStops(s) = 1

                ElseIf bln_Suit(s, Eight) Then
                    'Q98x RHO A, leads x-8-T, KJ vs Q9
                    sng_NumStops(s) = 1
                End If

            Else ' Qxx '
                If int_LHOs_Suit(1) = s Then
                    sng_NumStops(s) = 0
                Else
                    sng_NumStops(s) = 1
                End If
            End If

        ElseIf bln_Suit(s, Jack) And _
               int_SuitLen(s) > 3 _
        Then
            ' J???... '
            If bln_Suit(s, Ten) And _
               (bln_Suit(s, Nine) Or bln_Suit(s, 8)) _
            Then
                'JT9x or JT8x  A takes x, K-T, Q-8/9
                sng_NumStops(s) = 1

            ElseIf c3 = "J98" And _
                   int_SuitLen(s) > 4 _
            Then
                'J98xx  A-x, RHO leads x-8-T, K-x, Qx vs Jx
                sng_NumStops(s) = 1

            Else
                sng_NumStops(s) = 0.5
            End If

        ElseIf (c3 = "T98" Or c3 = "T97") And _
               int_SuitLen(s) > 4 _
        Then
            'T98xx -- Ax, 2-8-J, KQx vs T9x
            sng_NumStops(s) = 1
        End If


        ' Total number of suits stopped '
        ' outside of partner's 1st suit.'
        If sng_NumStops(s) > 0 And s <> pard Then
            If sng_NumStops(s) = 0.5 Then
                sng_NumStops(0) = sng_NumStops(0) + 0.5
            Else
                sng_NumStops(0) = sng_NumStops(0) + 1
            End If
        End If

    Next

End Sub


Private Sub Calculate_Shape()
    '************************************************
    ' Shape in BidBase can be one of the following: *
    '    1 = 1-suiter (6+/3-)                       *
    '    2 = 2-suiter (6+/4+, 5+/5+)                *
    '    3 = 3-suiter (4441, 5440)                  *
    '    b = balanced (4333, 4432, 5332)            *
    '    u = unbalanced (no singleton or void)      *
    '    s = singleton                              *
    '    v = void                                   *
    '    e = either sing-void                       *
    ' But a hand can be more than one of the above, *
    ' such as "1" and "u" and "e", and we don't     *
    ' know in advance which an entry will want.     *
    '                                               *
    ' So to make comparison of the hand to BidBase  *
    ' possible, the hand's features are stored in   *
    ' two different variables.                      *
    '                                               *
    ' The first 4 are in the str_Shape variable.    *
    ' The last 3 are in the str_Shortness variable. *
    ' Unbalanced is not assigned. Instead, if entry *
    ' specifies Unbalanced, we just check to make   *
    ' sure that str_Shape is not "b".               *
    '************************************************

    Dim s As Integer
    Dim n4 As Integer  ' # of 4-card suits
    Dim n5 As Integer  ' # of 5-card suits
    Dim n6 As Integer  ' # of 6+card suits
    Dim s1 As Integer  ' singleton  '
    Dim db As Integer  ' doubleton
    Dim v As Integer   ' void       '

    Erase int_DistribPts()

    str_Shape = ""     ' 1, 2, 3, b '
    str_Shortness = "" ' s, v, d     '

    ' Count suits, shortness: '
    For s = 1 To 4
        If int_SuitLen(s) > 5 Then
            n6 = n6 + 1

        ElseIf int_SuitLen(s) = 5 Then
            n5 = n5 + 1

        ElseIf int_SuitLen(s) = 4 Then
            n4 = n4 + 1

        ElseIf int_SuitLen(s) = 2 Then
            int_DistribPts(s) = int_DistribPts(s) + 1
            db = db + 1
            
        ElseIf int_SuitLen(s) = 1 Then
            s1 = s1 + 1
            int_DistribPts(s) = int_DistribPts(s) + 2

        ElseIf int_SuitLen(s) = 0 Then
            v = v + 1
            int_DistribPts(s) = int_DistribPts(s) + 3

        End If

        int_DistribPts(0) _
            = int_DistribPts(0) _
            + int_DistribPts(s)

    Next

    ' Record num of suits to 'str_Shape':  '
    If n6 = 1 Then
        If n5 = 1 Or n4 = 1 Then
            ' 6-5 or 6-4 or more '
            str_Shape = "2"

        Else ' 6-3-2-2, 7-3-2-1, etc.
            str_Shape = "1"
        End If

    ElseIf n5 = 2 Or n6 = 2 Or (n5 = 1 And n6 = 1) Then
        ' 5-5 or 6-6+ or 5-6+ '
        str_Shape = "2"

    ElseIf n5 + n4 = 3 Then
        ' 5-4-4-0 or 4-4-4-1 '
        str_Shape = "3"

    ElseIf n6 = 0 And int_DistribPts(0) < 2 Then
        ' 5-3-3-2, 4-4-3-2, 4-3-3-3 '

        str_Shape = "b"

    End If

    ' Record shortness to 'str_Shortness':  '
    If v > 0 Then
        ' We have a void.
        If s1 > 0 Then
            str_Shortness = "b" ' both
        Else
            str_Shortness = "v"
        End If

    ElseIf s1 > 0 Then
        ' We have a singleton.'
        str_Shortness = "s"

    ElseIf db > 0 Then
        str_Shortness = "d"

    End If

End Sub




