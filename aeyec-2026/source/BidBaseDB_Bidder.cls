VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bidder"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************
'          CardShark BidBase Bidding Module            *
'         Nelson Ford  www.aeyec.com/BidBase           *
'                                                      *
' Purpose: Code for a DLL which will make a bid for a  *
'          client-specified bridge hand by comparing   *
'          the hand's features to those specified for  *
'          each bid in the database for the same set   *
'          of prior calls. See the documentation for   *
'          a more thorough explanation, including the  *
'          specific format for each input/output.      *
'                                                      *
'          Inputs from client:                         *
'               1. Path and file name of database.     *
'               2. A bridge hand.                      *
'               3. Prior calls, if any.                *
'               4. Bid to show analysis for, if any.   *
'               5. Vulnerability                       *
'               6. Scoring method ("MPs" or "IMPs")    *
'               7. Test ID (adding/editing entries)    *
'               8. Cancel search                       *
'               9. "Get Disclosures" flag.             *
'              10. Trump suit (for Blackwood, etc.)    *
'                                                      *
'          Return to client:                           *
'               1. Bid found, if any.                  *
'               2. ID code for entry found.            *
'               3. Disclosures for bid.                *
'               4. Error message, if needed            *
'               5. Error code, if needed               *
'               6. Hand data (for editor; see docs)    *
' Project - References:                                '
'   Microsoft Office 14.0 Access Database Engine Lib.  '
' Project - Components - none.                         '
'                                                      * I've been reading how blue strains the eyes. '
' Notes:  Written in VB6                               * Sadly VB6 doesn't offer a muted yellow/brown.'
'           In VB6 under Tools, Options, Editor Format * So I'm using a black background but still    '
'           I use the following screen colors:         * using dark cyan background for comments. '
'               Normal Text - black on cyan            * cyan on black      '
'               Selection Text - cyan on dark blue     * dark blue on cyan  '
'               Syntax Error - red on white            *
'               Execution Point - black on yellow      *
'               Breakpoint - white on red              *
'               Comment Text - purple on lt. green     * white on dark cyan '
'               Keyword Text - drk. green on cyan      * green on black     '
'               Identifier Text - drk. blue on cyan    * yellow on black    '
'         You can set colors to your own taste, but I  *
'         STRONGLY recommend something similar to make *
'         it easier to spot various types of code, etc.*
'*******************************************************

Option Explicit ' All variables must be explicitly declared. '
Option Compare Binary ' In string comparisons, "A" <> "a".   '
Private WithEvents frm As frmBidBaseAnalysis
Attribute frm.VB_VarHelpID = -1

Private Const Version$ = "2023.12.27"
' 2023.12.27 Added "<A" and "<K" for no ace and no A or K. '
' 2022-10-24: Working on 1M-P testing. '

Private lng_HandVal_Left As Long
Private lng_HandVal_Top As Long
Private str_Looking4Majors As String
Private str_Looking4MajorsSuit As String
Private int_Looking4Minors As Integer
Private bln_ShowBefore As Boolean
Private bln_ShowAfter As Boolean
Private bln_ShowBySuit As Boolean
Private bln_ShowNever As Boolean
Private int_ShowDeactivated As Integer
Private int_DontShowDeactivated As Integer
Private int_OptByStrainAvg(4, 5, 2) As Integer ' See Sub GetDDA '
Private int_OptByStrainOrig(4, 5, 2) As Integer ' See Sub GetDDA '
'Client input variables:
Private int_OptBySuitAvg(4, 5, 2) As Integer ' 4=player, 5=suit, 2=0-2 no, 11, 101 DDA avg. '
Private int_OptBySuitOrig(4, 5, 2) As Integer ' 4=player, 5=suit, 2=0-2 no, 11, 101 DDA avg. '
Private bln_AlreadySearching As Boolean
Private str_AnalyzeBid As String 'Bid(s) to show analysis for; e.g.: 1S
Private bln_Cancel As Integer   ' Emergency exit from this module
Private bln_CheckAllPcts As Boolean
Private str_DBPathAndName As String
Private bln_GetDisclosure As Boolean  ' Flag: Client wants disclosures?
Private str_Job As String ' Client says to save an entry. '
Private int_OptAvg As Integer ' Optimum contract using averaging - used for computing bid. '
Private int_OptOrig As Integer ' Optimum contract on original - used for computing bid. '
Private int_Opt4PardAvg As Integer ' Optimum for partner. '
Private int_Opt4PardOrig As Integer ' Optimum for partner. '
Private int_Opt4OppAvg As Integer ' Optimum for opponents. '
Private int_Opt4OppOrig As Integer ' Optimum for opponents. '
Private str_PlayersHand As String
Private str_PriorCalls As String
Private str_BidToSave As String
Private str_ScoreMethod As String 'I=IMPs, M=MPs
Private bln_ShowAnalysis As Boolean
Private lng_ShowBid(1) As Long ' Position of box showing calculated bid. '
Private bln_ShowBidPosSet As Boolean
Private lng_TestID As Long 'ID# of entry that test hand should match.
Private bln_TestEntry As Boolean ' Test only the entry for lng_TestID.
Private str_TestEntryCombo As String
Private lng_FoundID As Long
Private str_Vul As String
Private lng_DealNum As Long ' For Sub SaveBidEntry. '
'Database access variables:
Private db_BidBase As Database
Private ws_BidBase As Workspace
Private rs_BidTable As Recordset
Private rs_BidNames As Recordset
Private rs_BidsQuery As Recordset
Private db_HandVal As Database
Private ws_HandVal As Workspace
Private rs_HandVal As Recordset
Private rs_HandValBidReq As Recordset
Private bln_AlreadyInitialized As Boolean
Private bln_HandValInitialized As Boolean
'Card naming constants:
Private Const Ace = 14
Private Const King = 13
Private Const Queen = 12
Private Const Jack = 11
Private Const Ten = 10
Private Const Nine = 9
Private Const Eight = 8
'Suit naming constants:
Private Const Clbs = 1
Private Const Dias = 2
Private Const Hrts = 3
Private Const Spds = 4
' Player's Holdings: '
Private bln_AlreadyHaveOPsData ' Flag for other player's data retrieval
Private str_OPsSpec As String
Private bln_Fit_Established As Boolean ' Fit is known to bidder.
Private bln_Fit_Confirmed As Boolean  ' Fit has been shown by bidder.
Private int_Fit_Suit(3) As Integer
Private int_Fit_Len(3) As Integer
Private int_Opps_FitSuit(2) As Integer
Private lng_BON As Long ' Bid Order Number of entry with hand being tested. '
Private lng_LHO As Long
Private lng_Pard As Long
Private lng_RHO As Long

    ' Flags: '
Private bln_CountFlag As Boolean
Private str_Flag(20) As String
Private int_Flag(20) As Integer
Private int_FlagNum As Integer
Private int_NumFlags As Integer
    ' My hand: '
Private bln_I_Bid_NT As Boolean
Private int_My_CueBid As Integer
Private int_My_Fit(3) As Integer
Private int_My_HelpSuit As Integer
Private int_My_RebidSuit(2) As Integer
Private int_My_SS(3) As Integer
Private int_My_Suit(4) As Integer
Private str_My_Disclosure(5, 50) As String
Private bln_WinnersOK As Boolean
Private int_MyNumAces As Integer
Private int_MyNumKings As Integer
Private bln_MyTrumpK As Boolean
Private bln_MyTrumpQ As Boolean
Private int_MyVoidSuit As Integer
Private int_PardsVoidSuit As Integer
    ' Pard: '
Private bln_Pard_Bid_NT As Boolean
Private int_Pards_CueBid As Integer
Private int_Pards_Fit(2) As Integer
Private int_Pards_HelpSuit As Integer
Private int_Pards_Pos As Integer ' Partner's position
Private int_Pards_RebidSuit(2) As Integer
Private int_Pards_SS As Integer ' Pard's short suit number
Private int_Pards_Suit(4) As Integer ' C=1...S=4 '
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
Private int_RHOs_CueBid
Private int_RHOs_Fit(2) As Integer
Private int_RHOs_HelpSuit As Integer
Private int_RHOs_Pos As Integer
Private int_RHOs_RebidSuit(2) As Integer ' Showed a rebiddable suit
Private int_RHOs_SS As Integer ' short suit
Private int_RHOs_Suit(2) As Integer
Private int_RHOs_SuitLen(4) As Integer
Private str_RHOs_Strength As String

Private sng_HandVal As Single
Private sng_HandNT_Val As Single
Private sng_HandSuit_Val As Single
Private str_ValSystem(20) As String ' Names of hand valuation systems.'
Private sng_ValPoints(20, 2) As Single ' 1=suit, 2=NT '
' Hand analysis variables: '
' (In alphabetical order,  '
'  ignoring prefixes.)     '
Private str_Activated As String
Private bln_AnalyzeBid As Boolean 'Flag: Has bid analysis been requested?
Private bln_AnalysisShowing As Boolean ' Flag: Analysis form showing
Private int_BidNum As Integer   ' Current bid #, starting with first non-pass
Private str_BlockedBid(1) As String  ' E.g.: In Drury, block later bids from bidding 2C/2D. '
Private str_BON As String       ' Bid Order Number '
Private int_BidRound As Integer ' Bidding round, starting with 1st seat
Private str_Bid(50) As String ' Bids, starting with first non-pass
Private str_Call(50) As String  ' Array of prior bids/calls including starting passes
Private int_CallNum As Integer  ' Current call #, starting with 1st seat
Private bln_ControlBidding As Boolean
Private str_Disclosure As String  ' Give client all possible bid meanings
Private int_DistribPts(4) As Integer  ' Distribution points by suit/hand
Private bln_GettingDiscl As Boolean  ' Entry found - getting disclosures
Private int_HCP(4) As Integer   ' High Card Points by hand/suit 0-4 where 0=hand
Private lng_ID As Long  'ID# of entry matching hand's specs
Private int_Intermediates(4) As Integer ' # of intermediates by suit/hand
Private int_1NT_Min As Integer ' Minimum HCPs for opening 1N.
Private int_1NT_Max As Integer ' Maximum HCPs for opening 1N.
Private int_2NT_Min As Integer ' Minimum HCPs for opening 2N.
Private int_2NT_Max As Integer ' Maximum HCPs for opening 2N.
Private int_3NT_Min As Integer
Private int_3NT_Max As Integer
Private int_NumAces As Integer ' for Blackwood '
Private int_NumKings As Integer ' ditto '
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
Private bln_Suit(4, 14) As Boolean ' Cards held (T/F) array for hand. (Remember: the Bidder only knows the cards for 1 hand.)
Private bln_SolidSuit(4) As Boolean ' See Sub CalculateWinners
Private int_SuitLen(5) As Integer  ' # of cards in each suit (5=NT to avoid errors)
Private str_SuitNames(4) As String ' Clubs, Diamonds, Hearts, Spades
Private bln_TrumpQueen As Boolean ' for RKCB '
Private bln_TrumpKing As Boolean ' ditto '
Private bln_TrumpAce As Boolean ' for Grand Slam Force [
Private int_TrumpSuit As Integer ' C=1, D=2, H=3, S=4 - For Roman KC Blackwood, etc. '
Private int_2ndAASuit As Integer ' 2nd suit sometimes used in Ace Asking. '
Private bln_xField As Boolean      ' Suit bid matches "x" specs (T/F)
Private str_xSuit As String        ' Suit bid when matching "x" specs.
Private bln_xSuitFound As Boolean
' BidBase Hand Valuation variables:                   '
Private str_CardsSpec As String
Private str_CardsSpecCombo As String
Private str_CardsSpecExtracted As String
Private str_CardsSpecOperator As String
Private int_CustPtsTotalSpec As Integer
Private int_CustPtsSpec(5) As Integer
Private str_FlagSpec As String
Private str_FitSpec As String
Private str_HandShape As String
Private str_HandSpec As String
Private int_HCPn(4) As Integer
Private int_LenSpec As Integer
Private str_LenSpec As String
Private str_PosSpec As String
Private int_Shpn(4) As Integer
Private int_SuitNum As Integer
Private int_SuitSpec As Integer
Private str_SuitSpec As String
Private sng_ValSpec As Single
Private str_ValSpec As String
Private bln_ValAdded As Boolean
' Error return variables:                            '
Public str_ErrMsg As String
Public int_ErrCode As Integer
' Error message codes:                                '
' 1: "No entries found with these Prior Bids."        '
' 2: "No matching entry found for this hand."         '
' 3: "Hand doesn't have 13 cards: "                   '
' 4: "Invalid character in hand."                     '
' 5: "Cards not in sequence."                         '
'98: "Search cancelled by user."                      '
'99: "Already searching for bid."                     '


Private Sub Class_Initialize()
    ' To avoid a time-out on initialization, '
    ' all start-up code is in the LookUpBid  '
    ' routine called when the Client asks    '
    ' for a bid. The bln_AlreadyInitialized  '
    ' flag ensures the initialization is not '
    ' repeated on subsequent calls.          '
End Sub


Private Sub Class_Terminate()
    ' Close data file: '
    If Not bln_AlreadyInitialized Then Exit Sub
    On Error Resume Next
    rs_BidNames.Close
    rs_BidTable.Close
    rs_BidsQuery.Close
    db_BidBase.Close
    ws_BidBase.Close
    On Error GoTo 0
End Sub


'                      Data from Client:                        '


Public Property Let DBPathAndName(ByVal strClientInput As String)
    ' Drive, path, and file name of BidBase file to use. '
    If Not bln_AlreadySearching Then
        str_DBPathAndName = strClientInput
    End If
End Property


Public Property Let PlayersHand(ByVal strClientInput As String)
    ' Bridge hand in the format of AK432-64-QJT3-T2  '
    ' in the suit order spades-hearts-diamonds-clubs '
    ' (although for analysis purposes, Spds=4 and    '
    ' Clbs=1 so that Spds>Hrts>Dias>Clbs).           '
    If Not bln_AlreadySearching Then
        str_PlayersHand = Trim$(strClientInput)
    End If
End Property


Public Property Let PlayersPosition(ByVal strClientInput As String)
    If Not bln_AlreadySearching Then
        str_Position = Trim$(strClientInput)
    End If
End Property


Public Property Let PriorCalls(ByVal strClientInput As String)
    ' Prior Calls, including starting Passes. '
    ' Format: "P-1C-1N-2C-D-R-2S", etc.       '
    If Not bln_AlreadySearching Then
        str_PriorCalls = UCase$(Trim$(strClientInput))
        str_PriorCalls = Replace(str_PriorCalls, "MW", "Mw") ' Minorwood '
    End If
End Property


Public Property Let SaveBid(ByVal strClientInput As String)
    str_BidToSave = Trim$(strClientInput)
End Property


Public Property Let SaveDealNum(ByVal lngClientInput As Long)
    lng_DealNum = lngClientInput
End Property


Public Property Let Vul(ByVal strClientInput As String)
    ' Vulnerability set by client:           '
    '   U(nfav.), F(av.), B(oth), N(either)  '
    '     V(ul bidder), X=not vul bidder.    '
    ' If testing an entry, vul may be blank. '
    ' If no vul given, any vul will do.      '
    If Not bln_AlreadySearching Then
        str_Vul = UCase$(Trim$(strClientInput))
    End If
End Property


Public Property Let Activated(ByVal strClientInput As String)
    ' This can be "   ", "  0", "0  ", "00 ", "  W", etc., '
    ' for Cnv, Sub, Pct.  This should only be in use when  '
    ' saving a new or changed entry in BidBase Editor.     '
    ' It is used in Function Activation_Ok. '
    If Not bln_AlreadySearching Then
        str_Activated = Trim$(strClientInput)
    End If
End Property


Public Property Let BidToAnalyze(ByVal strClientInput As String)
    ' If user is expecting a certain bid from the database '
    ' and doesn't get it, he may request that this module  '
    ' display why each entry with that bid (for the prior  '
    ' bids specified) has failed and we will show him on   '
    ' the form "frmBidBaseAnalysis".                       '
    '                                                      '
    ' The user may specify more than one bid to watch for, '
    ' and may also specify "ALL" to see analysis for all.  '
    If Not bln_AlreadySearching Then
        str_AnalyzeBid = Trim$(strClientInput)
        bln_AnalyzeBid = (str_AnalyzeBid <> "")
        ' In the main routine, LookUpBid, for each entry    '
        ' tested, if str_AnalyzeBid matches the entry's bid '
        ' or if it is "ALL" then Sub ShowAnalysis is called '
        ' to initialize the frmBidBaseAnalysis list box.    '
        ' The bln_AnalyzeBid is also set so that any test   '
        ' routine (like HCPs_Ok) that fails knows to show   '
        ' the reason in frmBidBaseAnalysis.lstAnalysis.     '
    End If
End Property


Public Property Let ScoringMethod(ByVal strClientInput As String)
    ' I = IMPs, M = MPs          '
    ' Appears in Pct_Used field. '
    If Not bln_AlreadySearching Then
        str_ScoreMethod = Trim$(UCase$(strClientInput))
        If str_ScoreMethod <> "" And _
           str_ScoreMethod <> "I" And _
           str_ScoreMethod <> "M" _
        Then
            str_ErrMsg = "Not a valid scoring method. " & _
                         "Must be blank or 'I' or 'M'."
        End If
    End If
End Property


Public Property Let TestID(ByVal lngClientInput As Long)
    ' If testing new/edited entry, client '
    ' (e.g.: BidBase Editor) should pass  '
    ' the entry's ID# to make sure the    '
    ' test hand matches only the entry    '
    ' being created/edited.  Game-playing '
    ' clients should leave TestID blank.  '
    ' If all entries in a section are     '
    ' being tested, lng_TestID should be  '
    ' set to the current entry's ID# so   '
    ' that the test hand will be used to  '
    ' check its entry's specs. '
    If Not bln_AlreadySearching Then
        lng_TestID = lngClientInput
    End If
End Property


Public Property Let TestEntry(ByVal clientInput As Boolean)
    ' Client provides an entry ID# and a hand  '
    ' and sets this flag to test the hand only '
    ' against the specified entry rather than  '
    ' searching for a bid for the hand.        '
    If Not bln_AlreadySearching Then
        bln_TestEntry = clientInput
    End If
End Property


Public Property Let CancelSearch(ByVal blnClientInput As Boolean)
    ' Normally,  the search should be too fast '
    ' to cancel, but if there is a hang-up in  '
    ' the code, this will allow an exit from   '
    ' the LookUpBid routine.                   '
    bln_Cancel = blnClientInput
End Property


Public Property Let GetDisclosures(ByVal blnClientInput As Boolean)
    ' If playing or practicing, client will '
    ' probably want disclosures, but if     '
    ' testing entries or running sims, then '
    ' getting disclosures would slow things '
    ' down needlessly.                      '
    bln_GetDisclosure = blnClientInput
End Property


Public Property Let CheckAllPcts(ByVal blnClientInput As Boolean)
    ' Normally, the Bidder ignores Pct's      '
    ' which are "x" or "0", but when an entry '
    ' is being saved in the Editor, we want   '
    ' to check all entries.                   '
    bln_CheckAllPcts = blnClientInput
End Property


Public Property Let TrumpSuit(ByVal intClientInput As Integer)
    ' The trump suit is needed for Ace Asking bids. '
    ' 1=Clubs,  2=Diamonds,  3=Hearts,  4=Spades.   '
    ' A second digit is for 6 key asking or exclusion
    ' blackwood, etc. '
    
    ' I don't think this is being used. '
    
    int_TrumpSuit = intClientInput
    If int_TrumpSuit Mod 10 = 9 Then
        bln_ControlBidding = True
    Else
        int_2ndAASuit = int_TrumpSuit Mod 10
    End If
    int_TrumpSuit = int_TrumpSuit \ 10
End Property


Public Property Let Job(ByVal strClientInput As String)
    ' Input = "save" = save a bidding entry. '
    ' This is done here because all the code '
    ' for saving entries is in the module.   '
    ' If not a save, then
    str_Job = strClientInput
    If str_Job = "Save" Then
        SaveBidEntry
    End If
End Property

' Variables for bid calculations: '

Public Property Let OptimumAvg(ByVal intClientInput As Integer)
    ' If no bid is found and Optimum contract is given, '
    ' calculate the bid to be made. See ComputeBid sub  '
    ' at the end of this module. '
    int_OptAvg = intClientInput
End Property

Public Property Let OptimumOrig(ByVal intClientInput As Integer)
    ' If both pairs bid, the DDA on the original layout '
    ' should be more accurate than averaging. '
    int_OptOrig = intClientInput
End Property

Public Property Let Opt4PardAvg(ByVal intClientInput As Integer)
    ' Used in ComputeBid sub. '
    int_Opt4PardAvg = intClientInput
End Property

Public Property Let Opt4PardOrig(ByVal intClientInput As Integer)
    ' Used in ComputeBid sub. '
    int_Opt4PardAvg = intClientInput
End Property

Public Property Let Opt4OppAvg(ByVal intClientInput As Integer)
    ' Used in ComputeBid sub. '
    int_Opt4OppAvg = intClientInput
End Property

Public Property Let Opt4OppOrig(ByVal intClientInput As Integer)
    ' Used in ComputeBid sub. '
    int_Opt4OppOrig = intClientInput
End Property

'Public Property Let


'                     Pass data back to Client                          '

Public Property Get Bid() As String
    Dim t As Double
    Dim t2 As Double
    Dim pr As String
    Dim hs As String
    Dim i As Integer
    Dim j As Integer

    ' Note: If there is an error anywhere,   '
    ' you do not get a local error. Instead, '
    ' it goes back to the client ("b = .bid")'
    ' which gives the error:  Operations is  '
    ' not supported for this type of object. '

    'If Not bln_AlreadySearching Then
        'bln_AlreadySearching = True
        int_ErrCode = 0
        str_Disclosure = ""
      
        Dim c(4) As String
        Dim pos As Integer
        
        str_ErrMsg = ""
        LookUpBid
        
        If str_AnalyzeBid <> "" Then
            int_DontShowDeactivated = frm.ck_ShowDeactivated + 1
            Unload frm
        End If
        Bid = str_Call(int_CallNum)
End Property


Private Sub SwapHS()
    ' Swaps the H & S suits in a test hand. '
    Dim i As Integer
    Dim j As Integer
    
    i = InStr(str_PlayersHand, "-")
    j = InStr(i + 1, str_PlayersHand, "-")
    If i = 0 Or j = 0 Or j < i Then
        ' Can't be, but... "
        str_ErrMsg = "Problem with the hand: " & str_PlayersHand
        int_ErrCode = 55
        Exit Sub
    End If                  ' 123456789-123456 '
                            ' J864-5-T98763-T2 '
    str_PlayersHand = Mid$(str_PlayersHand, i + 1, j - i) & _
                      Left$(str_PlayersHand, i) & _
                      Mid$(str_PlayersHand, j + 1)

End Sub


Public Property Let Combo(ByVal strClientInput As String)
    If Not bln_AlreadySearching Then
        ' This is passed when testing an entry '
        ' where M is used in Prior Bids. '
        str_TestEntryCombo = UCase$(Trim$(strClientInput))
    End If
End Property


Public Property Get id() As Long
    ' The ID of the entry for the bid, if any. '
    If Not bln_AlreadySearching Then
        If lng_ID > 0 Then
            id = lng_ID
        ElseIf int_ErrCode > 0 Then
            id = lng_FoundID
            On Error GoTo 0
        End If
    End If
End Property

Public Property Get bon() As String
    ' The Bid_Order_Number of another entry '
    ' which matches this target bid's specs.'
    If Not bln_AlreadySearching Then
        bon = rs_BidTable!Bid_Order_Number
    End If
End Property

Public Property Get Disclosure() As String
    ' Disclosure for bid made, if any. '
    If Not bln_AlreadySearching Then
        Disclosure = str_Disclosure
    End If
End Property

Public Property Get ErrMsg() As String
    ' Text error message, if any.    '
    ' Error code is in next routine. '
    ' Doesn't make sense: When bid is returned, '
    ' the next line is e = .ErrMsg and it has   '
    ' the error message in .ErrMsg, but when F8 '
    ' is pressed, it returns to this code and   '
    ' str_ErrMsg is blank, yet in the next Sub, '
    ' int_ErrCode is still there. '
    If Not bln_AlreadySearching Then
        ErrMsg = str_ErrMsg
    Else
        ErrMsg = "Already searching for bid."
    End If
    ' I commented out the following line '
    ' because I wasn't getting the msg.  '
    ' str_ErrMsg = "" ' Clear for next try. '
End Property

Public Property Get ErrCode() As Integer
    ' Error code, if any, to client. '
    If Not bln_AlreadySearching Then
        ErrCode = int_ErrCode
    Else
        ErrCode = 99
    End If
    ' Error message codes:                                '
    ' 1: "No bid found for the specified hand."           '
    ' 2: "No matching entry found for this hand's specs." '
    ' 3: "Hand doesn't have 13 cards: "                   '
    ' 4: "Invalid character in hand."                     '
    ' 5: "Cards not in sequence."                         '
    '55: "Impossible error happened."                     '
    '88: "Test hand matches earlier entry."               '
    '89: "Test hand doesn't match entry being saved."     '
    '96: "Hand's HCP don't match."                        '
    '97: "Unrecognized vulnerability specified."          '
    '98: "Search cancelled by user."                      '
    '99: "Already searching for bid."                     '
    int_ErrCode = 0 ' Clear for next input.
End Property


Public Property Get HandData(HCPs() As Variant, _
                             Qty() As Variant, _
                             DistPts() As Variant, _
                             Winners() As Variant, _
                             Losers() As Variant, _
                             QTs() As Variant, _
                             Stoppers() As Variant, _
                             Intermeds() As Variant, _
                             BidRound As Integer, _
                             posNum As Integer, _
                             valSys As String, _
                             valPts() As Variant, _
                             DontShowHandVal As Boolean _
                             ) As Variant

    ' Pass back data for the hand which was '
    ' provided by the client. You do not    '
    ' have to find a bid for the hand to do '
    ' this, just run the AnalyzeHand sub.   '

    bln_ShowNever = DontShowHandVal
    InitializeBidBaseDLL
    AnalyzeHand

    Dim i As Integer

    For i = 0 To 4
        HCPs(i) = int_HCP(i)
        Qty(i) = int_SuitLen(i)
        DistPts(i) = int_DistribPts(i)
        Winners(i) = sng_NumWinners(i)
        Losers(i) = int_NumLosers(i)
        QTs(i) = sng_QuickTricks(i)
        Stoppers(i) = sng_NumStops(i)
        Intermeds(i) = int_Intermediates(i)
    Next
    BidRound = int_BidRound
    posNum = int_PosNum
    ' Hand evaluation: '

    If valSys <> "" Then
        Eval_A_Hand valSys
        valPts(1) = sng_HandVal + sng_HandSuit_Val 'sng_ValPoints(0, 1)
        valPts(2) = sng_HandVal + sng_HandNT_Val 'sng_ValPoints(0, 2)
    End If
End Property


'                Bid Lookup Analysis Starts Here                 '


'***********************************************************
' Procedure: LookUpBid                                     *
' Purpose:                                                 *
    '  This is the main routine for finding a bid.         *
    '  It is called by the Get Bid routine, above.         *
    '                                                      *
    '  This routine retrieves each BidBase database entry  *
    '  in which the Prior Bids for an entry match the      *
    '  Prior Bids from Client. Then it checks the specs    *
    '  for the entry against the Client hand's specs.      *
    '  If they don't match, the next entries matching      *
    '  the Prior Bids are retrieved and tested until an    *
    '  entry matching the hand's specs is found or until   *
    '  there are no more entries matching the Prior Bids.  *
    '  If no bid is found, err.code & message are returned *
    '  and it's up to the Client whether or not to perform *
    '  calculations of its own to generate a bid.          *
    '  Note that since we get the "NEXT" entry each time,  *
    '  you cannot do any other searches of the database    *
    '  while running the test because that would mess up   *
    '  the pointer to the next entry.                      *
'***********************************************************

Private Sub LookUpBid()

    Dim xxx As Single
    Dim s As String
    Dim X As String
    Dim Majors As String
    Dim minors As String

    Erase str_BlockedBid()

    If int_ErrCode > 0 Then
        ' Input error from client. Abort. '
        Exit Sub
    End If
    bln_ShowNever = True

    bln_AlreadyInitialized = False
    If Not bln_AlreadyInitialized Then
        InitializeBidBaseDLL
    End If

    ' Analyze hand before testing entries: '
    AnalyzeHand
    If int_ErrCode Then
        Exit Sub
    End If

LookUpAgain:
    s = "SELECT * FROM Bids"

    If bln_TestEntry And lng_TestID > 0 Then
        ' Find entry for specified ID#: '
        s = s & " WHERE ID=" & lng_TestID
    Else
        ' Find first matching entry, if any, '
        ' for specified Prior Bids:          '
        s = s & " WHERE Prior_Bids='" & str_PriorBids & "'"
    End If

    s = s & " ORDER BY Bid_Order_Number;"
    Set rs_BidTable = db_BidBase.OpenRecordset(s, dbOpenSnapshot)
    If rs_BidTable.NoMatch Or rs_BidTable.EOF Then
        If bln_TestEntry And lng_TestID > 0 Then
            str_ErrMsg = "Entry #" & lng_TestID & " not found."
            GoTo SkipEntry
        Else
            If str_Call(1) = "P" Then
                X = " (Opening passes are ignored.)"
            Else
                X = ""
            End If
            
            ComputeOptBid
            'Stop
            Exit Sub
            MsgBox _
                "No entries found matching the prior bids: " & vbCrLf & _
                str_PriorCalls & vbCrLf & X
            
            
            On Local Error GoTo 0
            int_ErrCode = 1
            
            GoTo Handle1NOvercall
            
            'If int_OptAvg > 0 Then
            '    ComputeOptBid
            'End If
            'Exit Sub
        End If
    End If

    ' Loop through entries with matching Prior Bids. '
    ' ("ZZZZZZ" is a dividing line entry that marks  '
    '   the end of a Prior Bids section. It may not  '
    '   always be used, but it should be treated as  '
    '   a stopping point when it has been used.)     '
Dim zz1 As Long
    Do While rs_BidTable!Bid_Order_Number <> "ZZZZZZ" And _
             Trim$(rs_BidTable!Prior_Bids) = str_PriorBids
zz1 = zz1 + 1

      DoEvents
      ' See if this bid was specified by client to '
      ' be analyzed, and if so, show the details:  '
      Dim num As Long
      Dim sel As String  ' 0 if convention is deactivated. '
      Dim sel2 As String ' 0 if sub-convention is deactivated. '
      Dim nam As String ' Name of convention. '
      Dim doNotShow As Boolean

      num = Val("" & rs_BidTable!ConvID)
'If rs_BidTable!id = 1475 Then Stop ' NOTE: This is for ID#, not BON#. '
      If (str_AnalyzeBid = rs_BidTable!Bid And _
         rs_BidTable!Bid <> "") _
      Then
        bln_ShowAnalysis = True
        ShowAnalysis
        If int_DontShowDeactivated = 2 Then
            nam = ""
            LookUpBidName num, sel, sel2, nam
            If sel & sel2 & rs_BidTable!pct_Used <> "" Then
                ' The Option was chosen to not '
                ' show deactivated conventions.'
                GoTo SkipEntry
            End If
        End If
      Else
        bln_ShowAnalysis = False
      End If

      If str_AnalyzeBid = "ALL" Or _
         (InStr(str_AnalyzeBid, rs_BidTable!Bid) > 0 And _
            str_AnalyzeBid = rs_BidTable!Bid And _
            str_AnalyzeBid <> "") _
      Then
          bln_ShowAnalysis = True
          ShowAnalysis
          'frm.Show
          bln_AnalyzeBid = True

      ElseIf str_AnalyzeBid = rs_BidTable!Bid Or _
             (LCase$(Right$(rs_BidTable!Bid, 1)) = "x" And _
              Left$(str_AnalyzeBid & " ", 1) = Left$(rs_BidTable!Bid, 1)) _
      Then
          X = Right$(str_AnalyzeBid, 1) ' Suit
          If (X = "C" And InStr(LCase$(rs_BidTable!CNum), "x")) Or _
             (X = "D" And InStr(LCase$(rs_BidTable!DNum), "x")) Or _
             (X = "H" And InStr(LCase$(rs_BidTable!HNum), "x")) Or _
             (X = "S" And InStr(LCase$(rs_BidTable!SNum), "x")) _
          Then
              bln_ShowAnalysis = True
              ShowAnalysis
              frm.Show
              bln_AnalyzeBid = True
          Else
              int_DontShowDeactivated = frm.ck_ShowDeactivated + 1
              frm.Hide
              bln_AnalyzeBid = False
          End If
      End If
'If zz1 = 80 Then Stop
      ' Compare hand to entry specifications: '
      ' (Suits_Ok is last because it is the   '
      '  longest and most complicated test.)  '
'If rs_BidTable!id = 25240 Then Stop ' Then GoTo SkipEntry 'Stop  ' NOTE: This is for ID#, not BON#. '
'Debug.Print rs_BidTable!id
On Error GoTo WhistleStop ' Otherwise, errors return to calling program. '
CheckItOut:
      If Activation_Ok Then
       If TotalHCPs_Ok Then
        If CustPoints_Ok Then
         If DistribPts_Ok Then
          If Winners_Ok Then
           If Losers_Ok Then ' # in DB means <= # '
            If QuickTricks_Ok Then
             If Stoppers_Ok Then
              If Shape_Ok Then
               If Intermed_Ok Then
                If Vul_Ok Then
                 If Position_Ok Then
                  'If Eval_Pts_Ok Then   ' Val systems no longer used. '
                   If Suits_Ok Then
                       'Debug.Print rs_BidTable!id; " - OK"
                       '     PASSED ALL TESTS - MATCHING ENTRY FOUND       '
                       s = GetDataFromMatchingEntry ' including disclosures and lng_ID '
                       If str_ErrMsg = "Keep testing" Then
                           str_ErrMsg = ""
                       Else
                           If bln_ShowAnalysis Then 'Or lng_TestID > 0 Then
                               If lng_TestID <> rs_BidTable!id Then Exit Sub
                               If InStr(str_PriorCalls, "M") = 0 Then
                               
                               With frm
                                   .lstAnalysis.List(24) _
                                       = Space$(35) _
                                       & "Matching entry found."
                                   .lstAnalysis.ListIndex = 24
                                   .cmd_Cancel.Enabled = True
                                   .cmd_Continue.Enabled = True
                                   .cmd_Cancel.Caption = "Cancel"
                                   .MousePointer = 0
                                   Do While .cmd_Continue.Visible And _
                                            .cmd_Cancel.Visible
                                       DoEvents
                                   Loop
                                   If .cmd_Cancel.Caption <> "Cancel" Then
                                       str_Call(int_CallNum) = ""
                                       str_ErrMsg = "Search cancelled."
                                       int_ErrCode = 98
                                   End If
                               End With
                               End If
                           End If
                           Exit Sub
                       End If
                       '                                                   '
                   'End If
                   End If
                 End If
                End If
               End If
              End If
             End If
            End If
           End If
          End If
         End If
        End If
       End If
      End If

      xxx = DoEvents()
      'Debug.Print zz1;
      If bln_Cancel Then
          int_ErrCode = 98
          str_ErrMsg = "Search cancelled by user."
          Exit Sub
      End If

      ' In an entry fails to match the client-specified '
      ' hand in one of the above tests, then the test   '
      ' routine will disclose why in frmBidBaseAnalysis.'
      ' The user must click OK to continue, at which    '
      ' point the form is unloaded.                     '

      If bln_ShowAnalysis And (lng_TestID = rs_BidTable!id Or lng_TestID = 0) Then
          With frm
              .Visible = True
              .cmd_Cancel.Enabled = True
              .cmd_Continue.Enabled = True
              .MousePointer = 0
              If lng_TestID = rs_BidTable!id Then
                  .lstAnalysis.List(24) = _
                    Space$(30) & "NOTE: Test hand is for this entry."
              End If
              Do While .cmd_Continue.Visible
                  If .cmd_Cancel.Visible = False Then
                      If lng_TestID = rs_BidTable!id Then
                          Exit Do
                      End If
                      int_ErrCode = 98
                      str_ErrMsg = "Search cancelled by user."
                      Exit Sub
                  End If
                  DoEvents
              Loop
              .cmd_Continue.Visible = True
              If lng_TestID = rs_BidTable!id Then ' And str_Looking4Majors <> "M" Then
                  ' If entry failed on its own test hand, '
                  ' exit the bid finder whether the user  '
                  ' presses Cancel -or- Continue.         '
                  int_ErrCode = 97
                  If lng_TestID = rs_BidTable!id Then
                      str_ErrMsg = "Entry was saved, but the Test" & vbCr & _
                                   "hand didn't match entry's specs."
                  Else
                      str_ErrMsg = "Test hand didn't match entry's specs."
                  End If
                  Exit Sub
              End If

              int_DontShowDeactivated = .ck_ShowDeactivated.Value + 1
          End With
      End If
'If rs_BidTable!id = 3514 Then Stop
SkipEntry:
      If lng_TestID = rs_BidTable!id And int_ErrCode > 0 Then
          ' Entry being tested didn't match test hand. '
          ' If analysis was already turned on, this    '
          ' routine would have been exited above.      '
          ' Otherwise, turn on analysis and retest so  '
          ' that the user can see what's wrong.        '
          bln_AnalyzeBid = True
          str_AnalyzeBid = rs_BidTable!Bid

      ElseIf Not rs_BidTable.EOF Then
          ' Get next record, if any: '
          rs_BidTable.MoveNext
          If rs_BidTable.EOF Then
              Exit Do
          End If

      Else
          Exit Do
      End If
      DoEvents
    Loop

Handle1NOvercall:
    If int_OptAvg > 0 Then
        ' All 4 hands are known; no bid found. '
        ' As of Oct.2022, special treatment for 1N overcalls '
        ' has been implemented below. See the Help file. '
    Dim prNTstart As String ' Starting characters of auction. '
    Dim prNTO As String ' Temp holder of original str_PriorBids '

        prNTstart = " " & Left$(str_PriorBids & "       ", 7) & " "
                    ' e.g.: " 1H-1N-P "

        If InStr(" 1C-1N-P 1D-1N-P 1H-1N-P 1S-1N-P ", prNTstart) > 0 Then
            prNTO = str_PriorBids
            str_PriorBids = Mid$(str_PriorBids, 4)
            ' A bid was not found in the original section like 1H-1N-P '
            ' so we now look in 1N-P. If a bid is found, it will be    '
            ' processed above before getting back here; otherwise, it  '
            ' will end up in the "If int_OptAvg > 0" section but not in   '
            ' this subsection and it will go to ComputeOptBid below.   '
            GoTo LookUpAgain
            
        ElseIf prNTO <> "" Then
            ' No bid was found in 1N-P, so we put things back: '
            'str_PriorBids = prNTO
        End If
                
        ComputeOptBid
        Exit Sub
    End If

    'int_ErrCode = 2
    'str_ErrMsg = "No entry found matching this hand."
    lng_ID = 0
Exit Sub

WhistleStop:
    MsgBox Err.Description & vbCr & _
           "in BidBaseDB.LookUpBid." & vbCr & _
           "Ctrl-Break to see error."
    Resume 'Next
End Sub


Private Sub InitializeBidBaseDLL()
    If bln_AlreadyInitialized Then Exit Sub
        ' To avoid a time-out when loading the DLL,     '
        ' initialization does not take place until the  '
        ' first time the DLL is called by the Client,   '
        ' then a flag is set below to keep this routine '
        ' from executing again on subsequent calls.     '
        'ChDir App.Path ' Something was changing the default directory. '
        Set ws_BidBase = DBEngine.Workspaces(0) ' Not sure if wsp is needed.
        Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
        Set rs_BidTable = db_BidBase.OpenRecordset("Bids") ' Table in the db to use.
        Set rs_BidNames = db_BidBase.OpenRecordset("BidNames")
        Set rs_BidsQuery = db_BidBase.OpenRecordset("GridQuery")

        ' Set up suit name array: '
        str_SuitNames(1) = "Clubs"
        str_SuitNames(2) = "Diamonds"
        str_SuitNames(3) = "Hearts"
        str_SuitNames(4) = "Spades"

        Set frm = New frmBidBaseAnalysis

        bln_AlreadyInitialized = True
End Sub


'                        Start hand analysis routines                          '


Private Sub AnalyzeHand()
    '**************************************'
    ' Most features of a hand only need to '
    ' be analyzed once and then compared to'
    ' as many entries as match the client- '
    ' specified Prior Bids.                '
    '                                      '
    ' For easier tracking, routines are    '
    ' called in the order that specs are   '
    ' checked in the main sub, LookUpBid.  '
    '                                      '
    ' Where spec tests in that routine do  '
    ' not need prior computing or are not  '
    ' computed in a stand-alone routine,   '
    ' it is indicated below. For example,  '
    ' PctUsed is not calculated for the    '
    ' hand before testing entries because  '
    ' it can only be calculated when an    '
    ' entry specifies it. But Winners in   '
    ' a hand CAN be counted one time and   '
    ' used in testing multiple entries.    '
    '**************************************'

    Calculate_BidNum
    ConvertHandStringToArrays
    If int_ErrCode > 0 Then
        Exit Sub
    End If

    ' TotalHCPs are calculated in the Sub      '
    '     ConvertHandStringToArrays, above.    '
    ' CustPoints calculated in CustPoints_Ok.  '
    ' DistribPts calc below in Calculate_Shape.'
    Calculate_CustPoints
    Calculate_Winners
    Calculate_LosingTrickCount
    Calculate_QuickTricks
    Calculate_Position
    Calculate_Stoppers
    Calculate_Shape
    ' Intermed is calculated in the sub:   '
    '     ConvertHandStringToArray, above. '
    ' Vul. is specified by client.         '
    ' Suits: values used in Suits_Ok are   '
    ' calculated in other routines above.  '
    GetOpeningNTpts
End Sub


Private Sub Calculate_BidNum()
    '*************************************************
    ' From Prior Calls provided by Client, calculate *
    '     1. The number of the current Call.         *
    '     2. The number of the current Bid, starting *
    '           with the first non-passing Call).    *
    '           With Prior Calls of P-P-1C-P, the    *
    '           the current Call Number is 5 and     *
    '           the current Bid Number is 3.         *
    '     3. Current round of bidding (1, 2, etc.)   *
    '     4. Current bidder's position/chair (1-4).  *
    ' str_PriorBids = str_PriorCalls w/o opening P's.*
    '*************************************************

    Dim pb As String
    Dim pc As String
    Dim i As Integer
    Dim origPB As String

    Erase str_Bid()
    Erase str_Call()

    str_PriorBids = ""

    If str_PriorCalls = "" Then
        int_BidNum = 1
        int_CallNum = 1
        int_BidRound = 1
        int_PosNum = 1
        str_Position = "12FUN"
        ' 1=1st seat              '
        ' 2=2nd seat              '
        ' 3=3rd seat              '
        ' 4=4th seat              '
        ' F=First/Second seat     '
        ' T=Third/Fourth seat     '
        ' Y=passed hand           '
        ' N=unpassed              '
        ' P=partner is passed hand'
        ' U=partner is unpassed   '
        Exit Sub
    End If
    'Now there must be a Prior Call but not
    'necessarily a Prior Bid (e.g.: "P").

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
        str_Call(3) = "P"
        int_CallNum = 3
        str_Position = "1FP"
        If Mid$(pc, 6, 1) = "-" Then
            str_PriorBids = Trim$(Mid$(pc, 7))
        End If

    ElseIf Left$(pc, 3) = "P-P" Then
        str_Call(1) = "P"
        str_Call(2) = "P"
        int_CallNum = 2
        str_Position = "1FP" ' (1=1st bidding pos., F=1st/2nd seat, P=pard is passed. '
        If Mid$(pc, 4, 1) = "-" Then
            str_PriorBids = Trim$(Mid$(pc, 5))
        End If

    ElseIf Left$(pc, 1) = "P" Then
        str_Call(1) = "P"
        int_CallNum = 1
        str_Position = "1FU"
        If Mid$(pc, 2, 1) = "-" Then
            str_PriorBids = Trim$(Mid$(pc, 3))
        End If

    Else
        str_Position = "1FU"
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
        DoEvents
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


Private Sub Calculate_CustPoints()
    ' Called by Sub AnalyzeHand. '
    
    Dim n As Integer ' suit number 1-4 '
    Dim nV As Integer
    Dim s As String
    Dim X As String
    
    
    int_MyNumAces = 0
    int_MyNumKings = 0
    bln_MyTrumpK = False
    bln_MyTrumpQ = False
    int_PardsVoidSuit = 0
    
    ' Check for trump suit: '

    If int_TrumpSuit = 0 Then
        n = InStr(str_PriorBids, "@")
        If n = 0 Then Exit Sub
                                                               '   |   '
        s = Mid$(str_PriorBids, n - 1, 1) ' n = trump suit # in 4N@30  '
        X = Mid$(str_PriorBids, n + 1, 1) ' = "X" in 4H@XC '   '   |   '
        
        If s = "N" Then ' Ex: 4N@30 '
            int_TrumpSuit = Val(Mid$(str_PriorBids, n + 1, 1)) ' =3 in 4H30 '
            
        ElseIf X = "X" Then ' Exclusion Blackwood '
            ' Ex: 4H@XC : H=pard's void. @X=Exclusion. C=trump suit. '
            
            s = Mid$(str_PriorBids, n + 2, 1)
            int_TrumpSuit = InStr("CDHS", s) ' s = H '
            
            nV = Mid$(str_PriorBids, n - 1, 1) ' S '
            int_PardsVoidSuit = InStr("CDHS", nV) ' 3 '
            
            If Left$(str_Suit(int_PardsVoidSuit), 1) = "A" Then
                Dim voidAce As Integer
                voidAce = 1  ' Subtracted out of Ace count below. '
            End If
        
        Else
            s = Mid$(str_PriorBids, n + 1, 2)
            If s = "Mw" Then
                s = Mid$(str_PriorBids, n - 1, 1)
                int_TrumpSuit = InStr("CD", s)
            End If
        End If
    End If
    
    
    For n = 1 To 4 ' C to S '
        If Left$(str_Suit(n), 1) = "A" Then
            int_MyNumAces = int_MyNumAces + 1
        End If
        
        If InStr(str_Suit(n), "K") > 0 Then
            int_MyNumKings = int_MyNumKings + 1
            If n = int_TrumpSuit Then
                bln_MyTrumpK = True
            End If
        End If
    Next
    If int_MyNumAces > 0 Then int_MyNumAces = int_MyNumAces - voidAce
    
    If InStr(str_Suit(int_TrumpSuit), "Q") > 0 Then
        bln_MyTrumpQ = True
    End If
        
    If InStr(str_Suit(int_TrumpSuit), "K") > 0 Then
        bln_MyTrumpK = True
    End If
        
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

    ' Check for 13 cards and 3 dashes. '
    '   Example: "AK432-K85-QJT2-3"    '
    If Len(str_PlayersHand) <> 16 Then
        int_ErrCode = 3
        str_ErrMsg = _
            "Hand doesn't have 13 cards: " & _
            str_PlayersHand
        Exit Sub
    End If

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
            DoEvents
        Loop

        strCard = Mid$(str_PlayersHand, intCardPtr, 1)
        intCardNum = InStr(" 23456789TJQKA", strCard)
        ' Check if card is valid: '
        If intCardNum < 2 Then
            'Character in the hand is not 2-A.
            'Return error code and message:
            int_ErrCode = 4
            str_ErrMsg = strCard & " in " & _
                         str_SuitName(s) & _
                         " in " & str_PlayersHand & _
                         " is not a valid card."
            Exit Sub

        ElseIf intCardNum >= intLastCard Then
            'If card number is greater than the
            'last card,  it is out of sequence.
            'Return error code and message:
            int_ErrCode = 5
            str_ErrMsg _
                = "The " & strCard & " of " _
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

    For s = 1 To 4
        If int_SuitLen(s) > 4 And int_HCP(s) = 10 Then
            bln_SolidSuit(s) = True
        ElseIf int_SuitLen(s) > 5 And Left$(str_Suit(s), 4) = "AKQT" Then
            bln_SolidSuit(s) = True
        ElseIf int_SuitLen(s) > 8 And int_HCP(s) = 9 Then
            bln_SolidSuit(s) = True
        ElseIf int_SuitLen(s) > 7 And int_HCP(s) > 8 Then
            bln_SolidSuit(s) = True
        End If
    Next

    ' Zero total for hand: '
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
                'sng_NumRunners(s) = 1
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
    '***************************************************
    '   Calculations per www.PhillipAdlerBridge.com &  *
    '   www.bridgehands.com/L/Losing_Trick_Count.htm   *
    '***************************************************

    Dim s As Integer
    Dim sL As String
    Erase int_NumLosers()

    For s = 1 To 4
        If Len(str_Suit(s)) = 0 Then
            'No losers in a void.

        ElseIf Len(str_Suit(s)) = 1 Then
            If str_Suit(s) <> "A" Then
                int_NumLosers(s) = int_NumLosers(s) + 1
            End If

        ElseIf Len(str_Suit(s)) = 2 Then
            If str_Suit(s) = "AK" Then
                'No losers.
            ElseIf str_Suit(s) = "AQ" Then
                int_NumLosers(s) = int_NumLosers(s) + 0.5
            ElseIf Left$(str_Suit(s), 1) = "A" Then
                int_NumLosers(s) = int_NumLosers(s) + 1
            ElseIf Left$(str_Suit(s), 1) = "K" Then
                int_NumLosers(s) = int_NumLosers(s) + 1
            Else
                int_NumLosers(s) = int_NumLosers(s) + 2
            End If

        Else ' Suit len = 3+ '
            sL = Left$(str_Suit(s), 3)

            If sL = "AKQ" Then
                'No losers

            ElseIf Left$(str_Suit(s), 2) = "AK" Then 'AK2 to AKJ.
                int_NumLosers(s) = int_NumLosers(s) + 1
                ' To be consistent, it seems like AKJ should '
                ' just be .5, but that's not how they do it. '

                ' Losing_Trick_Count.htm says AQ= .5
                
            ElseIf sL = "AQJ" Or sL = "AQT" Then
                int_NumLosers(s) = int_NumLosers(s) + 1

            ElseIf sL = "AJT" Then
                int_NumLosers(s) = int_NumLosers(s) + 1.5

            ElseIf Left$(str_Suit(s), 1) = "A" Then 'A32 to AJ9.
                int_NumLosers(s) = int_NumLosers(s) + 2

            ElseIf sL = "KQJ" Then
                int_NumLosers(s) = int_NumLosers(s) + 1

            ElseIf sL = "KQT" Then
                int_NumLosers(s) = int_NumLosers(s) + 1.5

            ElseIf Left$(str_Suit(s), 1) = "K" Then 'Kxx to KJT
                int_NumLosers(s) = int_NumLosers(s) + 2
                
            ElseIf Left$(str_Suit(s), 2) = "QJ" Or _
                   Left$(str_Suit(s), 2) = "QT" _
            Then
                int_NumLosers(s) = int_NumLosers(s) + 2
                ' KQx is 2 losers and so is QJx? '
                ' KQx should be 2.5, but so be it.

            ElseIf Left$(str_Suit(s), 1) = "Q" Then 'Q2 to Q9.
                int_NumLosers(s) = int_NumLosers(s) + _
                                  2.5
            Else 'none of the top 3 honors
                int_NumLosers(s) = int_NumLosers(s) + 3
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
        sng_QuickTricks(s) = 0
        
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
    
    ' Controls are not the same as Stoppers when control   '
    ' bidding for slam purposes. See Sub Calculate_Controls'

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

    ' 2018 Update: I'm not sure where the idea came from   '
    ' of not counting a suit stopper if partner has bid    '
    ' the suit, but it is complicating things. For example '
    ' if N opens 1H and S bids 1S and N has Kxx in Spades  '
    ' and a bid like ID#17301 requires 3 suits stopped to  '
    ' bid 1N it would be nice if N had stoppers in C,D,S   '
    ' but that hand has Kxx-AQJTx-xx-ATx, so just 2 suits  '
    ' stopped outside of Spades, yet I prefer 1N over 2H.  '
    ' Plus in this case, S had QT94 but for a 1-over-1, he '
    ' could have had a lot worse. An alternative would be  '
    ' to require no more than 1 suit unstopped not counting'
    ' partner's, but that sounds complicated to code. '

GoTo SkipPard
    ' Get partner's suit: '
    If int_BidNum > 2 Then
        If int_Pards_Suit(1) = 99 Then
            GetOtherPlayersData
            pard = int_Pards_Suit(1)
        Else
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
SkipPard:

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
                If bln_Suit(s, 9) Or bln_Suit(s, 8) Then
                    ' QJT9 or QJT8 '
                    sng_NumStops(s) = 2
                Else
                    ' QJx...
                    sng_NumStops(s) = 1
                End If

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

                If bln_Suit(s, 8) Then ' QT8x looks like stopper to me. '
                    ' QT8x '
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
                    sng_NumStops(s) = 0.5
                End If
            End If

        ElseIf bln_Suit(s, Jack) And _
               int_SuitLen(s) > 3 _
        Then
            ' J???... '
            If int_SuitLen(s) > 4 Then
                sng_NumStops(s) = 1

            ElseIf bln_Suit(s, Ten) And _
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
        ' XXXoutside of partner's 1st suit.XXX' no more '
        If sng_NumStops(s) > 0 Then ' And s <> pard Then
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
    '    u = unbalanced (but no singleton or void)  *
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
    Dim v As Integer   ' void       '
    Dim c As String    ' card '

    Erase int_DistribPts()

    str_Shape = ""     ' 1, 2, 3, b '
    str_Shortness = "" ' s or v     '

    ' Count suits, shortness: '
    For s = 1 To 4
        If int_SuitLen(s) > 5 Then
            n6 = n6 + 1

        ElseIf int_SuitLen(s) = 5 Then
            n5 = n5 + 1

        ElseIf int_SuitLen(s) = 4 Then
            n4 = n4 + 1

        ElseIf int_SuitLen(s) = 2 Then
            c = Left$(str_Suit(s), 1)
            'If int_HCP(s) < 3 Then
                ' Don't count distribution for Qx, Jx. '
            '    If InStr("QJ", c) = 0 Then
                    ' Don't count distribution for Qx or Jx. '
                    ' Only count the high card points. '
            '        int_DistribPts(s) = int_DistribPts(s) + 1
            '    End If
            ' The above is too confusing when editing entries. '
            'Else
                int_DistribPts(s) = int_DistribPts(s) + 1
            'End If

        ElseIf int_SuitLen(s) = 1 Then
            s1 = s1 + 1
            'If int_HCP(s) = 0 Then
                ' Only count distribution for 2-10. '
                int_DistribPts(s) = int_DistribPts(s) + 2
            'End If

        ElseIf int_SuitLen(s) = 0 Then
            v = v + 1
            'If int_HCP(s) = 0 Then
                ' Only count distribution for 2-10. '
                int_DistribPts(s) = int_DistribPts(s) + 3
            'End If

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

    ElseIf n5 = 2 Or _
           n6 = 2 Or _
           (n4 > 0 And _
              (n5 = 1 Or n6 = 1)) Or _
           (n5 = 1 And n6 = 1) _
    Then
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
        str_Shortness = "v"

    ElseIf s1 > 0 Then
        ' We have a singleton.'
        str_Shortness = "s"

    End If

End Sub


Private Sub Calculate_Position()
    ' Calculate "position" variable where '  Examples:    '
    '   Y = bidder is a passed hand       '  P-P-1S-P-*   '
    '   N = bidder is an unpassed hand    '  1C-P-1H-P-*  '
    '   1 = bidder is in 1st seat         '  * (any call) '
    '   2 = bidder is in 2nd seat         '  1D-* or P-*  '
    '   3 = bidder is in 3rd seat         '  1D-P-*, P-P-*'
    '   4 = bidder is in 4th seat         '  1C-P-1H-* etc'
    '   F = opener is 1st or 2nd seat     '  bid or P-bid '
    '   T = opener is 3rd or 4th seat     '  P-P-*, PPP-* '
    '   Y = bidder is a passed hand       '  P-1C-1D-P-*  '
    '   N = bidder is not a passed hand   '  1H-1S-P-2S-* '
    '   P = partner is a passed hand      '  P-1C-*       '
    '   U = partner is an unpassed hand   '  1C-P-*       '
    ' The variable str_Position can have  '
    ' more than one value because the     '
    ' codes are not mutually exclusive &  '
    ' an entry may specify 1 code rather  '
    ' than another, both of which match   '
    ' the current bidder's position.      '  / Y = a passed hand: P-1C-1D-P-*    '
    ' Example: Y1P - see the codes above. '--  3 = 3rd seat:      x-x-*          '
    ' That's why you see --               '  \ P = pard passed:   P-P-P-1C-P-P-* '
    '  str_Position = str_Position & ...  '                       ^

    ' Note that the Player Number which   '
    ' comes after Round and before Prior  '
    ' Bids in the Editor excludes leading '
    ' Passes while the Position value in  '
    ' the drop-down list box in the Editor'
    ' DOES include leading Passes, so the '
    ' two are not comparable. The former  '
    ' value is int_BidNum while Position  '
    ' is int_CallNum. '

    ' Also note that being a passed hand  '
    ' refers to the first chance to bid,  '
    ' not his most recent turn. '

    Dim p As String
    Dim n As Integer

    n = int_CallNum Mod 4
    
    ' See if bidder is unpassed: '
    If int_CallNum < 5 And int_CallNum > 0 Then
        ' Example: If int_CallNum = 4 then player is '
        ' in 4th seat and could not have passed yet. '
        str_Position = int_CallNum & "N" ' = unpassed hand '

    ElseIf str_Call(int_CallNum) <> "P" Then
        ' If more than 4 calls, then  '
        ' this player's position num  '
        ' (e.g.: 3) is also the array '
        ' pointer to his first call in'
        ' the array, like str_Call(3).'
        str_Position = "N" ' = unpassed hand '

    Else
        str_Position = "Y" ' = passed hand '
    End If

    ' Add F, T: '
    If int_CallNum < 3 Then
         ' In 1st/2nd chair: '
        str_Position = str_Position & "F"
    Else ' In 3rd/4th chair: '
        str_Position = str_Position & "T"
        'If int_BidNum = 1 Then
        '    ' Opener '
        '    str_Position = str_Position & int_PosNum
        'End If
    End If

    ' Check responder's pos.: '
    int_Pards_Pos = int_CallNum - 2
    ' e.g.: you=3, pard=3-2=1,  '
    If int_Pards_Pos < 1 Then
        ' e.g.: you=1, 1-2=-1, -1+4=3 '
        int_Pards_Pos = int_Pards_Pos + 4
    End If
    If str_Call(int_Pards_Pos) <> "P" Or _
       int_CallNum < 3 _
    Then ' pard didn't pass last time '
         ' or he hasn't bid yet.      '
        str_Position = str_Position & "U" ' = pard is Unpassed '
    Else
        str_Position = str_Position & "P" ' = pard is Passed '
    End If
End Sub


Sub GetOpeningNTpts()
    ' Find an entry for an opening bid of 1NT which '
    ' has not been deactivated in System, Sub, or   '
    ' Pct Used.  That sets the range for HCP1N. All '
    ' other references to HCPs for NT bids should be'
    ' like Min-1, Max+1, HCP1N, HCP2N, HCP3N.       '
    '                                               '
    ' For example if HCP1N is 15-17, then Min is 15 '
    ' and Max is 17. A "good" 14 is Min-1 and a bad '
    ' 18 is Max+1.  This saves a lot of space in    '
    ' BidBase by not having a complete set of the   '
    ' same types of entries for all different ranges'
    ' of notrump openings such as 13-15, 12-14, 11- '
    ' 13, and even 10-12. '

    Dim num As Long
    Dim nm As String
    Dim selec As String
    Dim selec2 As String
    Dim s As String

    Set rs_BidTable = db_BidBase.OpenRecordset("Bids")
    With rs_BidTable
    ''''''''''''''''
    .Index = "Bid"
    .Seek "=", "1N"

    If rs_BidTable.NoMatch Then
        str_ErrMsg = "Opening 1N bid not found." ' Really shouldn't be possible."
        int_ErrCode = 55
        Exit Sub
    End If

    ' Make sure entry is activated. '
    Do
        DoEvents
        If Val(rs_BidTable!HCP) > 0 And rs_BidTable!Prior_Bids = "" Then
            num = Val("" & rs_BidTable!ConvID)
            LookUpBidName num, selec, selec2, nm

            If selec = "" And _
               selec2 = "" And _
               rs_BidTable!pct_Used = "" And _
               nm = "Opening 1N HCP Specification" _
            Then
                Exit Do
            End If
        End If

        If .EOF Then
            MsgBox "Opening 1N bid not selected in database."
            int_ErrCode = 1
            Exit Sub
        End If

        .MoveNext
        DoEvents
    Loop

    int_1NT_Min = Val(!HCP)
    int_1NT_Max = Val(Right$("  " & !HCP, 2))


    .Seek "=", "2N"

    If rs_BidTable.NoMatch Then
        str_ErrMsg = "Opening 2N bid not found." ' Really shouldn't be possible."
        int_ErrCode = 1
        Exit Sub
    End If

    ' Make sure entry is activated
    Do
        DoEvents
        If Val(rs_BidTable!HCP) > 0 And rs_BidTable!Prior_Bids = "" Then
            num = Val("" & rs_BidTable!ConvID)
            LookUpBidName num, selec, selec2, nm

            If selec = "" And selec2 = "" And rs_BidTable!pct_Used = "" Then
                Exit Do
            End If
        End If

        If .EOF Then
            MsgBox "Opening 2N bid not selected in database."
            int_ErrCode = 1
            Exit Sub
        End If

        .MoveNext
        DoEvents
    Loop

    int_2NT_Min = Val(!HCP)
    int_2NT_Max = Val(Right$("  " & !HCP, 2))


    .Seek "=", "3N"

    If rs_BidTable.NoMatch Then
        str_ErrMsg = "Opening 3N bid not found." ' Really shouldn't be possible."
        int_ErrCode = 1
        Exit Sub
    End If

    ' Make sure entry is activated
    Do
        If Val(rs_BidTable!HCP) > 0 And _
           rs_BidTable!Prior_Bids = "" _
        Then
            num = Val("" & rs_BidTable!ConvID)
            LookUpBidName num, selec, selec2, nm

            If selec = "" And _
               selec2 = "" And _
               rs_BidTable!pct_Used = "" _
            Then
                Exit Do
            End If
        End If

        If .EOF Then
            MsgBox "Opening 3N bid not selected in database."
            int_ErrCode = 1
            Exit Sub
        End If

        .MoveNext
        DoEvents
    Loop

    int_3NT_Min = Val(!HCP)
    int_3NT_Max = Val(Right$("  " & !HCP, 2))
    ''''''''
    End With
End Sub


'                      Start entry analysis routines                          '


Private Function Activation_Ok() As Boolean
    '************************************************
    ' Check Pct_Used field to see if this entry has *
    ' been disabled. If not and "M", "I" or a # is  *
    ' specified, compare it to the client specs or  *
    ' to a randomly generated %, as appropriate.    *
    '                                               *
    ' Field Values:                                 *
    '    Blank -- always consider entry.            *
    '    "0" (zero) or "x" -- always ignore entry.  *
    '    "M" -- for Matchpoint scoring only.        *
    '    "I" -- for IMPs scoring only.              *
    '    "1" to "99" -- % of time to use entry.     *
    '                                               *
    ' Also checks the Selected field in BidNames to *
    ' see if this entry's Convention is turned off. *
    '************************************************

    Dim ttlPct As Integer
    Dim pct As Integer
    Dim bon As String ' Bid Order Number may contain letters. '
    Dim num As Long
    Dim sel As String
    Dim sel2 As String
    Dim nam As String
    Dim a As String
    Dim i As Integer
'If rs_BidTable!id = 110 Then Stop
    LookUpBidName num, sel, sel2, nam

    If rs_BidTable!Bid = str_BlockedBid(0) Or _
       rs_BidTable!Bid = str_BlockedBid(1) _
    Then
        'frm.lstAnalysis.List(6) = Space$(10) & "Bid of " & rs_BidTable!Bid & " is blocked by a convention."
        Exit Function

    ElseIf Right$(rs_BidTable!Bid_Order_Number, 4) = "9999" And _
           Left$(rs_BidTable!Bid_Order_Number, 1) < "9" And _
           sel <> "" And _
           sel2 <> "" And _
           rs_BidTable!pct_Used <> "" _
    Then
        If str_BlockedBid(0) = "" Then
            str_BlockedBid(0) = rs_BidTable!Bid
        Else
            str_BlockedBid(1) = rs_BidTable!Bid
        End If
        Exit Function

    ElseIf bln_TestEntry Then
        ' Must accept this entry despite Pct setting.'
        Activation_Ok = True
        Exit Function
    End If

    On Local Error GoTo 0 ' LocalErr

    With rs_BidTable
        If lng_TestID = !id And lng_TestID <> 0 Then
            Activation_Ok = True
            Exit Function
        End If

        ' Check convention's "BidNames.Selected" code:
        num = Val("" & !ConvID)

        If num > 0 Then
        '''''''''''''''
        LookUpBidName num, sel, sel2, nam ' sel = "selected" '

        If Trim$(str_Activated) <> "" And _
           sel = "" And _
           sel2 = "" And _
           !pct_Used & "" = "" _
        Then
            ' str_Activated only has a value when '
            ' BidBase Editor is saving an entry   '
            ' in which Cnv, Sub, or Pct = 0, so   '
            ' since the entry being checked is not'
            ' deactivated, it fails to match the  '
            ' entry being saved: '
            If bln_ShowAnalysis Then
                frm.lstAnalysis.List(6) = _
                   Space$(10) & _
                   "Entry is activated; doesn't conflict with hand being saved."
            End If
            Exit Function
        ElseIf Trim$(str_Activated) <> "" Then
             ' being saved is deactivated but '
             ' not the entry being checked. '
            If bln_ShowAnalysis Then
                frm.lstAnalysis.List(6) = _
                  Space$(10) & "Entry is deactivated; " & _
                  "doesn't conflict " & _
                  "with hand being saved."
            End If
            Exit Function
        End If

        If Not (sel = "" Or !id = lng_TestID Or _
               (sel = "W" And (int_PosNum = 1 Or int_PosNum = 3)) Or _
               (sel = "T" And (int_PosNum = 2 Or int_PosNum = 4))) Or _
           Not (sel2 = "" Or !id = lng_TestID Or _
               (sel2 = "W" And (int_PosNum = 1 Or int_PosNum = 3)) Or _
               (sel2 = "T" And (int_PosNum = 2 Or int_PosNum = 4))) _
        Then   ' "W"=We  "T"=They  ""=both  "0"=neither '
            If bln_ShowAnalysis Then
              If (sel <> "0" Or sel2 <> "0") Then
                With frm.lstAnalysis
                    For i = 1 To 10
                        .AddItem ""
                    Next
                    .ListIndex = 6
                    If sel = "0" Then
                        .List(6) = Space$(30) & "Convention not activated by user."
                    ElseIf sel2 = "0" Then
                        .List(6) = Space$(25) & _
                                   "This Convention's sub not activated by user."
                    Else
                        .List(6) = Space$(20) & "Cnv = " & sel & _
                                   Space$(10) & "Sub = " & sel2
                    End If
                End With
              ElseIf !pct_Used = "0" Then
                frm.lstAnalysis.List(6) = _
                   Space$(30) & "Entry is not activated by user."
              End If
            End If
            Activation_Ok = False ' This was True with sel="0" and bombed. 05-13-2022 '
            Exit Function
        End If
        '''''''''
        End If

        If lng_TestID > 0 And Val(!pct_Used) > 0 Then
            If lng_TestID = Val(!id) Then
                Activation_Ok = True
            End If
            Exit Function
        End If

        If Trim$("" & !pct_Used) = "" Or _
           bln_CheckAllPcts Or _
           Trim$("" & !pct_Used) = str_ScoreMethod _
        Then

            Activation_Ok = True
            Exit Function

        ElseIf Val(!pct_Used) > 0 Then
            ' BidBase entry has a Pct_Used number.    '
            ' Get a random number and see if it falls '
            ' within the specified range (i.e.:       '
            ' random number <= specified Pct          '
            ' Example: If Pct = 20 and random # = 13, '
            '          then Ok, but not if # = 32.    '

            bon = !Bid_Order_Number
            Randomize Timer
            pct = Fix(Rnd * 100 + 1)

            Do
                DoEvents
                If (pct - ttlPct) <= Val(!pct_Used) Or _
                   .EOF Or _
                   lng_TestID > 0 _
                Then
                    ' If end of file, last entry must '
                    ' be Okayed for further testing:  '
                    ' Or when testing, accept any Pct.'
                    Activation_Ok = True
                    GoTo PctOk
                End If

                ttlPct = ttlPct + Val(!pct_Used)

                .MoveNext
                If .EOF Then
                    .MovePrevious
                    Activation_Ok = True
                    GoTo PctOk
                End If
                If bon <> !Bid_Order_Number Or _
                   Val(!pct_Used) = 0 _
                Then
                    ' Alternative entries must have the '
                    ' same BON. This entry doesn't, so  '
                    ' go back to prior entry and OK it: '
                    On Local Error Resume Next
                    .MovePrevious
                    On Local Error GoTo 0
                    Activation_Ok = True
                    GoTo PctOk
                End If
                DoEvents
            Loop
            ' No Do-Loop limits because exits are made   '
            ' if a Pct matches or if EOF or if Bid Order '
            ' Number changes or if next PctUsed field no '
            ' longer has a number in it.                 '


        ElseIf (Left$(!pct_Used & " ", 1) = "M" Or _
                Left$(!pct_Used & " ", 1) = "I") And _
               (Left$(!pct_Used & " ", 1) = str_ScoreMethod Or _
                lng_TestID > 0) _
        Then  ' "M"=matchpoints   "I"=IMPs   "D"=Double Dummy Analysis   '
              ' A number of 1-99 in the Pct field allows BB to randomly  '
              '   select a bid the given % of the time. UPDATE: I am not '
              '   doing that any more, but some older entries do.        '
              ' x followed by a number disables the % field but allows   '
              '   you to reinstate by by removing the x.                 '
            Activation_Ok = True
            Exit Function
        End If

    End With


    If bln_ShowAnalysis Then
        frm.Visible = True
        With frm.lstAnalysis
            .ListIndex = 6
          If rs_BidTable!pct_Used = "0" Then
            .List(6) = Space$(30) & "Entry is not activated by user."
          Else
            .List(6) = Space$(40) & "Pct Used = " & _
                       rs_BidTable!pct_Used
          End If
        End With
    End If
    Exit Function

PctOk:
    ' Redo analysis form for new entry: '
    If bln_ShowAnalysis Then
        ShowAnalysis
        frm.Show
    End If
Exit Function

LocalErr:
    MsgBox Err.Description & vbCr & _
            "in 'Activation_Ok' function." & vbCr & _
            "ID#" & rs_BidTable!id

End Function


Private Function TotalHCPs_Ok() As Boolean
    '**************************************************
    ' Purpose: See if the hand's HCPs (which were     *
    '          calculated in Sub ConvertHandToSuits)  *
    '          matches the entry's specs.             *
    '**************************************************

    Dim strHCP As String
    Dim intN As Integer
    Dim intMinHCPs As Integer
    Dim intMaxHCPs As Integer
    Dim r1 As String
    Dim n As Integer
    Dim minNT As Integer
    Dim maxNT As Integer

    On Local Error GoTo 0
    If rs_BidTable!HCP = "" Then
        TotalHCPs_Ok = True
        Exit Function
    End If

'If rs_BidTable!id = 1845 Then Stop

    strHCP = LCase$(Trim$("" & rs_BidTable!HCP)) ' Entry's specs '
    If UCase$(Left$(strHCP, 1)) = "T" Then
        Stop ' Should not be a T in HCP. '
        strHCP = Mid$(strHCP, 2)
    End If

    If strHCP = "" Then
        TotalHCPs_Ok = True
        Exit Function
    End If

    If Left$(str_PriorBids, 2) = "1N" Or _
       (Left$(strHCP, 3) = "min" And _
        Left$(rs_BidTable!Bid & "  ", 2) = "1N") Or _
       Left$(strHCP & "     ", 5) = "hcp1n" _
    Then
        minNT = int_1NT_Min
        maxNT = int_1NT_Max

    ElseIf Left$(str_PriorBids, 2) = "2N" Or _
           (Left$(strHCP, 3) = "min" And _
            Left$(rs_BidTable!Bid & "  ", 2) = "2N") Or _
           Left$(strHCP & "     ", 5) = "hcp2n" _
    Then ' Something's wrong "!Bid can't be "hcp2n" which is HCP. '
        minNT = int_2NT_Min
        maxNT = int_2NT_Max
    End If

    If strHCP = "min" Then
        ' Bidder opened NT, responder made an '
        ' invitational bid.                   '
        strHCP = "" & minNT

    ElseIf Left$(LCase$(strHCP) & "   ", 3) = "min" Then
        n = Val(Right$(strHCP, 2))
        If n = 0 Then ' Just "min+"
            strHCP = (minNT + 1) & "-" & maxNT

        Else
            strHCP = "" & (minNT + n)
        End If

    ElseIf UCase$(strHCP) = "HCP1N" Or _
           UCase$(strHCP) = "HCP2N" _
    Then
        ' The original HCP range. '
        strHCP = minNT & "-" & maxNT

    ElseIf LCase$(Left$(strHCP & " ", 1)) = "p" Then
        ' Partner opened NT. Compute responses from '
        ' minimum NT opening HCP range. Example: If '
        ' HCPs spec for this entry is "P24-25" then '
        ' if the opening range is 15-17, then the   '
        ' HCP spec for this entry is 24 minus 15 =  '
        ' 9 and 25 minus 15 = 10, or "9-10". If the '
        ' opening range is 12-14, then the range for'
        ' this entry is 24-12=12 and 25-12=13, or   '
        ' 12-13. This is a unique situation because '
        ' of 1N (and 2N) and its responses having   '
        ' such tightly defined HCP ranges, but which'
        ' also have a number of alternative ranges. '
        If Len(strHCP) > 2 Then
            strHCP = Mid$(strHCP, 2) ' Get rid of "P"  '
            If Left$(strHCP, 1) = "<" Then
                strHCP = "<" & (Val(Mid$(strHCP, 2)) - minNT)
            Else
                r1 = "" & (Val(strHCP) - minNT)
                If Len(strHCP) > 3 Then ' e.g.: 24-25 '
                    strHCP = r1 & "-" & _
                             (Val(Right$(strHCP, 2)) - _
                             minNT)
                Else
                    strHCP = r1 & Right$(strHCP, 1)
                End If
            End If
        End If
        ' Fixed up str_HCP, now test it below: '
    ElseIf Left$(strHCP, 3) = "1n+" Then
        intMinHCPs = int_1NT_Min + Val(Mid$(strHCP & " ", 4))
        intMaxHCPs = int_1NT_Min + Val(Mid$(strHCP & "    ", 6))
        If intMaxHCPs <= intMinHCPs Then
            MsgBox "Max is " & int_1NT_Max & vbCr & _
                   "Min is " & int_1NT_Min
            Exit Function
        End If

        If int_HCP(0) >= intMinHCPs And _
           int_HCP(0) <= intMaxHCPs _
        Then
            TotalHCPs_Ok = True
            Exit Function
        End If

        If bln_ShowAnalysis Then
            frm.Visible = True
            frm.lstAnalysis.ListIndex = 4
            frm.lstAnalysis.List(4) = _
              frm.lstAnalysis.List(4) & _
              "            NO MATCH:  Hand has " & _
              int_HCP(0)
        End If
        Exit Function
    End If

    intMinHCPs = Val(strHCP)

    If Left$(strHCP, 1) = "<" Then 'e.g.: <10
        intMinHCPs = 0
        intMaxHCPs = Val(Mid$(strHCP, 2)) - 1

    ElseIf Left$(strHCP, 1) = ">" Then 'e.g.: >10
        intMaxHCPs = 40

    End If

    intN = InStr(strHCP, "-")
    If intN > 0 Then
        If intN = Len(strHCP) Then
            'e.g.: can be 6- but normally use <7 instead
            intMaxHCPs = Val(strHCP)
            intMinHCPs = 0
        Else                 'e.g.: 6-10
            intMaxHCPs = Val(Mid$(strHCP, intN + 1))
        End If

    ElseIf Right$(strHCP, 1) = "+" Then 'e.g.: 6+

        intMaxHCPs = 40

    ElseIf Left$(strHCP, 1) <> "<" And _
           Left$(strHCP, 1) <> ">" _
    Then
        intMaxHCPs = intMinHCPs

    End If

    If int_HCP(0) >= intMinHCPs And _
        int_HCP(0) <= intMaxHCPs _
    Then

        TotalHCPs_Ok = True

        Dim im As Integer
        Dim adj As Integer

        im = Val(rs_BidTable!Intermed)
        If im > 0 Then
            adj = int_HCP(0) - intMinHCPs
            If adj > 0 Then
               ' int_Intermediates(0) = int_Intermediates(0) + adj
               ' This is screwy.  int_Intermediates(0) is "Intermediates" '
               '                  which is in the 2nd column of Inputs.   '
            End If
        End If

    ElseIf bln_ShowAnalysis Then
        frm.Visible = True
        With frm.lstAnalysis
            .ListIndex = 4
            .List(4) = .List(4) & _
                "            NO MATCH:  Hand has " & int_HCP(0)
        End With
    'ElseIf str_Looking4Majors = "" Then
    '    int_ErrCode = 96
    End If

End Function


Private Function CustPoints_Ok()
    '*****************************************************
    ' Purpose:  Custom point assignments can be made for *
    '           A-K-Q-J-T.  Example: A=2, K=1.           *
    ' So first calculate the total custom points for the '
    ' hand and then compare the total to the total in the'
    ' entry.
    '*****************************************************

    Dim honQty(5) As Integer
    Dim cp(5) As Single
    Dim cardsSpecified As Boolean
    Dim rp As String
    Dim ReqdPts As Single
    Dim ttlPts As Single
    Dim blnLessThan As Boolean
    Dim s As Integer
    Dim cd As Integer
    Dim c As Integer

    ' Using   "" &   before DAO field '
    ' prevents a null error from DAO. '
    ' Could use On Error Resume Next, '
    ' but that can hide other errors. '

    With rs_BidTable
        If !Cust_Pts & "" = "" Then
            ' Custom points not specified. '
            CustPoints_Ok = True
            Exit Function
        End If

        ' Convert entry's CP specs to array: '
        'cp(0) = Val("" & !Cust_Pts)
        cp(1) = Val("" & !A_Cust)
        cp(2) = Val("" & !K_Cust)
        cp(3) = Val("" & !Q_Cust)
        cp(4) = Val("" & !J_Cust)
        cp(5) = Val("" & !T_Cust)

        ' Required (total) custom points: '
        rp = Val("" & !Cust_Pts)
        ReqdPts = Val(rp)

        ' Identify math sign for later use: '
        If ReqdPts = 0 And _
           Len("" & !Cust_Pts) > 1 _
        Then 'ex: <4, >4
            ReqdPts = Val(Mid$("" & !Cust_Pts, 2))
            If Left$("" & !Cust_Pts, 1) = "<" Then
                blnLessThan = True
            End If
        End If

        ' If nothing specified, exit with OK: '
        If ReqdPts = 0 And _
           cardsSpecified = False _
        Then
            CustPoints_Ok = True
            Exit Function
        End If

        ' Sum up Custom Points and number of each honor: '
        For s = 1 To 4
            For cd = Ten To Ace
                If bln_Suit(s, cd) Then
                    ' In cp array, Ace=1, Ten=5.    '
                    ' In suit array, Ace=14, Ten=10.'
                    ' Next line converts the card # '
                    ' to cp array index (as 14 to 1)'
                    c = 15 - cd
                    ttlPts = ttlPts + cp(c)
                    honQty(c) = honQty(c) + 1
                End If
            Next
        Next

        ' If Custom Points are required, see if this hand '
        ' matches:                                        '
        If ReqdPts > 0 Then
            If blnLessThan Then
                If ttlPts < ReqdPts Then
                    CustPoints_Ok = True
                    Exit Function
                End If

            ElseIf ttlPts >= ReqdPts Then
                ' Notice that 9, 9+, and >8 are all the same. '
                ' Any value not preceded by "<" is considered '
                ' a minimum, not a specific value.            '
                CustPoints_Ok = True
                Exit Function
            End If

        ' If no total custom points were specified, then '
        ' it must be a spec for actual number of honors: '
        ElseIf honQty(1) >= cp(1) Then
            If honQty(2) >= cp(2) Then
                If honQty(3) >= cp(3) Then
                    If honQty(4) >= cp(4) Then
                        If honQty(5) >= cp(5) Then
                            CustPoints_Ok = True
                        Exit Function
                        End If
                    End If
                End If
            End If
        End If

        ' Custom Points do not match. '

        If bln_ShowAnalysis Then
            With frm.lstAnalysis
                .ListIndex = 7
                If ReqdPts > 0 Then
                    .List(7) = .List(7) & _
                                "         NO MATCH:  Hand has " & _
                                ttlPts
                ElseIf cp(0) > 0 Then
                    .List(7) = "Hon.reqd. = " & cp(0) & _
                               "         NO MATCH:  Doesn't have " & _
                               "specified honors."
                Else
                    CustPoints_Ok = True
                End If
            End With
        End If
    End With
End Function


Private Function DistribPts_Ok()
    ' Check Distribution Points '
    ' BidBase format examples:  '
    '       2+                  '
    '       <3                  '
    '       >1                  '
    '       3-4                 '
    '       T14+ (T=total HCP+) '

    Dim tdp As String 'total dist. points
    Dim tdpn As Integer 'tdp as number
    Dim sign As String  'math sign
    Dim totalDistribHCP As Boolean
    Dim tp As Integer
    Dim minDistribPts As Integer
    Dim maxDistribPts As Integer
    Dim dash As Integer
    Dim s As Integer
    Dim sp(4)  As Integer ' suit points '
    Dim ls1 As Integer ' long suit 1 '
    Dim ls2 As Integer ' long suit 2 '

    'Set distibution points required:
    tdp = Trim$("" & rs_BidTable!Dist_Pts)
'If rs_BidTable!id = "18336" Then Stop
    If tdp = "" Then
        ' None required. '
        DistribPts_Ok = True
        Exit Function
    End If

    If tdp = "Rule20" Then
        ' Add HCP to 2 longest suit lengths. '
        ' If result is 20+, open the bidding.'
        ' http://fourseasonsbridge.com/harold/IL2/12.pdf '
        For s = 1 To 4
            If sp(s) = 1 Then
                If int_HCP(s) > 3 Then
                    ' Don't count singleton honor except Ace. '
                    sp(s) = 4
                End If
            ElseIf sp(s) = 1 Then
                If int_HCP(s) > 2 Then
                    ' Don't count Qx or Jx. '
                    sp(s) = int_HCP(s)
                End If
            Else
                sp(s) = int_HCP(s)
            End If
            If int_SuitLen(s) > ls1 Then
                ls1 = int_SuitLen(s)
            ElseIf int_SuitLen(s) > ls2 Then
                ls2 = int_SuitLen(s)
            End If
        Next
        If sp(1) + sp(2) + sp(3) + sp(4) + ls1 + ls2 >= 20 Then
            DistribPts_Ok = True
            Exit Function
        End If
    End If

    ' Compute tdpn and see if there '
    ' is a math sign:
    tdpn = Val(tdp)
    If tdpn = 0 And Len(tdp) > 1 Then
        ' Must be a sign in front  '
        ' of the number, like  <4 .'
        tdpn = Val(Mid$(tdp, 2))
        sign = Left$(tdp, 1)
    End If

    ' If "T" is before specified points, '
    ' then the spec is for distribution  '
    ' points plus HCPs.                  '
    If sign = "T" Then
        totalDistribHCP = True
        If Val(Mid(tdp, 2, 1)) = 0 Then
            If Val(Mid(tdp, 2, 1)) = 0 Then
                sign = Mid$(tdp, 2, 1) ' e.g.: T<13
                tdpn = Val(Mid$(tdp, 3))
            Else
                sign = Mid$(tdp, 2, 2) ' e.g.: T<=13
                tdpn = Val(Mid$(tdp, 4))
            End If

        ElseIf Val(Right$(tdp, 2)) < 10 Then ' e.g.: T13+
            sign = Right$(tdp, 1)
        End If
    End If

    dash = InStr(tdp, "-")    ' e.g.: 3-4

    ' Calculate min and max distribution '
    ' points:                            '

    If sign = "<" Then     'e.g.: <4
        minDistribPts = 0
        maxDistribPts = tdpn - 1

    ElseIf sign = ">" Then 'e.g.: >1
        minDistribPts = tdpn + 1
        maxDistribPts = 40

    ElseIf Right$(tdp, 1) = "+" Then 'e.g.: 2+
        minDistribPts = tdpn
        maxDistribPts = 40

    ElseIf dash Then
        If dash = Len(tdp) Then 'e.g.: 1-
            ' This is not a recommended format.
            minDistribPts = 0
            maxDistribPts = tdpn

        Else                 'e.g.: 2-3
            minDistribPts = tdpn
            maxDistribPts = Val(Mid$(tdp, dash + 1))
        End If

    ElseIf tdpn > 0 Then     'e.g.: 2
        minDistribPts = tdpn
        maxDistribPts = minDistribPts
    End If


    ' Compare actual points to min and max: '

    If totalDistribHCP Then
        ' Total distribution plus HCP specified: '
        tp = int_DistribPts(0) + int_HCP(0)
        If tp >= minDistribPts And _
           tp <= maxDistribPts _
        Then
            DistribPts_Ok = True

            Dim im As Integer
            Dim adj As Integer

            im = Val(rs_BidTable!Intermed)
            If im > 0 Then
                adj = tp - minDistribPts
                If adj > 0 Then
                    totalDistribHCP = totalDistribHCP + adj
                End If
            End If

            Exit Function
        End If

    Else ' Only distribution points specified: '
        If int_DistribPts(0) >= minDistribPts And _
           int_DistribPts(0) <= maxDistribPts _
        Then
            DistribPts_Ok = True
            Exit Function
        End If
    End If

    ' Test failed. '

    If bln_ShowAnalysis Then
        With frm.lstAnalysis
            frm.Visible = True
            .ListIndex = 5
            If totalDistribHCP Then
                tdp = "            NO MATCH:  Total points = " & tp
            Else
                tdp = "            NO MATCH:  Distrib. pts = " & int_DistribPts(0)
            End If
            .List(5) = .List(5) & tdp
            Exit Function
            .List(5) = .List(5) & tdp & _
                int_DistribPts(0)
        End With
    End If

End Function


Private Function Winners_Ok()
    ' See if Winners specified match the hand. '
    Dim n As Integer

    If Left(rs_BidTable!Winrs & " ", 1) = "<" Then
        n = Val(Mid$(rs_BidTable!Winrs & " ", 2))
        If sng_NumWinners(0) < n Then
            Winners_Ok = True
            bln_WinnersOK = True
            Exit Function
        End If

    ElseIf Trim$("" & rs_BidTable!Winrs) = "" Or _
       sng_NumWinners(0) >= Val("" & rs_BidTable!Winrs) _
    Then
        Winners_Ok = True
        bln_WinnersOK = sng_NumWinners(0) >= Val("" & rs_BidTable!Winrs)
        Exit Function
    End If

    If bln_ShowAnalysis Then
        With frm.lstAnalysis
            .ListIndex = 8
            .List(8) = .List(8) & _
                "            NO MATCH:  Number of winners = " & _
                sng_NumWinners(0)
        End With
    End If

End Function


Private Function Losers_Ok()
    ' See if Losers specified match the hand. '
    Dim Losers As String
    Dim minLosers As Boolean
    
    Losers = Trim$("" & rs_BidTable!Losers)

    If Right$(" " & Losers, 1) = "+" Then
        ' # losers more than the specified #. '
        If int_NumLosers(0) >= Val(Losers) Then
            Losers_Ok = True
            Exit Function
        End If

    ElseIf Losers = "" Or _
       (int_NumLosers(0) <= Val(Losers) And _
        Val(Losers) <> 0) _
    Then
        ' # losers
        Losers_Ok = True
        Exit Function

    End If

    If bln_ShowAnalysis Then
        ' Failed test. '
        With frm.lstAnalysis
            .ListIndex = 9
            .List(9) = .List(9) & _
                "            NO MATCH:  Number of losers = " & _
                int_NumLosers(0)
        End With
    End If

End Function


Private Function QuickTricks_Ok()
    ' See if Quick Tricks specified match the hand. '

    If Trim$(rs_BidTable!QTs) = "" Or _
       sng_QuickTricks(0) >= Val(rs_BidTable!QTs) _
    Then

        QuickTricks_Ok = True

    ElseIf bln_ShowAnalysis Then
        ' Failed test. '
        With frm.lstAnalysis
            .ListIndex = 10
            .List(10) = .List(10) & _
                "            NO MATCH:  Hand has " & _
                sng_QuickTricks(0)
        End With
    End If
End Function


Private Function Stoppers_Ok()
    '*******************************************'
    ' Purpose: See if the current entry's specs '
    '          for suits stopped match the hand.'
    '                                           '
    ' Number of suits stopped for the hand are  '
    ' calculated in Sub Calculate_Stoppers().   '
    '*******************************************'

    If Trim$(rs_BidTable!Stopped_Suits) = "" Then
        Stoppers_Ok = True

    ElseIf sng_NumStops(0) >= Val(rs_BidTable!Stopped_Suits) Then
        Stoppers_Ok = True

    ElseIf bln_ShowAnalysis Then
        '  Failed test.  '
        With frm.lstAnalysis
            .ListIndex = 12
            .List(12) = .List(12) & _
                "            NO MATCH:  Number of stops = " & _
                sng_NumStops(0)
        End With

    End If

End Function


'****************************************************
' Procedure: Shape_Ok                               *
' Purpose:                                          *
    ' The values in the Shape field in the database *
    ' come from the drop-down list box for shape in *
    ' the BidBase view-edit program:                *
    '    1 = 1-suiter (6+/3-)                       *
    '    2 = 2-suiter (6+/4+, 5+/5+)                *
    '    3 = 3-suiter (4441, 5440)                  *
    '    b = balanced (4333, 4432, 5332)            *
    '    u = unbalanced (no singleton or void)      *
    '    s = singleton                              *
    '    v = void                                   *
    '    e = either sing-void                       *
'****************************************************
'
Private Function Shape_Ok()
    Dim shp As String
    Dim shpOk As Boolean

    shp = LCase$(Trim$("" & rs_BidTable!Shape))
    shpOk = False

    Select Case shp
        Case ""
            shpOk = True

        Case "1", "2", "3", "b"
            ' 1-, 2-, 3-suiter or balanced
            If str_Shape = shp Then
                shpOk = True
            End If

        Case "u"
            ' Unbalanced
            If str_Shape <> "b" Then
                shpOk = True
            End If

        Case "s", "v", "e"
            ' Singleton, void, or either.               '
            ' str_Shortness can only be "s", "v" or "". '
            If (shp = "e" And str_Shortness <> "") Or _
               str_Shortness = shp _
            Then
                shpOk = True
            End If

    End Select

    Shape_Ok = shpOk

    If shpOk = False And _
       bln_ShowAnalysis _
    Then
        With frm.lstAnalysis
            .ListIndex = 15
            If shp = "b" Then
                .List(15) = .List(15) & _
                            "            NO MATCH:  Hand is not balanced."
            Else
                .List(15) = .List(15) & _
                            "            NO MATCH:  Shape is " & str_Shape
            End If
       End With
    End If

End Function


'************************************
' Procedure: Intermed_Ok            '
' Purpose:                          '
    ' Test for Intermediates (8-10) '
    ' In the routine for testing    '
    ' HCPs, int_Intermediates was   '
    ' adjusted for the excess of    '
    ' total HCPs over required HCPs.'
    ' This is because it makes no   '
    ' sense to say that 12+ HCPs +  '
    ' 1 intermediate is required,   '
    ' but that a hand with 13 HCPs  '
    ' fails because it has a Jack   '
    ' instead of a 9.               '
'************************************

Private Function Intermed_Ok()
    Dim i As Integer   ' # intermediates required for suit.      '
    Dim adj As Integer ' Adjustment for extra HCPs above minimum.'
    Dim im As Integer
    Dim iStr As String

    With rs_BidTable

        If "" & !Intermed = "" Then
            ' None specified '
            Intermed_Ok = True
            Exit Function
        End If

        iStr = !Intermed

        i = Val(iStr)

    End With

    ' Test: '

    If iStr = "0" Then ' Or iStr = "1" Then
        ' When 0 intermediates specified, it   '
        ' it seems logical that this is meant  '
        ' to indicate poor intermediates, so   '
        ' the entry should be rejected if hand '
        ' has MORE than 0. '
        'If int_Intermediates(0) >= i Then
        If int_Intermediates(0) = 0 And i = 0 Then
            Intermed_Ok = True
        End If

    ElseIf int_Intermediates(0) >= i Then
        ' Note: int_Intermediates(0) is adjusted '
        ' for extra HCPs in Sub TotalHCPs_OK     '
        Intermed_Ok = True

    ElseIf bln_ShowAnalysis Then
        On Local Error Resume Next
        With frm.lstAnalysis
            .ListIndex = 11
            .List(11) = .List(11) & _
                        "            NO MATCH:  Hand has " & _
                        (int_Intermediates(0))
        End With
        On Local Error GoTo 0
    End If

End Function


Private Function Vul_Ok()
    ' Check vulnerability.      '
    ' Values in BidBase can be: '
    '    F = favorable / they   '
    '    U = unfavorable / we   '
    '    B = both               '
    '    N = neither            '
    '    V = vulnerable bidder  '
    '    X = not vul. bidder    '
    '
    '    blank = any vul is ok  '

    Dim v As String ' vul spec from BB '

    v = rs_BidTable!Vul

    If v = str_Vul Or v = "" Or _
       (v = "X" And _
         (InStr("UV", str_Vul) = 0 Or _
          str_Vul = "")) _
    Then
        ' If playing and the Vul spec is "", then   '
        ' any vulnerability is okay, but if testing '
        ' an entry, use that entry's specified vul. '
        Vul_Ok = True

    ElseIf bln_ShowAnalysis Then
        ' Test failed. '
        With frm.lstAnalysis
            .ListIndex = 13
            If str_Vul = "" Then ' any '
                v = "Any vul ok"
            ElseIf str_Vul = "B" Then ' both vul '
                v = "Both vul"
            ElseIf str_Vul = "N" Then ' neither vul '
                v = "Neither vul"
            ElseIf str_Vul = "U" Then ' Unfav '
                v = "Unfav. vul"
            ElseIf str_Vul = "F" Then ' Fav '
                v = "Favorable vul"
            ElseIf str_Vul = "V" Then ' Vul bidder '
                v = "Vul bidder"
            ElseIf str_Vul = "X" Then ' Non-vul bidder '
                v = "Bidder not vul"
            End If
            .List(13) = .List(13) & _
                "            NO MATCH:  " & v
        End With
    End If
End Function


Private Function Position_calc(c() As String)
    Dim i As Integer
    Dim posNum As Integer
    
    ' str_PriorCalls is used instead of str_PriorBids  '
    ' because we are looking for position relative to  '
    ' dealer, not to opening bidder. For example, "N"  '
    ' means player is Not a passed hand which can only '
    ' be determined from PriorCalls: P-P-1S-P-?? would '
    ' be a passed hand but PriorBids is "1S-P-??".     '
    If str_PriorCalls <> "" Then
        For i = 1 To Len(str_PriorCalls)
            If Mid$(str_PriorCalls, i, 1) = "-" Then
                posNum = posNum + 1
                ' Get passed hands: '
                If posNum < 5 And _
                   Mid$(str_PriorCalls, i - 1, 1) = "P" _
                Then
                    c(posNum) = "P"
                End If
            End If
        Next
        ' Example "1C-P-1H": posNum = 2. "1C" = 1. '
        ' posNum+2 is the # of player bidding now. '
        posNum = posNum + 2
        If posNum > 4 Then
            posNum = posNum Mod 4
            If posNum = 0 Then posNum = 4
        End If
    End If
    Position_calc = posNum
End Function


Private Function Position_Ok()
    ' Check position of current player  '
    ' (e.g.: 1st seat, 2nd, 3, 4, etc.) '
    ' AND his partner. '

    ' Remember str_PriorCalls starts    '
    ' with any leading passes and       '
    ' str_PriorBids does not...         '
    ' If str_PriorCalls is blank, then  '
    ' this player is                    '
    '    1 (first seat) or              '
    '    F (First or Second seat) or    '
    '    N (an unpassed hand)           '
    ' and partner is also unpassed (U). '

    ' Normally, whether a player is a   '
    ' passed/unpassed hand only refers  '
    ' to his action in the first round. '
    ' If str_PriorCalls has more than 3 '
    ' calls in it, then we go back to   '
    ' the 1st round to see what he did. '

    ' The variable str_Position was calculated   '
    ' in Sub Calculate_BidNum. It is all the     '
    ' abbreviations which apply to the current   '
    ' hand for which we are trying to find a bid.'
    ' It typically looks something like "1NFU".  '
    ' If the position specified in the database  '
    ' for the current bid entry is "F" then that '
    ' is one of the current hand's abbreviations '
    ' and it passes. '

    Dim position As String
    Dim problem As String
    Dim posNum As Integer
    Dim dashes As String
    Dim c(4) As String ' first 4 calls (which may be "P") '
    Dim i As Integer
    Dim pb As String ' prior bids of entry being tested '
    
'If rs_BidTable!id = 25240 Then Stop
    
    On Local Error GoTo PositionError
    position = Trim$(UCase$("" & rs_BidTable!pos))
             ' Position spec for current bid entry. '

    If position = "" Then
        Position_Ok = True
        Exit Function
    End If

    posNum = Position_calc(c())
    
    Position_Ok = False
    If posNum = 0 Then posNum = 1
    If position = "" Or position = "1" Or position = "F" Then
        If posNum = 1 Then
            Position_Ok = True
            Exit Function
        End If
        problem = "The bidder's position is not 1st chair."

    ElseIf position = "2" Then
        If posNum = 2 Then
            Position_Ok = True
            Exit Function
        End If
        problem = "The bidder's position is not 2nd chair."

    ElseIf position = "4" Then
        If posNum = 4 Then
            Position_Ok = True
            Exit Function
        End If
        problem = "The bidder's position is not 4th chair."
        'Stop
' If bidding has gone P-P-P-, then the player is in position="4" '
' but posNum calculated in "posNum=Position_calc(c())" is "0" so '
' when 4th chair is specified in the calling app but the entry   '
' does not specify that the entry is for 4th chair, it does not  '
' reject the entry. '
        

    ElseIf position = "3" Or position = "T" Then
        ' 3-15-2020: In the line above was 'or position = "N"'  '
        ' but that conflicted with 'ElseIf position = "n" below.'
        If posNum = 3 Then ' ?? Or (posNum = 4 And position <> "N") Then
            Position_Ok = True
            Exit Function
        End If
        problem = "The bidder's position is not 3rd chair." ' (or 4th) chair."

    ElseIf position = "Y" Then
        ' Y = bidder is passed '
        If c(posNum) = "P" Then
            Position_Ok = True
            Exit Function
        End If
        problem = "The bidder is not a passed hand as specified."

    ElseIf position = "N" Then
        ' N = bidder is unpassed. '
        If int_CallNum < 5 Then
            Position_Ok = True
        ElseIf str_Call(posNum) <> "P" Then
            Position_Ok = True
        End If
        Exit Function
        
        If c(posNum) = "N" Or position = "1" Then
            Position_Ok = True
            Exit Function
        End If
        problem = "The bidder is not an unpassed hand as specified."

    ElseIf position = "P" Then
        ' P = bidder's pard is passed. '
        If posNum > 2 Then
            If c(posNum - 2) = "P" Then
                Position_Ok = True
                Exit Function
            End If
        Else
            If c(posNum + 2) = "P" Then
                Position_Ok = True
                Exit Function
            End If
        End If
        problem = "Bidder's pard is not a passed hand as specified."

    ElseIf position = "U" Then ' 2020-06-13: And posNum > 2 Then
        ' U = pard is unpassed. '
        If posNum < 3 Then posNum = posNum + 2
        If c(posNum - 2) <> "P" Then
            Position_Ok = True
            Exit Function
        End If
        problem = "Bidder's pard is not an unpassed hand as specified."

    End If
    DoEvents

    If bln_ShowAnalysis Then
        If lng_TestID > 0 Then
            ' If an entry ID was specified, this is being '
            ' called by the Editor to test an entry and   '
            ' normally would not show starting passes.    '
            Position_Ok = True
            Exit Function
        End If

        ' A match was not found. '
        With frm.lstAnalysis
            .ListIndex = 14
            .List(14) = .List(14) & _
                        " -- " & problem
        End With
    End If
Exit Function

PositionError:
    MsgBox Err.Description & vbCr & _
           "in Function Position_Ok for" & vbCr & _
           "entry ID#" & rs_BidTable!id
    Stop
    Resume
    On Local Error GoTo 0
End Function


Private Function Suits_Ok() As Boolean
    '*******************************'
    ' See if Suit specs match hand. '
    ' This routine calls:           '
    '   Combo_Ok and                '
    '   Suit_Ok which calls         '
    '       SuitNum_Ok and          '
    '       SuitHCP_Ok.             '
    '*******************************'
    
    ' Say that x-Num = "6+,4+" and Spds = "x,<4" '
    ' First we test x and if it fails, then we   '
    ' test for the alternate value of Spds.      '
    Dim num1 As String
    Dim num2 As String
    Dim num3 As String
    Dim num4 As String
    Dim numC As String
    Dim pts1 As String
    Dim pts2 As String
    Dim pts3 As String
    Dim pts4 As String
    Dim ptsC As String
    Dim sOk(4) As Boolean
    Dim dbNum(4) As String
    Dim dbHCP(4) As String
    Dim s As Integer
    Dim c As String
    Dim cSuit As String
    Dim m As String
    Dim bln_xField As Boolean
    Dim i As Integer
    Dim ix As Integer
    Dim longest_xSuit As Integer
    Dim compareXsuit As Integer
    Dim compareXnum As String
    Dim compareXpts As String
    Dim cmpr As String

On Local Error GoTo Suits_Ok_Error
    
    bln_xSuitFound = True
    
    With rs_BidTable
        dbNum(4) = !SNum
        dbHCP(4) = !SHCP
        dbNum(3) = !HNum
        dbHCP(3) = !HHCP
        dbNum(2) = !DNum
        dbHCP(2) = !DHCP
        dbNum(1) = !CNum
        dbHCP(1) = !CHCP
    
'If rs_BidTable!id = 21231 Then Stop

        ' See if M or m start Combo field. '
        ' If so and if not an actual combo,'
        ' replace the value above with the '
        ' value in the Combo field.        '
        numC = Trim$("" & !cmNum)
        ptsC = Trim$("" & !cmHCP)
        'm = str_Looking4Majors
        'If m = "M" Then m = "H" ' Otherwise m = "S". '
        ' It's only M when testing entries and '
        ' both H & S are tested, starting with '
        ' H and if no bid is found, then S. '
            
        ' See if M is part of an actual Combo spec '
        ' such as "Mx=10+" = "Hx=10+" which means  '
        ' that the total HCP for H and the x suit  '
        ' is 10+. Of course "MD=10+" would be the  '
        ' total of H and D, but MH=10+ or MS=10+   '
        ' won't work because that could be H+H or  '
        ' S+S, so such an entry can't be in 1M-P.  '
        
        ' Combo Num: '
        If Left$(numC & " ", 1) = "M" Then
            c = Mid$(numC, 2, 1)
            If c = "=" Then
                ' M=H means the entry is for H only. Ditto for M=S. '
                cSuit = Mid$(numC, 3, 1)
                c = cSuit
            End If
            
            If InStr("CDHS", c) = 0 Then
                ' Simply M3+ or whatever. '
                ' Put M specs into H/S:   '
                If m = "H" Then
                    dbNum(3) = Mid$(numC, 2) ' dbNum(3) is hearts '
                    If InStr(ptsC, "M") > 0 Then
                        dbHCP(3) = Mid$(ptsC, 2)
                    End If
                Else
                    dbNum(4) = Mid$(numC, 2) ' dbNum(4) is spades '
                    If InStr(ptsC, "M") > 0 Then
                        dbHCP(4) = Mid$(ptsC, 2)
                    End If
                End If
                numC = ""
            ElseIf Left$(numC, 2) <> "M=" Then ' Mx=10+ or MC<8, etc. '
                ' This is a Combo spec. Change M to H/S and leave it. '
                numC = m & Mid$(numC, 2)
            End If
                
        ' Combo HCP: '
        ElseIf Left$(ptsC & " ", 1) = "M" Then
            c = Mid$(ptsC, 2, 1)
            If InStr("CDHSx", c) = 0 Then
                ' Simply M3+ or whatever. '
                ' Put M specs into H/S:   '
                If m = "H" Then
                    dbHCP(3) = Mid$(ptsC, 2)
                Else
                    dbHCP(4) = Mid$(ptsC, 2)
                End If
                ptsC = ""
            Else ' Mx=10+ or MC<8, etc. '
                ' This is a Combo spec. Change M to H/S and leave it. '
                ptsC = m & Mid$(ptsC, 2)
            End If
        End If
                
        If str_Looking4Majors <> "" Then
            ' Ready to rock and roll: '
            GoSub SeeIfSuitsOk
            ' Yes, GoSub is that tool of the Devil,  '
            ' but in a new Sub, I would have to pass '
            ' all the variables I've set up and I've '
            ' got bigger fish to fry at the moment.  '
            If bln_xSuitFound Then
                For i = 1 To 4
                    If Not sOk(i) And dbNum(i) = "x" Then
                        ' If dbNum has no alternative to x and '
                        ' some x suit is ok, then this is Ok.  '
                        sOk(i) = True
                    End If
                Next
            End If
            
            If sOk(1) And sOk(2) And sOk(3) And sOk(4) Then
                ' Remember (1)=C and (4)=S. '
                ' Suits are ok, check Combo: '
                If Combo_Ok(numC, ptsC) Then
                    Suits_Ok = True
                    Exit Function
                Else
                    str_xSuit = ""
                    bln_xSuitFound = False
                    Suits_Ok = False
                End If
            End If
        End If
        
        ' No bid found in the 1M-P (etc.) sections. '
        ' Reset variable & look in 1H/S-P sections: '
    ' If not an "M" entry, from here down is '
    ' the first testing that will be done.   '
RerunFunction:
        dbNum(3) = !HNum
        dbHCP(3) = !HHCP
        dbNum(4) = !SNum
        dbHCP(4) = !SHCP
        ' C & D are still the same. '
        If str_Looking4Majors = "M" Then
            str_Looking4Majors = "S" ' H was just checked. '
            i = Suits_Ok
        End If
'Exit Function
'If rs_BidTable!id = 24445 Then Stop

        GoSub SeeIfSuitsOk ' Doesn't check Combined suits specs. '
        ' dbNum() was getting cleared somewhere, so setting it up again: '
        dbNum(4) = !SNum
        dbHCP(4) = !SHCP
        dbNum(3) = !HNum
        dbHCP(3) = !HHCP
        dbNum(2) = !DNum
        dbHCP(2) = !DHCP
        dbNum(1) = !CNum
        dbHCP(1) = !CHCP
        For i = 1 To 4
            If Not sOk(i) Then
                If InStr(dbNum(i), "x") > 0 And InStr(dbNum(i), ",") = 0 Then
                    ' H: x,2-4  D:x,2-4 -- If H matched x but D didn't, '
                    ' the entry is okay but when D doesn't match x and  '
                    ' an alternative is specified but not matched, then '
                    ' the entry is no good. '
                    If bln_xSuitFound Then
                        sOk(i) = True
                    End If
                End If
            End If
        Next
        If sOk(1) And sOk(2) And sOk(3) And sOk(4) And _
           (bln_xSuitFound Or !xNum = "") _
        Then ' Above: either an x-suit was found or no x field specified. '
            ' Suits are ok, check Combo specs: '
            If Combo_Ok(numC, ptsC) Then
                If longest_xSuit > 0 Then
                    str_xSuit = Mid$("CDHS", longest_xSuit, 1)
                End If
                Suits_Ok = True
            End If
        Else
            str_xSuit = ""
            bln_xSuitFound = False
            Suits_Ok = False
        End If
    End With
Exit Function
        
        
SeeIfSuitsOk:
''''''''''''
    For s = Clbs To Spds
        ' See if trump suit was specified: '
        If InStr(dbHCP(s), "Trmp") > 0 Then
            int_TrumpSuit = s
            If dbHCP(s) = "Trmp" Then
                dbHCP(s) = ""
            Else
                dbHCP(s) = Replace(dbHCP(s), "Trmp", "")
                dbHCP(s) = Replace(dbHCP(s), ",", "")
            End If
        End If
        ' Get rid of blanks: '
        dbNum(s) = Replace(dbNum(s), " ", "", , 99)
        dbHCP(s) = Replace(dbHCP(s), " ", "", , 99)
    Next

    str_xSuit = ""
    bln_xField = False
    bln_xSuitFound = False ' Are both needed? '

    Dim s1 As Integer
    Dim s2 As Integer
    Dim s3 As Integer
    Dim xNum As String

'If rs_BidTable!id = 20066 Then Stop

    xNum = Trim$("" & rs_BidTable!xNum)
    ' xNum is Num for the "x" field: "1+,<2" '
    ' dbNum(s) could be "x" and so would be  '
    ' tested for "1+" and then for "<2".     '
    
    ' If xNum has "U" in it, go up the suits '
    ' from Clubs to Spades, else go down.    '
    If Right$(" " & xNum, 1) = "U" Then
        s1 = Clbs
        s2 = Spds
        s3 = 1 'step increment in For statement below. '
        If xNum <> "" Then xNum = Left$(xNum, Len(xNum) - 1)
    Else
        s1 = Spds
        s2 = Clbs
        s3 = -1
    End If
    
    For s = s1 To s2 Step s3
        num1 = dbNum(s) ' This is Num for suit(s).
        num2 = ""
        num3 = ""
        num4 = ""
        pts1 = dbHCP(s)
        pts2 = ""
        pts3 = ""
        pts4 = ""

' Example: Test Diamonds for a control, '
' which is a A, K, singleton or void:   '
'                                       '
'     num HCP                           '
' x    1+,<2   1/2 [1 of top 2]         '
' Dia    x                              '
'                                       '
' If Num Dia = 1+ and HCP= A or K then  '
'     Suit is okay.                     '
' ElseIf Dia Num < 2 then               '
'     Dia must be singleton or void     '
'     Suit is okay.                     '
' End If                                '

        If LCase$(num1) = "x" Or LCase$(pts1) = "x" Then
            ' Change num1 from "x" to xNum's value: '
            num1 = xNum
            pts1 = Trim$("" & rs_BidTable!xHCP)
            ' If num1 for a suit is something like   '
            ' "x,3+" with Points of, say, "ab", then '
            ' the points go with the 3+ since x has  '
            ' its own points field. (Bears repeating)'
        End If

        If LCase$(Left$(dbNum(s) & " ", 1)) = "x" Then
            If InStr(dbNum(s), ",") > 0 Then
                If LCase$(Left$(dbHCP(s) & " ", 1)) <> "x" Then
                    If Len(dbHCP(s)) > 0 Then
                        If InStr(dbHCP(s), ",") = 0 Then
                            ' e.g.: If dbNum = "x,3-5" and dbHCP = 4+ '
                            ' then 4+ must go with "3-5", not with x  '
                            ' because "x" has its own HCP field.      '
                            ' If Num= 4,3-5 and Points= 5 then 5 goes '
                            ' with 4, not with 3-5. '
                            pts1 = " ," & pts1
                        End If
                    End If
                End If
            End If
        End If

        ' Split alternatives, if any: '
        i = InStr(num1, ",")
        ix = InStr(xNum, ",")
        If i > 0 Then
            If ix > 0 And i <> ix Then
                num3 = Trim$(Mid$(num1, ix + 1))
                num2 = Trim$(Left$(num1, ix - 1))
            Else
                num2 = Trim$(Mid$(num1, i + 1))
            End If
            'num1 = Trim$(Left$(num1, i - 1))
            ' Changed line above to the one below '
            ' because the code was not working. '
            If InStr(num1, "x") > 0 Then
                If ix = 0 Then ' e.g.: x '
                    num1 = xNum
                Else ' e.g.: x,<2 '
                    num1 = Trim$(Left$(xNum, i - 1))
                End If
            Else ' e.g.: 5,6 '
                num1 = Left$(num1, i - 1)
            End If
            If num1 = "" Then num1 = num2
        End If

        i = InStr(pts1, ",")
        If i > 0 Then
            pts2 = Trim$(Mid$(pts1, i + 1))
            pts1 = Trim$(Left$(pts1, i - 1))
        End If

        If num1 = "" And pts1 = "" Then
            sOk(s) = True
            GoTo GetNextSuit
        End If
        
            Do
                sOk(s) = Suit_Ok(s, num1, pts1)
                'If s = 2 Then Stop
                
                If sOk(s) = False Then
                    ' Say that 3 suits have "x". Only 1 has to meet '
                    ' the x specs for all the suits to be ok. '
                    If num2 <> "" Then
                        ' Check alternative num spec: '
                        num1 = num2
                        i = InStr(num1, ",")
                        If i > 0 Then
                            num2 = Trim$(Mid$(num1, i + 1))
                            num1 = Trim$(Left$(num1, i - 1))
                        Else
                            num2 = ""
                        End If

                        pts1 = pts2
                        
                        sOk(s) = Suit_Ok(s, num1, pts1)

                    ElseIf pts2 <> "" Then
                        ' Check alternative HCP spec: '
                        pts1 = pts2
                        i = InStr(pts1, ",")
                        If i > 0 Then
                            pts2 = Trim$(Mid$(pts1, i + 1))
                            pts1 = Trim$(Left$(pts1, i - 1))
                        Else
                            pts2 = ""
                        End If

                    Else
                        If LCase$(dbNum(s)) = "x" And _
                           (dbHCP(s) = "" Or _
                            LCase$(dbHCP(s)) = "x") _
                        Then
                            ' If x, it is not an error '
                            ' unless ALL x suits fail. '
                            'sOk(s) = True

                        ElseIf bln_ShowAnalysis Then
                            With frm.lstAnalysis
                                'If Not frm.Visible Then
                                '    ShowAnalysis
                                'End If
                                frm.Visible = True
                                i = 23 - s
                                .ListIndex = i
                                .List(i) = .List(i) & _
                                    "            NO MATCH:  Suit = " & _
                                    Chr$(34) & str_Suit(s) & Chr$(34)
                            End With
                        End If
                        Exit Do
                    End If

                Else ' sOk(s) = True '
                
                    If InStr(dbNum(s), "x") = 1 And dbNum(s) <> "x," & num1 Then
                        str_xSuit = Mid$("CDHS", s, 1)
                        If InStr(xNum, "<") > 0 Then ' e.g.: "<2" '
                            ' If we are looking for a short suit, we '
                            ' can quit now that we have found it.    '
                            For i = s1 + s3 To s2 Step s3
                                dbNum(i) = Replace(dbNum(i), "x", "")
                            Next
                            Exit Do
                        End If
                        
                        If longest_xSuit = 0 Or _
                           int_SuitLen(s) > int_SuitLen(longest_xSuit) _
                        Then
                            longest_xSuit = s
                        End If

                    ElseIf LCase$(Right$(" " & num1, 1)) = "x" Or _
                           LCase$(Right$(" " & pts1, 1)) = "x" Then
                        ' This doesn't make sense. num1 is set to whatever '
                        ' is in the x's Num box, so num1 can't be "x".     '
                        ' Also, Points should never have "x" in it any more'
                        ' unless num1 is something like "x,4+", then pts1  '
                        ' is for the 4+, not for the x. '
                        Stop
                        compareXsuit = s
                        compareXnum = num1
                        compareXpts = pts1

                    End If

                    Exit Do

                End If

                DoEvents
           Loop

        'End If
        
        If str_xSuit <> "" Then
            bln_xSuitFound = True
        End If
        
        If sOk(s) = False Then
            ' Exiting when s = 3 and s=2 is true. '
            ' But we do want to exit if a spec is false, "
            ' so this needs to be fixed. '
'            Stop
            If dbNum(s) <> "x" Then ' I think this fixes the above. '
                ' If it is "x", then we have to keep testing the other suits. '
                Exit For
            End If
        End If

GetNextSuit:
    ' Even if this suit fit x, we may need to '
    ' look for a longer suit, for example.    '

    Next
    
    If longest_xSuit > 0 Then
        str_xSuit = Mid$("CDHS", longest_xSuit, 1)
    End If

    'See if xNum or xHCP had a value not met by any suit
    If str_xSuit = "" And _
       (Trim$(rs_BidTable!xNum) <> "" Or _
        Trim$(rs_BidTable!xHCP) <> "") _
    Then
        bln_xSuitFound = False
        If bln_ShowAnalysis Then
            With frm.lstAnalysis
                .ListIndex = 18
                .List(18) = .List(18) & _
                "            NO MATCH:  Doesn't match 'x' specs."
            End With
        End If
        Return 'Exit Function
    End If

    If compareXsuit > 0 And str_xSuit <> "" Then
        s = InStr("CDHS", str_xSuit)
        Dim n1 As Integer
        Dim n2 As Integer
        If compareXnum <> "" Then
            n1 = int_SuitLen(compareXsuit)
            n2 = int_SuitLen(s)
            If Len(compareXnum) = 3 Then
                cmpr = Left$(compareXnum, 2) 'e.g.: <=x
            Else
                cmpr = Left$(compareXnum, 1) 'e.g.: >x
            End If
            If cmpr = "<" Then
                sOk(compareXsuit) = (n1 < n2)
            ElseIf cmpr = ">" Then
                sOk(compareXsuit) = (n1 > n2)
            ElseIf cmpr = "=>" Or cmpr = ">=" Then
                sOk(compareXsuit) = (n1 >= n2)
            ElseIf cmpr = "=<" Or cmpr = "<=" Then
                sOk(compareXsuit) = (n1 <= n2)
            End If
        End If

        If compareXpts <> "" Then
            n1 = int_HCP(compareXsuit)
            n2 = int_HCP(s)
            If Len(compareXpts) = 3 Then
                cmpr = Left$(compareXpts, 2) 'e.g.: <=x
            Else
                cmpr = Left$(compareXpts, 1) 'e.g.: >x
            End If
            If cmpr = "<" Then
                sOk(compareXsuit) = (n1 < n2)
            ElseIf cmpr = ">" Then
                sOk(compareXsuit) = (n1 > n2)
            ElseIf cmpr = "=>" Or cmpr = ">=" Then
                sOk(compareXsuit) = (n1 >= n2)
            ElseIf cmpr = "=<" Or cmpr = "<=" Then
                sOk(compareXsuit) = (n1 <= n2)
            End If
        End If
    End If

Return

Suits_Ok_Error:
    MsgBox Err.Description & vbCr & _
           "Error in Function Suits_Ok."
    ' In edit mode, just Resume to find the error. '
    Stop
    Resume
    On Local Error GoTo 0
Return

End Function


Private Function Combo_Ok(num1 As String, pts1 As String) As Boolean
    ' Check combo specs, if any. '
    ' Example: pts1 = "SHD>9"    '

    Dim comboOk As Boolean
    Dim s1 As Integer ' 1st suit in spec '
    Dim s2 As Integer ' 2nd suit in spec '
    Dim s3 As Integer ' 3rd suit in spec '
    Dim sign As String
    Dim sTtl As Integer
    Dim sPts As Integer
    Dim ttl As Integer
    Dim ttl2 As Integer
    Dim dash As Integer
    'Dim num1 As String
    'Dim pts1 As String
    Dim numOk As Boolean
    Dim numKC As Integer
    Dim gerber As Boolean
    Dim numAces As Integer
    Dim numKeys As Integer
    Dim ref As String
    Dim t As String
    Dim i As Long
    Dim j As Long
    Dim signK As String
    Dim signQ As String
    Dim c As Long ' comma position '
    Dim p As String ' p = a temp variable for pts1 '
    Dim s As String ' s = an individual spec within pts1 '

    ' Check combined suits specs: '
    'num1 = Trim$("" & rs_BidTable!cmNum) ' "" &  prevents an error if cmNum is null. '
    'pts1 = Trim$("" & rs_BidTable!cmHCP)
'If rs_BidTable!id = 17341 Then Stop

On Local Error GoTo Combo_Ok_Error

    If num1 = "" And pts1 = "" Then
        Combo_Ok = True
        Exit Function
    End If

    ' Check suit lengths:                           '
    If Len(num1) > 3 Then
        ' Combining quantities:          '
        ' (e.g.: SC>8, CHD=11-13, DC<4 ) '
        s1 = InStr("CDHS", Mid$(num1, 1, 1))
        If s1 > 0 Then
            s2 = InStr("CDHS", Mid$(num1, 2, 1))
            s3 = InStr("CDHS", Mid$(num1, 3, 1))

            If s3 > 0 Then ' CDH=10-12, DHS>9, etc.
                sign = Mid$(num1, 4, 1)
                sTtl = int_SuitLen(s1) + _
                       int_SuitLen(s2) + _
                       int_SuitLen(s3)
                ttl = Val(Mid$(num1, 5))

            Else ' CD=8-9, SH>8, etc.
                sign = Mid$(num1, 3, 1)
                sTtl = int_SuitLen(s1) + int_SuitLen(s2)
                ttl = Val(Mid$(num1, 4))
            End If

            ' Do the math: '
            If sign = "=" Then ' CH=8, CDH=11-12, etc.
                dash = InStr(num1, "-")
                If dash > 0 Then ' range of qty. specified. '
                    ttl2 = Val(Mid$(num1 + "  ", dash + 1))
                    If sTtl < ttl Or _
                       sTtl > ttl2 _
                    Then
                        GoTo ExitCombo
                    End If
                ElseIf sTtl <> ttl Then
                    GoTo ExitCombo
                End If

            ElseIf sign = "<" Then 'CD<4, CDH<9
                If sTtl >= ttl Then GoTo ExitCombo

            ElseIf sign = ">" Then ' SH>8, DHS>11
                If sTtl <= ttl Then GoTo ExitCombo
            End If

        End If
    End If

    numOk = True
    If pts1 = "" Then
        comboOk = numOk
        GoTo ExitCombo
    End If

' Slam Bidding specs example:             '
' BidBase entry specs:                    '
'     4N-P-5S: 2 or 4 keys with trump Q   '
'     2A-TK+TQ,1A+TK+TQ,4A-TK+TQ,3A+TK+TQ '
' User:        ^^^^^^^^                   '
'     KQ983-T4-A94-QJ3                    '
'     ^^       ^      = 1A+TK+TQ          '

    If InStr(pts1, "T") = 0 And _
       InStr(pts1, "A") = 0 And _
       InStr(pts1, "K") = 0 And _
       InStr(pts1, "Q") = 0 _
    Then
        ' No slam bidding characters in the Combo fields. '
        GoTo SlamBiddingDone
    End If

    ' Specs are in the Suits - Combo-HCP box. '
    ' Like other entries, each entry has a    '
    ' test hand for verifying that the specs  '
    ' do what is expected. '
    
    ' User's combo variables are calculated in '
    '  Sub Calculate_CustPoints                '
    
    ' If basic Grand Slam Force is being used: '
    ' Suits - Comb.-HCP should have "TA+TK,TA+TQ,TK+TQ". '
    ' This specifies 2 of the top 3 trumps. '

    'If rs_BidTable!id = 17341 Then Stop

    pts1 = Replace(pts1, " ", "")
    ' RKCB 1430: '
    ' 5C: 1 or 4 keys = 1A-TK,0A+TK,4A-TK,3A+TK. No void if 4 keys which would be 5N. '
    ' 5D: 0 or 3 keys = 0A-TK,3A-TK,2A+TK. No void if 3 keys which would be 6x. '
    ' 5H: 2 or 5 keys = 2A-TK-TQ,1A+TK-TQ,4A+TK-TQ '
    ' 5S: 2 or 5 keys = 2A-TK+TQ,1A+TK+TQ,4A+TK+TQ '
    ' 5N: 2 or 4 keys = 2A-TK-TQ,1A+TK-TQ,3A+TK + some void, so can't have 4 aces. '
    ' 6C: 1 or 3 keys = 1A-TK,0A+TK,4A-TK,3A+TK + C void '
    '     If C's are trumps, a void can't be shown - takes you past 6C. '
    ' 6D: 1 or 3 keys = 1A-TK,0A+TK,4A-TK,3A+TK + D void '
    '     If D's are trumps, 6C shows C void, 6D shows H/S void. '
    ' 6H: 1 or 3 keys = 1A-TK,0A+TK,4A-TK,3A+TK + H void '
    '     If H's are trumps, 6C/6D = void suit bid, 6H shows S void. '
    
    '    int_MyNumAces  '
    '    int_MyNumKings '
    '    bln_MyTrumpK   '
    '    bln_MyTrumpQ   '
    '    int_MyVoidSuit '
    Dim nA As Integer
    
    s = int_MyNumAces & "A"
    
    If bln_MyTrumpK Then
        s = s & "+TK"
    Else
        s = s & "-TK"
    End If
    
    If InStr(pts1, s & "-TQ") Then
        If Not bln_MyTrumpQ Then
            comboOk = True
        End If
    ElseIf InStr(pts1, s & "+TQ") Then
        If bln_MyTrumpQ Then
            comboOk = True
        End If
    ElseIf InStr(pts1, s) Then
        comboOk = True
    End If
    
    GoTo ExitCombo ' Just doing 4N Ace asking '


    ' Count Kings: '

    ' King counting is done for Ace asker after   '
    ' pard has shown his number of Aces. '
        i = InStr(pts1, "K")
    If i > 1 Then
        If Val(Mid$(pts1, i - 1, 1)) > 0 Then
            ' Has to be some # in front of K, so count 'em: '
            int_NumKings = 0
            For s1 = 1 To 4
                If bln_Suit(s1, 13) And _
                    Not (bln_TrumpKing And _
                         s1 = int_TrumpSuit And _
                         InStr(pts1, "-TK") > 0) _
                Then
                    int_NumKings = int_NumKings + 1
                End If
            Next
            If InStr(pts1, int_NumKings & "K") > 0 Then
                comboOk = True
                GoTo ExitCombo
            End If
        End If
    End If

    ' Each specification is separated by a comma.      '
    ' Unlike individual suits Combos may have >2 specs.'
    ' If one spec passes, then Combo is ok; otherwise, '
    ' try the next spec until done. Example used here: '
    ' pts1 = 2A+Q,1A+TK+Q,4A+TK+Q     This is the spec '
    ' for the bid 5S = 2 or 5 key cards plus trump Q.  '
    ' There are 3 specs: 2A+Q - 1A+TK+Q - 4A+TK+Q.     '
    ' Below we check them one at a time in a loop:     '

    p = pts1 ' Example = 2A+Q,1A+TK+Q,4A+TK+Q '
                       ' 123456... '
    Do
        DoEvents
        c = InStr(p, ",") ' c=6 in the example shown above. '
        If c = 0 Then
            ' There is only one spec. '
            s = p
            p = ""
        Else
            s = Left$(p, c - 1) ' 1st spec from example is 2A+Q. '
            ' Remove the spec currently being processed: '
            p = Mid$(p, c + 1)  ' 1A+TK+Q,4A+TK+Q remains. '
            ' Next loop s = 1A+TK+Q and p = 4A+TK+Q '
            ' Then s = 4A+TK+Q and p = "". '
        End If

        ' 2A+Q is the same as 2A-TK+Q as is 1A+TK and '
        ' 1A+TK-Q, etc., so we must adjust for that.  '
        s = Replace$(s, "-Q", "-TQ")
        If Right$(s, 1) = "A" And Len(s) = 2 Then
            ' i.e.: s = "#A" '
            s = s & "-TK-TQ"
        ElseIf Right$(s, 1) = "K" And Len(s) = 5 Then
            ' i.e.: s = "#A+TK" or "#A-TK" '
            s = s & "-TQ"
        ElseIf Right$(s, 1) = "Q" And Len(s) = 5 And _
               InStr(s, "K") = 0 _
        Then
            ' i.e.: s = "#A+TQ" or "#A-TQ" '
            s = Left$(s, 2) & "-TK" & Right$(s, 3)
        End If
        ' Now the "specs" include #A and TK and TQ. '
        ' If just "A" was entered with no number: '
        If s = "A" Then s = "1A"

        ' Look for # Aces matching the spec: '
        If InStr(s, numAces & "A") = 0 Then
            ' Example: "1A+TK+TQ" and # Aces=2. '
            ' #A held does not match the spec.  '
            GoTo NextSpec
        End If

        ' # Aces match; now '
        ' check trump King. '
        If InStr(s, "-TK") Then
            If bln_TrumpKing Then
                ' "-TK" is specified and but TK is held. '
                GoTo NextSpec
            End If
        ElseIf InStr(s, "A+TK") Then ' "...A+TK..." '
            If Not bln_TrumpKing Then
                GoTo NextSpec
            End If
        End If

        ' #A and TK match.   '
        ' Check trump Queen: '
        If InStr(s, "-TQ") Then
            If bln_TrumpQueen Then
                ' "-TQ" is specified and but TQ is held. '
                GoTo NextSpec
            End If
        ElseIf InStr(s, "+TQ") Then ' "...+TQ..." '
            If Not bln_TrumpQueen Then
                GoTo NextSpec
            End If
        End If
        ' Spec checks out; now exit: '
        comboOk = True
        GoTo ExitCombo

NextSpec:
        DoEvents
    Loop While p <> ""

SlamBiddingDone:

    ' Examples for the below: SH>5, HDC>10, Hx<9
    s1 = InStr("CDHSx", Left$(pts1 & " ", 1))
    If Len(pts1) > 3 And s1 <> 0 Then
        If s1 = 5 Then
            ' Get x spec, if any: '
            s1 = InStr(" CDHS", str_xSuit) - 1
            If s1 < 1 Then s1 = 0
        End If

        sPts = 0
        ttl = 0
        s2 = 0
        s3 = 0

        If s1 > 0 Then
            s2 = InStr("CDHSx", Mid$(pts1 & "   ", 2, 1))
            If s2 = 5 And str_xSuit <> "" Then
                s2 = InStr(" CDHS", str_xSuit) - 1
                If s2 < 1 Then s3 = 0
            ElseIf s2 > 0 Then
            End If
            
            s3 = InStr("CDHS", Mid$(pts1 & "    ", 3, 1))
            If s3 > 0 Then 'new code:  = 5 And str_x_Suit <> "" Then
'                s3 = InStr(" CDHS", str_xSuit) - 1
'                If s3 < 1 Then s3 = 0
'            End If
'            If s3 = 5 Then ' CDH=10-12, DHS>9, etc.
                sign = Mid$(pts1 & "     ", 4, 1)
                sPts = int_HCP(s1) + _
                       int_HCP(s2) + _
                       int_HCP(s3)
                ttl = Val(Mid$(pts1 & "      ", 5))

            Else ' CD=8-9, SH>8, etc.
                sign = Mid$(pts1 & "    ", 3, 1)
                sPts = int_HCP(s1) + _
                       int_HCP(s2)
                ttl = Val(Mid$(pts1 & "     ", 4))
            End If
            
            If num1 = "" And pts1 = "" Then
                comboOk = True
            End If

            'Do the math:
            If sign = "=" Then
                dash = InStr(num1, "-")
                If dash > 0 Then ' CD=6-9 or CDH=12-15 '
                    ttl2 = Val(Mid$(num1 + "  ", dash + 1))
                Else ' e.g.: CD=6+ or CDH=12+ '
                    ttl2 = ttl
                End If

                If dash > 0 Then
                    If sPts < ttl Or sPts > ttl2 Then
                        comboOk = False
                        GoTo ExitCombo
                    End If
                ElseIf sPts <> ttl Then
                    comboOk = False
                    GoTo ExitCombo
                End If

            ElseIf sign = "<" Then
                If sPts >= ttl Then
                    comboOk = False
                    GoTo ExitCombo
                End If

            ElseIf sign = ">" Then
                If sPts <= ttl Then
                    comboOk = False
                    GoTo ExitCombo
                End If
            End If
            comboOk = True
        End If
    End If

ExitCombo:
'''''''''
    If comboOk Then
        Combo_Ok = True

    ElseIf bln_ShowAnalysis Then
        With frm.lstAnalysis
            .ListIndex = 23

            If InStr(pts1, "A") > 0 Then
                If InStr(pts1, "TK") > 0 Then
                    ref = "# Keys = " & numKeys
                Else
                    ' If TK not specified, only asking for aces: '
                    ref = "# Aces = " & int_NumAces
                End If
            ElseIf InStr(pts1, "K") > 0 Then
                ref = "# Kings = " & int_NumKings
            Else
                ref = "Suit quantity and HCPs = " & sTtl & ", " & sPts
            End If

            ref = "    NO Match"
            .List(23) = .List(23) & _
            "" & _
            ref
        End With
    End If
Exit Function

Combo_Ok_Error:
    MsgBox Err.Description & vbCr & _
           "Error in Function Combo_Ok."
    Stop
    ' Use Resume to see error. '
    Resume
    On Local Error GoTo 0

End Function


Private Function Suit_Ok( _
            ByVal s As Integer, _
            ByVal num As String, _
            ByVal HCPs As String) _
        As Boolean

    ' If "sNum" or "sHCP" had an "x" spec or '
    ' alternative specs, they were replaced  '
    ' with a single value in the Suits_Ok    '
    ' function which calls this function.    '
    Dim orHCP As String
    Dim qOk As Boolean ' Quantity passes? '
    Dim pOk As Boolean ' HCPs passes?     '

On Local Error GoTo Suit_Ok_Error
    
    qOk = True
    pOk = True
    bln_xField = False
    
'If rs_BidTable!id = 21219 Then Stop 'This is ID#, not BON. '
    
    If num <> "" And LCase$(Right$(num, 1)) <> "x" Then
        ' Check suit quantity: '
        
        qOk = SuitQty_Ok(s, num)
        
        If qOk And HCPs = "" Then
            Suit_Ok = True
            Exit Function
            
        ElseIf qOk Or LCase(Left$(HCPs & " ", 2)) = "or" Then
            HCPs = Replace(HCPs, "OR", "or")
            
            If (HCPs = "" Or num = "" Or qOk = True Or _
                (Left$(HCPs, 2) = "or") And _
                 Len(HCPs) > 2) _
            Then ' And LCase$(Right$(HCPs, 1)) <> "x" Then
                
                If Left$(HCPs, 2) = "or" Then
                    orHCP = Mid$(HCPs, 3)
                Else
                    orHCP = HCPs
                End If
                ' Check suit HCPs:  '
                pOk = SuitHCP_Ok(s, orHCP)
                
                If (Left$(HCPs, 2) <> "or" And _
                    (qOk = False Or pOk = False)) Or _
                   (Left$(HCPs, 2) = "or" And _
                    qOk = False And _
                    pOk = False) _
                Then
                    Suit_Ok = False
                    Exit Function
                End If
            End If
            Suit_Ok = True
            Exit Function
        End If
    End If
Exit Function

Suit_Ok_Error:
    MsgBox Err.Description & vbCr & _
           "Error in Function Suit_Ok."
    Stop
    ' Use Resume to see the error. '
    Resume
    On Local Error GoTo 0
    
End Function


Private Function SuitQty_Ok( _
            ByVal s As Integer, _
            ByVal num As String) _
        As Boolean
    'Check number of cards in suit:
'If rs_BidTable!id = 19084 And s = 1 Then Stop
    Dim nm As Integer
    Dim nm2 As Integer
    Dim s2 As Integer
    Dim signR As String
    Dim signL As String
    Dim X As String
    Dim i As Integer

If InStr("MH", str_Looking4Majors) > 0 And str_Looking4Majors > "" Then
    i = 3
Else
    i = 4
End If
'If rs_BidTable!id = 22441 And s = 3 Then Stop
    
    ' If saving an entry with M in the Combo field, '
    ' this Function may be called 4 times for H/S:  '
    ' 1. M=H, s=4 '
    ' 2.      s=3 '
    ' 3. M=S, s=4 '
    ' 4.      s=3 '
    
    On Local Error GoTo SuitQtyErrorTrap

    nm = Val(num)
    signR = Right$(num, 1)
    signL = Left$(num, 1)
    If Val(signL) = 0 And Len(num) > 0 Then
        nm = Val(Mid$(num, 2))
    End If
    
    If num = "0" Then
        If int_SuitLen(s) = 0 Then
            SuitQty_Ok = True
        End If

    ElseIf signR = "+" Then ' e.g.:  4+   '
        If int_SuitLen(s) >= nm Then
            SuitQty_Ok = True
        End If

    ElseIf signR = "-" Then ' e.g.: 4-   '
        If int_SuitLen(s) < nm Then
            SuitQty_Ok = True
        End If

    ElseIf signL = "<" Then
        ' e.g.:  <=4  or  <=H  or  <6  or  <S '

        If Left$(num, 2) = "<=" Then
            ' e.g.:  <=4  or  <=H  '
            s2 = InStr("CDHS", Mid$(num, 3, 1))
            If s2 = 0 Then
                ' Number specified, not suit. '
                ' e.g.:  <=4    '
                If int_SuitLen(s) <= _
                   Val(Mid$(num, 3, 1)) _
                Then
                    SuitQty_Ok = True
                End If

            Else ' Suit specified. '
                 ' e.g.:  <=H      '
                If int_SuitLen(s) <= _
                   int_SuitLen(s2) _
                Then
                    SuitQty_Ok = True
                End If
            End If

        Else  ' Just <  not  <=   '
              ' e.g.:  <4  or  <H '
            s2 = InStr("CDHS", Mid$(num, 2, 1))
            If s2 = 0 Then
                ' e.g.:  <4    '
                If int_SuitLen(s) < _
                   Val(Mid$(num, 2, 1)) _
                Then
                    SuitQty_Ok = True
                End If
            Else
                ' e.g.:  <H    '
                If int_SuitLen(s) < _
                   int_SuitLen(s2) _
                Then
                    SuitQty_Ok = True
                Else
                    SuitQty_Ok = False
                End If
            End If
        End If

    ElseIf signL = ">" Then
        ' e.g.:  >=4  or  >=H  or  >6  or  >S '

        If Left$(num, 2) = ">=" Then
            ' e.g.:  >=4  or  >=H   '
            s2 = InStr("CDHS", Mid$(num, 3, 1))

            If s2 = 0 Then
                ' e.g.:  >=4    '
                If int_SuitLen(s) >= _
                   Val(Mid$(num, 2, 1)) _
                Then
                    SuitQty_Ok = True
                End If

            Else
                ' e.g.:  >=H    '
                If int_SuitLen(s) >= _
                   int_SuitLen(s2) _
                Then
                    SuitQty_Ok = True
                End If
            End If

        Else ' Just  >  not >=    '
             ' e.g.:  >4  or  >H  '
            s2 = InStr("CDHS", Mid$(num, 2, 1))

            If s2 = 0 Then
                ' e.g.:  >4     '
                If int_SuitLen(s) > _
                   Val(Mid$(num, 2, 1)) _
                Then
                    SuitQty_Ok = True
                End If
            Else
                ' e.g.:  >H     '
                If int_SuitLen(s) > _
                   int_SuitLen(s2) _
                Then
                    SuitQty_Ok = True
                End If
            End If
        End If

    ElseIf signL = "=" Then
        s2 = InStr("CDHS", Mid$(num, 2, 1))

        If s2 = 0 Then
            ' e.g.:  =4    '
            If int_SuitLen(s) = _
               Val(Mid$(num, 2, 1)) _
            Then
                SuitQty_Ok = True
            End If

        Else
            ' e.g.:  =H    '
            If int_SuitLen(s) = _
               int_SuitLen(s2) _
            Then
                SuitQty_Ok = True
            End If
        End If

    ElseIf InStr(num, "-") > 0 Then
        ' e.g.:  2-4    '
        nm2 = Val(Mid$(num, InStr(num, "-") + 1))
        If nm2 = 0 Then
            ' May be like "2-H": min=2, max=# of Hearts
            X = Mid$(num, InStr(num, "-") + 1)
            i = InStr("CDHS", X)
            If i > 0 Then
                nm2 = int_SuitLen(i)
            End If
        End If

        If int_SuitLen(s) >= nm And _
           int_SuitLen(s) <= nm2 _
        Then
            SuitQty_Ok = True
        End If

    ElseIf num <> "" And nm > 0 Then
        ' e.g.:  4
        If int_SuitLen(s) = nm Then
            SuitQty_Ok = True
        End If

    End If
    
Exit Function

SuitQtyErrorTrap:
    MsgBox Err.Description & vbCr & _
           "Error in Function SuitQty_Ok." & vbCr & _
           "Ctrl-Break to see the error. "
    Resume 'Next
    On Local Error GoTo 0
End Function


Private Function SuitHCP_Ok( _
            ByVal s As Integer, _
            ByVal HCPs As String) _
        As Boolean
    ' Check strength of suit.  '
    ' Format of HCPs can be:   '
    '   3 -- specific number   '
    '        0 = void          '
    '   3+                     '
    '   >2 -- same as 3+       '
    '   <4                     '
    '   2-4 -- range           '
    '   A -- Ace. 5+A=AJ,AQ,AK '
    '     aA = AK, AQ, not KQ  '
    '     bA = AJT only, etc.  '
    '   K/k -- King            '
    '   Q/q -- Queen           '
    '   A,K or A,K,Q or K,Q    '
    '      are also allowed.   '
    '   <A -- no Ace.          '
    '   <K -- no Ace or King.  '
    '   or... -- e.g.: or<4    '
    '   >H -- more than hearts '
    '   >C, >D, <S -- as above '
    '   ---------------------  '
    '   a, b, c were replaced  '
    '   with H## but the a,b,c '
    '   checks were left in    '
    '   for old entries.       '
    '   a -- AK, KQ, AQ        '
    '   b -- AJT, KJT, QJT     '
    '   c -- KQJ, AQJ, AKQ     '
    '   H## -- # of honors     '
    '   SM -- Solid Minor      '
    '   SS -- Solid Suit       '
    '   ---------------------  '
    '   2stop -- 2+ stoppers   '
    '   1stop -- 1+ stopper(s) '
    '   hstop -- half+ stop(s) '
    '   0stop -- no stoppers   '
    '   FRC -- 1st rnd. ctrl.  '
    '          void or Ace     '
    '   SRC -- 2nd rnd. ctrl.  '
    '   i2 -- 2+ intermediates '
    '   4+i2 -- 4+ HCPs  and   '
    '        2+ intermediates  '
    '   P -- Protected honors  '
    '     RHO or pard bid suit '
    '   s1>s2 -- suit1>suit2   '
    '     ex: C>D, S>C, etc.   '
    '   x/y -- x of top y      '
    '     3/5 = QJT to AKQ     '
    '     3/5<A = no ace: KJT  '
    ' In Comb. Points field:   '
    '   kc1,4 -- RCKB 1 or 4   '
    '   kc0,3 --      0 or 3   '
    '   kc2,5+Q with trump Q   '
    '   kc2,5-Q w/o trump Q    '

    Dim ok As Boolean
    Dim ok2 As Boolean
    Dim originalHCPs As String
    Dim ptsNeeded As Integer
    Dim left1 As String
    Dim left2 As String
    Dim left3 As String
    Dim leftSuit(5) As String
    Dim iPtr As Integer
    Dim intermedReqd As Integer
    Dim honorsList As String
    Dim honCombo As String
    Dim signL As String
    Dim signR As String
    Dim extra As Boolean
    Dim i As Integer
    Dim h As Integer
    Dim j As Integer
    Dim n As Integer
    Dim t As Integer
    Dim X As String
    Dim qtyOrHCP As Boolean

'If rs_BidTable!id = 21899 Then Stop

On Local Error GoTo SuitHCP_OK_Error

    If HCPs = "SS" And _
       bln_SolidSuit(s) _
    Then
        SuitHCP_Ok = True
        Exit Function
    End If
    
    If InStr(HCPs, "<A") > 0 Then
        If Left$(str_Suit(s), 1) = "A" Then
            Exit Function
        End If
        HCPs = Replace(HCPs, "<A", "")
    ElseIf InStr(HCPs, "<K") > 0 Then
        If Left$(str_Suit(s), 1) = "A" Or Left$(str_Suit(s), 1) = "K" Then
            Exit Function
        End If
        HCPs = Replace(HCPs, "<K", "")
    End If

    If LCase$(Left$(HCPs & "  ", 2)) = "or" Then
        ' Note: "or" is handled in the calling sub. '
        'Stop ' Single-step back to the calling sub to find it. '
        qtyOrHCP = True
        If Len(HCPs) < 4 Then
            MsgBox "HCP = 'or ' only."
            Exit Function
        End If
        HCPs = Mid$(HCPs, 3)
    End If

    'Get rid of blanks:
    HCPs = Replace(HCPs, " ", "", , 99)

    If HCPs = "" Then
        SuitHCP_Ok = True
        Exit Function
    End If

    ok = False
    originalHCPs = HCPs
    HCPs = LCase$(HCPs)

    ' Have Help Suit: '
    If HCPs = "hh" Then
        If Len(str_Suit(s)) < 3 Then
            SuitHCP_Ok = True
        ElseIf Left$(str_Suit(s), 1) = "A" Or _
               Left$(str_Suit(s), 1) = "K" _
        Then
            SuitHCP_Ok = True
        End If
        Exit Function
    End If

    ' Ask for Help Suit: '
    ' (Doesn't have A or K.) '
    If HCPs = "ha" Then
        If Len(str_Suit(s)) > 2 Then
            If Left$(str_Suit(s), 1) <> "A" Or _
               Left$(str_Suit(s), 1) <> "K" _
            Then
                SuitHCP_Ok = True
            End If
        End If
        Exit Function
    End If

    ' Intermediates required? '
    iPtr = InStr(HCPs, "i")
    If iPtr > 0 Then
        intermedReqd = Val(Mid$(HCPs, iPtr + 1))
        If iPtr = 1 Then ' e.g.: HCPs = "i2"
            HCPs = ""
            If int_HCP(s) > 4 Then
                SuitHCP_Ok = True
            Else
                If int_HCP(s) > 0 Then
                    intermedReqd = intermedReqd - 1
                End If
                SuitHCP_Ok = (int_Intermediates(s) >= _
                              intermedReqd)
            End If
            Exit Function
        Else
            HCPs = Left$(HCPs, iPtr - 1)
        End If
    End If

    ' If more than one spec and/or  '
    ' intermediates required, HCPs  '
    ' are stripped down to just the '
    ' first spec, such as 4+i2,x is '
    ' now just 4+. 'originalHCPs'   '
    ' keeps the original "4+i2,x".  '
    ' If only intermediates were    '
    ' specified, then HCPs = "".    '
    '                               '
    ' (Can't be as now written.)    '

If HCPs <> "" Then

    signR = Right$(HCPs, 1)
    signL = Left$(originalHCPs, 1)

    For i = 1 To 5
        leftSuit(i) = Left$(str_Suit(s) & " ", i)
    Next
    ' If I trimmed the blanks, then   '
    ' leftSuit(3) could be "AK", then '
    ' then instr("AKQ", "AK") would   '
    ' pass but not ("AKQ", "AK ").    '
    
    If originalHCPs = "P" Then
        ' Check if any honors are protected '
        ' by the bidding. E.g.: If K-J held,'
        ' they are not protected if LHO bid '
        ' them but they are if RHO bid them.'
        ' This *should* be restricted to the'
        ' 4th seat in the first round of    '
        ' bidding, such as after: 1H-P-1S   '
        ' where Kx of Hrts is not protected '
        ' because LHO bid them, but Kx of   '
        ' Spades is protected because RHO   '
        ' bid them. The applies more to the '
        ' Majors because 1C by RHO does not '
        ' mean as much as 1H or 1S.         '

        ' This routine is not as accurate   '
        ' as hand evaluation routines where '
        ' we actually retrieve the other    '
        ' players' bid entries to make sure '
        ' that a real suit was bid, rather  '
        ' than an artificial bid.           '

        'left1Suit = Trim$(Left$(str_Suit(s) & " ", 1))
        'left2Suit = Trim$(Left$(str_Suit(s) & "  ", 2))
        'left3Suit = Trim$(Left$(str_Suit(s) & "   ", 3))
        ok = True
        ' True except for the following: '
        If leftSuit(2) = "AQ" Or _
           leftSuit(1) = "K" Or _
           leftSuit(1) = "Q" _
        Then

            If leftSuit(3) <> "KQJ" And _
               leftSuit(3) <> "QJT" And _
               leftSuit(3) <> "AQJ" _
            Then
                Dim LHOs As String
                Dim RHOs As String

                If int_BidNum = 4 Then
                    LHOs = Right$(str_Call(int_BidNum - 3), 1)
                    RHOs = Right$(str_Call(int_BidNum - 1), 1)
                    If InStr("CDHS", LHOs) = s Then
                        ok = False
                    ElseIf InStr("CDHS", RHOs) <> s Then
                        ok = False
                    End If
                End If
            End If
        End If

    '  Check for stoppers specified: '

    ElseIf HCPs = "1stop" Then
        ok = (sng_NumStops(s) >= 1)
        extra = (sng_NumStops(s) > 1)

    ElseIf HCPs = "2stop" Then
        ok = (sng_NumStops(s) > 1)
        extra = (sng_NumStops(s) > 2)

    ElseIf HCPs = "hstop" Then
        ok = (sng_NumStops(s) > 0)
        extra = (sng_NumStops(s) > 0.5)

    ElseIf HCPs = "0stop" Then
        ok = (sng_NumStops(s) < 1)
        If InStr(HCPs, ",") = 0 Then
            SuitHCP_Ok = ok
            Exit Function
        End If
        
    '  Check for Ace or King specified: '

    ElseIf InStr(originalHCPs, "A") > 0 Then
        ' Since case matters with "A"/"a", '
        ' and HCPs has been lowercased, we '
        ' must go back to originalHCPs to  '
        ' check for "A", "A,K", "4+A, etc. '
        ok = bln_Suit(s, Ace)
        If ok And Left$(originalHCPs, 1) <> "A" Then
            i = InStr(originalHCPs, "A")
            X = Mid$(originalHCPs, i - 1, 1)
            If Not (X = " " Or X = ",") Then
                ' e.g.: not "A" or "5+,A" '
                ' Testing bA, 5+A, etc.   '
                ' to make sure the spec   '
                ' before the A is okay.   '
                X = Left$(originalHCPs, i - 1)
                ok2 = SuitHCP_Ok(s, X)
                ok = ok2
                ' e.g.: For "aA", call this routine for just "a". '
                ' ok was set to True above if suit has an Ace,    '
                ' and is reset to false if first part is not ok.  '
            End If
        End If

        If bln_Suit(s, King) And Right$(originalHCPs, 1) = "K" Then
            ok = True
        End If

    ElseIf originalHCPs = "<K" Then
        ok = (bln_Suit(s, King) = False And _
              bln_Suit(s, Ace) = False)

    ElseIf InStr(HCPs, "k") > 0 Then
        ' Case doesn't matter with "k".
        ok = bln_Suit(s, King)

    ElseIf InStr(HCPs, "q") > 0 Then
        ok = bln_Suit(s, Queen)

    ElseIf InStr(HCPs, "/") = 2 Then
        ' #/#. E.g.: 2/3 = 2 of top 3. '
        n = Val(Mid$(HCPs, 3, 1))
        If bln_Suit(s, Ace) Then j = 1 Else j = 0
        For i = 13 To 13 - n + 2 Step -1
            If bln_Suit(s, i) Then
                j = j + 1
            End If
        Next
        ok = j >= Val(Mid$(HCPs, 1, 1))
        
    '  Check for Controls specified: '

    ElseIf HCPs = "frc" Then
        'First Round Control (void or Ace)
        ok = int_SuitLen(s) = 0 Or _
              bln_Suit(s, Ace)

    ElseIf HCPs = "src" Then
        'Second Round Control (singleton or King)
        ok = int_SuitLen(s) = 1 Or _
              bln_Suit(s, King)
              
    End If

    If ok Then
        SuitHCP_Ok = True
        Exit Function
    End If


    '  Do the math:  '
    ok2 = True
    
    ' High card specs: (e.g.: "3H4") '
    i = InStr(originalHCPs, "H")
    If i > 0 Then
        h = Val(Mid$(originalHCPs & " ", i + 1, 1))
        t = 0
        If bln_Suit(s, Ace) Then t = t + 1
        If bln_Suit(s, King) Then t = t + 1
        If bln_Suit(s, Queen) Then t = t + 1
        If bln_Suit(s, Jack) Then t = t + 1
        If bln_Suit(s, 10) Then
            t = t + 1
            If bln_Suit(s, 9) Then
                t = t + 1
                If bln_Suit(s, 8) Then t = t + 1
            End If
        End If
        If t < h Then ok2 = False
    End If

    If signL = "<" Then ' e.g.: <5 '
        ok = int_HCP(s) < Val(Mid$(HCPs, 2))

    ElseIf (Val(HCPs) > 0 And _
            Mid$(HCPs, 2) <> "stop") Or _
           HCPs = "0" _
    Then
        ptsNeeded = Val(HCPs)

        If signR = "+" Then ' e.g.:  4+ '
            ok = int_HCP(s) >= ptsNeeded
            extra = int_HCP(s) > ptsNeeded
            ' I'm going to arbitrarily count '
            ' 2 extra HCPs as = to 1 intermed.
            intermedReqd = intermedReqd - _
                           (int_HCP(s) - ptsNeeded) \ 2
            If intermedReqd < 0 Then
                intermedReqd = 0
            End If

        ElseIf Right$(HCPs, 1) = "-" Then ' 4-
            ' Should be using "<5", but just in case: '
            ok = int_HCP(s) <= ptsNeeded

        ElseIf Mid$(HCPs, 2, 1) = "-" Then
            ' Range, e.g.: 4-7 '
            ok = (int_HCP(s) >= ptsNeeded And _
                  int_HCP(s) <= Val(Mid$(HCPs, 3)))

        ElseIf InStr(HCPs, "/") = 0 Then
            ' No sign or range, so must be '
            ' a specific number, like "4". '
            ' (Could be "1stop", but that  '
            ' was fixed above.)            '
            ok = (int_HCP(s) = ptsNeeded)

        End If

    ElseIf Len(HCPs) > 0 And InStr(HCPs, "stop") = 0 Then
        ' At this point, HCPs must be a '
        ' code like a, b, c, ab, etc.   '
        ' Convert into actual cards:    '

        left1 = Left$(originalHCPs, 1)
        left2 = Left$(originalHCPs & " ", 2)

        If left2 = "ab" Then
            honorsList = "AK,AQ,KQ"
            ok = InStr(honorsList, leftSuit(2))
            If Not ok And leftSuit(3) <> "" Then
                honorsList = "AJT,KJT,QJT"
                ok = InStr(honorsList, leftSuit(3))
            End If

        ElseIf left2 = "bc" Then
            honorsList = "AJT,KJT,QJT,KQJ,AQJ,AKQ,AKJ,AKQJ"
            ok = InStr(honorsList, leftSuit(3))

        ElseIf left1 = "d" Then
            honorsList = "AKQJ,KQJT,AKQT,AKQJT"
            ok = InStr(honorsList, leftSuit(4))

        ElseIf left1 = "c" Then
            honorsList = "AKQ,AKJ,AQJ,KQJ"
            ok = InStr(honorsList, leftSuit(3))

        ElseIf left1 = "a" Or _
               left2 = "ac" _
        Then ' "c" is redundant with "a" '
             ' If any "a" spec is good,  '
             ' then some "c" spec must be.
            honorsList = "AK,AQ,KQ"
            ok = InStr(honorsList, leftSuit(2))

        ElseIf left1 = "b" Then
            honorsList = "AJT,KJT,QJT"
            ok = InStr(honorsList, leftSuit(3))

        ElseIf left1 = "H" Then
            h = Val(Mid$(HCPs, 2, 1))
            i = Val(Mid$(HCPs & " ", 3, 1)) ' Ex: 4 of top 5 : h=4 and i=5 '
            If i = 0 Then i = h
            If h = 2 Then ' 2 of top 3/4/5 '
                If i = 3 Then
                    honorsList = "AK,AQ,KQ"
                ElseIf i = 4 Then ' 2 out of AKQJ '
                    honorsList = "AK,AQ,AJ,KQ,KJ,QJ"
                ElseIf i = 5 Then ' 2 out of AKQJT '
                    honorsList = "AK,AQ,AJ,AT,KQ,KJ,KT,QJ,QT"
                End If
                ok = InStr(honorsList, leftSuit(2))
            
            ElseIf h = 3 Then ' 3 of the top 6 (AKQJT9) '
                ' By using ">", i gets all the combos '
                ' from the i's before.
                ' the
                If i > 2 Then honorsList = "AKQ"
                If i > 3 Then honorsList = honorsList & ",AKJ,AQJ,KQJ"
                If i > 4 Then honorsList = honorsList & ",AKT,AQT,AJT,KQT,QJT"
                If i > 5 Then honorsList = honorsList & ",AK9,AQ9,AJ9,KQ9,KJ9,KT9,QJ9"
                ok = InStr(honorsList, leftSuit(3))
            
            ElseIf h = 4 Then ' 4 of the top 6 (AKQJT9) '
                If i > 3 Then honorsList = "AKQJ"
                If i > 4 Then honorsList = honorsList & ",AKQT,AKJT,AQJT,KQJT"
                If i > 5 Then honorsList = honorsList & ",AKQ9,AQT9,AJT9,KQT9,QJT9"
                If i > 5 Then honorsList = honorsList & ",AKJ9,AQJ9,AJT9,KQT9,KJT9,KT9,QJ9"
                ok = InStr(honorsList, leftSuit(4))
            
            ElseIf h = 5 Then ' 5 of the top 6 (AKQJT9) '
                honorsList = "AKQJT,AKQJ9,AKQT9,AKJT9,AQJT9,KQJT9"
                ok = InStr(honorsList, leftSuit(5))
            End If
        End If

        If i > 0 Then
            n = Val(Mid$(leftSuit(4), i, 1))
            If n = 9 Or n = 0 Then
                ok = InStr(honorsList, Left$(leftSuit(4), i))
            End If
        End If

        i = InStr("shdc", Right$(HCPs, 1))

        If i > 0 And s <> i Then
          If 1 = 2 Then ' See note below '
            ok = int_HCP(s) - int_HCP(i)
            ' Can't figure this out.  '
            ' A4-K65-QT9872-42        '
            ' Looking for a suit with '
            ' HCP = c (KQJ, AQJ, AKQ) '
            ' but int_HCP(s)-int_HCP(i)
            ' where s=2 and i=4, but  '
            ' that is -2 so ok=True.  '
            ' What I want to know is  '
            ' if 1st 3 cards in Dia   '
            ' are KQJ, AQJ, or AKQ.   '
          End If

        End If

    End If

    ok = ok And ok2 ' ok2 is for #Honors specified. '

End If

    ' Check intermediates:              '
    ' (intermedReqd is computed above.) '
    If (ok Or HCPs = "") And _
       intermedReqd > 0 _
    Then

        If extra Then
            ' In analysis above, the hand '
            ' had more HCPs than required '
            ' KQ2 should not be rejected  '
            ' for not having an intermed  '
            ' if KT2 is okay.             '
            ' An alternative is to add .5 '
            ' for 10 and .25 for 9 and    '
            ' for 8 with a 9 or 10, then  '
            ' instead of 3+i say 3.5+.    '
            ' & "4" is ok with "KT98".    '
            ' This could also be used for '
            ' total hand HCP.             '
            intermedReqd = intermedReqd - 1
        End If

        ok = (int_Intermediates(s) >= intermedReqd)
    End If


    ' Check RKCB 1430 responses:  '
    ' 5C = 1 or 4  DB code: kc1,4 '
    ' 5D = 0 or 3           kc0,3 '
    ' 5H = 2 or 5 w/o trump Q : kc2,5-Q '
    ' 5S = 2 or 5 with trump Q: kc2,5+Q '
    If s = 5 Then ' combined suits '
        MsgBox "Stop SuitHCP_Ok 1"
    End If


    SuitHCP_Ok = ok
    
Exit Function

SuitHCP_OK_Error:
    MsgBox Err.Description & vbCr & _
           "Error in Function SuitHCP_OK"
    Stop
    ' Use Resume to see error. '
    Resume
    On Local Error GoTo 0
End Function


Private Sub ShowAnalysis()
    Dim num As Long
    Dim sel As String
    Dim sel2 As String
    Dim nam As String

    bln_AnalyzeBid = True
    With frm.lstAnalysis
        Load frm ' frmBidBaseAnalysis
        'frm.Show  ' Don't show until told to. '
        .Clear
        .AddItem "Bid: " & rs_BidTable!Bid & _
                  "   Bid Order #: " & rs_BidTable!Bid_Order_Number & _
                  "   ID#: " & rs_BidTable!id
        If rs_BidTable!Test_Hand <> "" Then
            .List(0) = .List(0) & "   Sample: " & rs_BidTable!Test_Hand
        End If

        If Left$(rs_BidTable!Explanation, 2) = "S:" Then
            ' Some entries have something like "S:4, H:4, D:4, C:1, HCP:8+" '
            ' in the Explanation field. Don't know why. Don't want to see it.
            .AddItem "Descrip.:   (none)"

        ElseIf Len(rs_BidTable!Explanation) < 50 Then
            .AddItem "Descrip.:  " & rs_BidTable!Explanation
        Else
            .AddItem "     " & rs_BidTable!Explanation
        End If

        num = Val("" & rs_BidTable!ConvID)
        If num > 0 Then
            LookUpBidName num, sel, sel2, nam
        End If
        .AddItem "Bid Name:  " & nam & "   Deactivation: " & _
                 sel & sel2 & rs_BidTable!pct_Used
        .AddItem "Requirements for call--                 " & _
                 "(Hand being bid: " & str_PlayersHand & ")"
        .AddItem "HCPs:       " & rs_BidTable!HCP
        .AddItem "Distr.pts:  " & rs_BidTable!Dist_Pts
        .AddItem ""
        .AddItem "Custom pts: " & rs_BidTable!Cust_Pts
        .AddItem "Winners:    " & rs_BidTable!Winrs
        .AddItem "Losers:     " & rs_BidTable!Losers
        .AddItem "Quik Triks: " & rs_BidTable!QTs
        .AddItem "Intermed:    " & rs_BidTable!Intermed
        .AddItem "Stoppers:   " & rs_BidTable!Stopped_Suits
        .AddItem "Vulner.:    " & rs_BidTable!Vul
        .AddItem "Position:     " & rs_BidTable!pos
        .AddItem "Shape:        " & rs_BidTable!Shape
        .AddItem ""
        .AddItem "             Num    HCP"
        .AddItem "'x' :     " & _
                  Space$(7 - Len(rs_BidTable!xNum)) & rs_BidTable!xNum & _
                  Space$(8 - Len(rs_BidTable!xHCP)) & rs_BidTable!xHCP
        .AddItem "Spds:" & _
                  Space$(7 - Len(rs_BidTable!SNum)) & rs_BidTable!SNum & _
                  Space$(10 - Len(rs_BidTable!SHCP)) & rs_BidTable!SHCP
        .AddItem "Hrts:  " & _
                  Space$(7 - Len(rs_BidTable!HNum)) & rs_BidTable!HNum & _
                  Space$(10 - Len(rs_BidTable!HHCP)) & rs_BidTable!HHCP
        .AddItem "Dias: " & _
                  Space$(7 - Len(rs_BidTable!DNum)) & rs_BidTable!DNum & _
                  Space$(10 - Len(rs_BidTable!DHCP)) & rs_BidTable!DHCP
        .AddItem "Clbs: " & _
                  Space$(7 - Len(rs_BidTable!CNum)) & rs_BidTable!CNum & _
                  Space$(10 - Len(rs_BidTable!CHCP)) & rs_BidTable!CHCP
        .AddItem "Cmbo:    " & _
                  Left$(rs_BidTable!cmNum & Space$(9), 9) & _
                  " " & rs_BidTable!cmHCP
        .Refresh
    End With
    With frm
        .cmd_Cancel.Enabled = False
        .cmd_Continue.Enabled = False
        .MousePointer = 11
    End With
End Sub


Private Function GetDataFromMatchingEntry() As String
    '*********************************************
    ' Purpose: The bid from an entry with specs  *
    ' matching the specified hand can usually be *
    ' used as-is, but sometimes there are other  *
    ' considerations:                            *
    '   1. If the bid has "x" in it, such as 2x, *
    '        it must be converted to the suit.   *
    '   2. If the user is testing an entry he    *
    '        just created or changed, then the   *
    '        entry's ID will have been specified *
    '        and if the matching entry found by  *
    '        this routine does not match that ID *
    '        the user must be notified and given *
    '        the choice of continuing testing or *
    '        cancelling the search.              *
    '   3. Get the info to disclose.             *
    ' Other data generated by this routine       *
    ' include disclosures.                       *
    '*********************************************

    Dim m As String
    Dim ans As Integer
    Dim num As Long
    Dim sel As String
    Dim sel2 As String
    Dim nam As String
    Dim i As Long
    Dim q As Long
    Dim s As Long

On Local Error GoTo GetDataFromMatchingEntry_Error

    With rs_BidTable
        str_BON = !Bid_Order_Number
        If lng_TestID > 0 And _
           !id <> lng_TestID And _
           (InStr(str_Position, !pos) Or _
           !pos = "") _
        Then
            int_ErrCode = 88
            str_ErrMsg = "Test Hand for entry" & vbCr & _
                         "ID#" & lng_TestID & _
                         "  BON#" & !Bid_Order_Number & vbCr & _
                         "also matches specs" & vbCr & _
                         "for entry ID#" & !id & _
                         "  BON#" & str_BON
            lng_FoundID = !id
            Exit Function
            
          If 1 = 2 Then
            If Trim(!Prior_Bids) <> "" Then
                m = m & "Prior bids: " & !Prior_Bids & vbCr
            End If
            m = m & "Entry #" & str_BON
            m = m & "      Pct: " & !pct_Used
            m = m & "      Bid: " & !Bid
            num = Val("" & !ConvID)
            LookUpBidName num, sel, sel2, nam
            If Trim(nam) <> "" Then
                m = m & vbCr & _
                    "Bid Name:" & vbCr & _
                    "   " & nam
                If sel <> "" Then
                    m = m & vbCr & _
                        "Conv.Code:    " & sel & vbCr & _
                        "Subconv.Code: " & sel2
                End If
            End If
            m = m & vbCr & vbCr + "Keep testing?" & vbCr & vbCr & _
                    "(Click No to compare these two entries.)"
            ans = MsgBox(m, vbYesNo)
            If ans = vbNo Then
                int_ErrCode = 88
                str_ErrMsg = "Entries #" & lng_TestID & _
                             " and #" & vbCr & _
                             !Bid_Order_Number & ":" & _
                             !id & " conflict."
            ElseIf ans = vbYes Then
                str_ErrMsg = "Keep testing"
            End If
          End If
        Else
            'If !id = 14697 Then Stop
            If InStr(LCase$(!Bid), "x") > 0 Then
                If str_xSuit = "" And bln_WinnersOK Then
                    For i = 1 To 4
                        If int_SuitLen(i) > q Then
                            q = int_SuitLen(i)
                            s = i
                        End If
                    Next
                    str_xSuit = Mid$("CDHS", s, 1)
                End If

                str_Call(int_CallNum) = _
                    Left$(!Bid, 1) & _
                    str_xSuit
            Else
                str_Call(int_CallNum) = _
                    !Bid

            End If

            lng_ID = !id

            If bln_GetDisclosure Then
                GetDiscl
            End If

        End If
    End With
Exit Function

GetDataFromMatchingEntry_Error:
    MsgBox Err.Description & vbCr & _
           "in Sub GetDataFromMatchingEntry for" & vbCr & _
           "entry ID#" & rs_BidTable!id
    Stop
    Resume ' To go back to code causing error. '
    Resume Next
    On Local Error GoTo 0
End Function


Private Sub GetDiscl()
    ' Loop thru all entries matching Prior Calls '
    ' and selected bid to get their disclosures  '
    ' to give to client. As a reminder: you can- '
    ' not just give the user the disclosure for  '
    ' the selected entry because that gives him  '
    ' more specific information than to which he '
    ' is entitled. For example, if there are two '
    ' entries for "1H-P-4H" where one is weak and'
    ' preemptive and the other is expecting to   '
    ' make, the user is not entitled to know     '
    ' which bid was chosen and why, just that it '
    ' could be either-or. However, if the chosen '
    ' bid is conventional, only that bid's specs '
    ' need to be shown. If it is not convention- '
    ' al (BON #900000+) a matching bid does not  '
    ' need to be shown if its entry is deactivated.

    Dim s As String
    Dim i As Long

    If str_Call(int_CallNum) = "Pass" Then
        Exit Sub
    End If

    ' Add the disclosure for the entry '
    ' being used for the bid. '
    str_Disclosure = ""
    AddToDisclosure

    If rs_BidTable!Bid_Order_Number < "900000" Then
        ' If BON is < 900000 then it is conventonal  '
        ' and disclosures for entries with the same  '
        ' bid which are not part of the same conven- '
        ' tional do not have to be shown. '
        Exit Sub
    End If

    ' Get the first entry matching the Prior Bids: '
    s = "SELECT * FROM Bids"
    s = s + " WHERE Prior_Bids='" & str_PriorBids & "'"
    s = s + " ORDER BY Bid_Order_Number;"
    Set rs_BidTable = db_BidBase.OpenRecordset(s, dbOpenSnapshot)

    ' Look for entries matching the bid made: '

    Do While rs_BidTable!Bid_Order_Number <> "ZZZZZZ" And _
             rs_BidTable!Prior_Bids = str_PriorBids
        DoEvents
        If rs_BidTable!Bid = str_Call(int_CallNum) And _
           (rs_BidTable!pct_Used = "" Or _
            Val(rs_BidTable!pct_Used) <> 0) _
        Then
            'Debug.Print rs_BidTable!id; "  "; rs_BidTable!Bid_Order_number
            AddToDisclosure
        End If

        ' Get next record, if any: '
        If Not rs_BidTable.EOF Then
            rs_BidTable.MoveNext
            If rs_BidTable.EOF Then
                rs_BidTable.MovePrevious
                If i = rs_BidTable!id Then
                    ' Just repeating. '
                    Exit Do
                ElseIf rs_BidTable!id = lng_ID Then
                    Exit Do
                End If

            End If
        Else
            Exit Do
        End If
        DoEvents
        i = rs_BidTable!id
    Loop

End Sub


Private Sub AddToDisclosure()

    Dim s As String
    Dim X As String
    Dim sb As String
    Dim sh As String
    Dim a As String
    Dim i As Integer
    Dim nb As String
    Dim nn As Integer
    Dim num As Long
    Dim sel As String
    Dim sel2 As String
    Dim nam As String
    Dim p As String
    Dim v As String

With rs_BidTable
    num = Val("" & !ConvID)
    LookUpBidName num, sel, sel2, nam
    'If Trim(nam) <> "" Then
    '    If sel = "" Or sel <> "0" Then
    '        nam = nam & "  Cvn: " & sel & "  Sub: " & sel2
    '    End If
    'End If

    If sel = "0" Or sel2 = "0" Or _
       (InStr(str_Position, !pos) = 0 And _
        !pos <> "") Or _
       (InStr(str_Vul, !Vul) = 0 And _
        !Vul <> "") _
    Then
        ' Entry is deactivated or position '
        ' or vul don't match, so it is not '
        ' an alternative to the bid made,  '
        ' so don't show its specs. '
        Exit Sub
    End If

    s = !Bid & _
        ".  BON:" & !Bid_Order_Number & _
        ".  ID#:" & !id & ". "
    If nam <> "" Then
      s = s & vbCrLf & Chr$(34) & nam & Chr$(34)
    End If
    s = s & vbCrLf

    If !pos <> "" Then
        p = !pos
        s = s & "Position:"
        If Len(p) > 1 Then s = s & vbCrLf
        sb = ""
        If InStr(p, "Y") Then
            sb = "  Bidder is a passed hand." & vbCrLf
        End If
        If InStr(p, "N") Then
            If sb <> "" Then sb = sb & "  or b" Else sb = "  B"
            sb = sb & "idder is an unpassed hand." & vbCrLf
        End If
        If InStr(p, "1") Then
            If sb <> "" Then sb = sb & "  or b" Else sb = "  B"
            sb = sb & "idder is in 1st seat." & vbCrLf
        End If
        If InStr(p, "2") Then
            If sb <> "" Then sb = sb & "  or b" Else sb = "  B"
            sb = sb & "idder is in 2nd seat." & vbCrLf
        End If
        If InStr(p, "3") Then
            If sb <> "" Then sb = sb & "  or b" Else sb = "  B"
            sb = sb & "idder is in 3rd seat." & vbCrLf
        End If
        If InStr(p, "4") Then
            If sb <> "" Then sb = sb & "  or b" Else sb = "  B"
            sb = sb & "idder is in 4th seat." & vbCrLf
        End If
        If InStr(p, "F") Then
            If sb <> "" Then sb = sb & "  or o" Else sb = "  O"
            sb = sb & "pener is 1st or 2nd seat." & vbCrLf
        End If
        If InStr(p, "T") Then
            If sb <> "" Then sb = sb & "  or o" Else sb = "  O"
            sb = sb & "pener is 3rd or 4th seat." & vbCrLf
        End If
        If InStr(p, "P") Then
            If sb <> "" Then sb = sb & "  or p" Else sb = "  P"
            sb = sb & "artner is a passed hand." & vbCrLf
        End If
        If InStr(p, "U") Then
            If sb <> "" Then sb = sb & "  or p" Else sb = "  P"
            sb = sb & "artner is an unpassed hand." & vbCrLf
        End If
        s = s & sb
    Else
        s = s & "Position:  any." & vbCrLf
    End If

    If !Vul <> "" Then v = !Vul Else v = "any"
        If v = "F" Then
            v = "favorable"
        ElseIf v = "U" Then
            v = "unfavorable"
        ElseIf v = "B" Then
            v = "both"
        ElseIf v = "N" Then
            v = "neither"
        ElseIf v = "V" Then
            v = "vul. bidder"
        ElseIf v = "X" Then
            v = "non-vul. bidder"
        Else
            v = "any"
        End If
    s = s & "Vulnerability:  " & v & ".   "
    If !Shape <> "" Then
        s = s & "Shape:  " & !Shape & "."
    End If
    s = s & vbCrLf
    X = s
    ' Strength:                       '
    If !HCP <> "" Then
        s = s & "Hand HCP: " & !HCP & ".  "
    End If
    
    If !Dist_Pts <> "" Then
        s = s & "Distrib: " & !Dist_Pts & "."
    End If
    
    If X <> s And s <> "" Then
        s = s & vbCrLf
        X = s
    End If
    
    If !Winrs <> "" Then
        s = s & "Winners: " & !Winrs & ".  "
    End If
    
    If !Losers <> "" Then
        s = s & "Losers: " & !Losers & ".  "
    End If
    
    If !Stopped_Suits <> "" Then
        s = s & "Stops: " & !Stopped_Suits & "."
    End If
    
    If X <> s And s <> "" Then
        s = s & vbCrLf
        X = s
    End If
    
    ' Suit lengths & HCPs:               '
    sb = ""
    If !xNum <> "" Then sb = sb & "x:" & !xNum & ".  "
    If !SNum <> "" Then sb = sb & "S:" & !SNum & ".  "
    If !HNum <> "" Then sb = sb & "H:" & !HNum & ".  "
    If !DNum <> "" Then sb = sb & "D:" & !DNum & ".  "
    If !CNum <> "" Then sb = sb & "C:" & !CNum & ".  "
    If !cmNum <> "" Then sb = sb & "cmb:" & !cmNum & "."
    If sb <> "" Then s = s & "Qty: " & sb
    If sb <> "" Then
        sb = sb & vbCrLf
    End If

    sb = ""
    If !xHCP <> "" Then sb = sb & "x:" & !xHCP & ".  "
    If !SHCP <> "" Then sb = sb & "S:" & !SHCP & ".  "
    If !HHCP <> "" Then sb = sb & "H:" & !HHCP & ".  "
    If !DHCP <> "" Then sb = sb & "D:" & !DHCP & ".  "
    If !CHCP <> "" Then sb = sb & "C:" & !CHCP & ".  "
    If sb <> "" Then
        sb = sb & vbCrLf
    End If
    If !cmHCP <> "" Then
        sb = sb & "Suit combo:  " & !cmHCP & ".  "
    End If
    
    If sb <> "" Then
        s = s & vbCrLf & "HCP: " & sb
    End If
    
    If X <> s And s <> "" Then
        s = s & vbCrLf
        X = s
    End If

    ' Asks pard for/to bid:             '
    sb = ""
    If !Artificial <> "" Then
        sb = sb & "Artificial.   "
    End If
    
    If !Alert = "A" Then
        sb = sb & "Alert  "
    End If
    
    If Left$(!Explanation, 1) <> "[" And _
       !Explanation <> "" _
    Then
        'sb = sb & "Explanation:" & vbCrLf ' & "   "
        sb = sb & Chr$(34) & !Explanation & Chr$(34)
    End If
    
    If sb <> "" Then
        s = s & sb
    End If
    
    If X <> s And s <> "" Then
        s = s & vbCrLf
        X = s
    End If

    sb = ""
    a = "" & LCase$(!Asks)
    If a = "" Then
        'skip it'
    ElseIf a = "g" Then
        sb = sb & "game force"
    ElseIf a = "1" Then
        sb = sb & "bid again"
    ElseIf a = "poc" Then
        sb = sb & "pass or correct"
    ElseIf a = "p" Then
        sb = sb & "pick a suit/slam"
    ElseIf a = "a" Then
        sb = sb & "show Aces"
    ElseIf a = "k" Then
        sb = sb & "show Kings"
    ElseIf a = "kc" Then
        sb = sb & "show # of Keys"
    ElseIf Left$(a, 1) = "h" Then
        sh = Mid$(a, 2, 1)
        i = InStr("cdhs", sh)
        sh = str_SuitNames(i)
        sb = sb & "for help in " & sh
    ElseIf Left$(a, 1) = "c" Then
        sh = Mid$(a, 2, 1)
        i = InStr("cdhs", sh)
        sh = str_SuitNames(i)
        sb = sb & "for controls in " & sh
    ElseIf Left$(LCase$(a), 1) = "b" Then
        sb = LCase$(Mid$(a + " ", 2, 1))
        Select Case Mid$(a + "  ", 3, 1)
            Case "m"
                sh = sh & "a minor"
            Case "M"
                sh = sh & "a major"
            Case Else
                i = InStr("cdhs", sh)
                sh = str_SuitNames(i)
        End Select
        sb = sb & "to bid " & sh
    ElseIf Trim$(!Asks) <> "" Then
        sb = sb & "to " & !Asks
    End If
    
    If sb <> "" Then
        s = s & sb
    End If
    
    If X <> s And s <> "" Then
        s = s & vbCrLf
    End If
    
    If s <> "" And InStr(str_Disclosure, s) = 0 Then
        If Right$(s, 4) <> vbCrLf & vbCrLf Then
            s = s & vbCrLf
        End If
        str_Disclosure = str_Disclosure & s
    End If
End With

End Sub


Private Sub LookUpBidName(num As Long, _
                          sel As String, _
                          sel2 As String, _
                          nam As String)

    With rs_BidNames
        If nam = "" Then
            .Index = "Number"
            .Seek "=", num
        Else
            .Index = "Bid_Name"
            .Seek "=", nam
        End If

        If .NoMatch = True Then
            Exit Sub
        End If

        num = !IDNum
        sel = Trim$("" & !Selected)
        sel2 = Trim$("" & !SubCat_Selected)
        nam = !Bid_Name
    End With
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

With rs_BidsQuery

    ' Get LHO's suits: '
    j = 1
    n = int_LHOs_Pos
    Do While j < 3 And _
             str_Bid(n) <> "" And _
             str_Bid(n) <> "P"

        DoEvents
        pb = MakePB(n)
        b = str_Bid(n)
        If Val(b) > 0 Then ' Some # of some suit bid.'
            GetOPsBid pb, b
            If .NoMatch = False And _
               .EOF = False _
            Then
                If !Artificial = "" Or _
                   IsNull(!Artificial) _
                Then
                    If xSuitEqualSuitBid(Right$(b, 1)) Then
                        s = Right$(b, 1)
                        int_LHOs_Suit(j) = InStr("CDHS", s)
                        j = j + 1
                    End If
                Else
                    MsgBox "Stop GetOtherPlayersData 1"
                    ' Get real suits.
                End If
            End If
        End If
        n = n + 4
        DoEvents
    Loop

    ' Get RHO's suits: '
    j = 1
    n = int_RHOs_Pos
    Do While j < 3 And _
             str_Bid(n) <> "" And _
             str_Bid(n) <> "P"

        DoEvents
        pb = MakePB(n)
        b = str_Bid(n)
        If Val(b) > 0 Then ' Some # of some suit bid.'
            GetOPsBid pb, b
            If .NoMatch = False And _
               .EOF = False _
            Then
                If !Artificial = "" Or _
                   IsNull(!Artificial) _
                Then
                    If xSuitEqualSuitBid(Right$(b, 1)) Then
                        s = Right$(b, 1)
                        int_RHOs_Suit(j) = InStr("CDHS", s)
                        j = j + 1
                    End If
                Else
                    MsgBox "Stop GetOtherPlayersData 1"
                    ' Get real suits.
                End If
            End If
        End If
        n = n + 4
        DoEvents
    Loop

    ' Get Pard's suits: '
    j = 1
    k = 1
    n = int_Pards_Pos
    Do While j < 3 And _
             str_Bid(n) <> "" And _
             str_Bid(n) <> "P"

        DoEvents
        pb = MakePB(n) ' Prior Bids for pard's bid.'
        b = str_Bid(n)   ' Pard's bid.'
        If Val(b) > 0 Then ' Some # of some suit bid.'
            GetOPsBid pb, b
            With rs_BidsQuery
            If .NoMatch = False And _
               .EOF = False _
            Then
                If !Artificial = "" Or _
                   IsNull(!Artificial) _
                Then
                    If xSuitEqualSuitBid(Right$(b, 1)) Then
                        s = Right$(b, 1)
                        int_Pards_Suit(j) = _
                               InStr("CDHS", s)

                        If int_Pards_Suit(j) = Clbs Then
                            int_Pards_SuitLen(j) = Val("" & !CNum)

                        ElseIf int_Pards_Suit(j) = Dias Then
                            int_Pards_SuitLen(j) = Val("" & !DNum)

                        ElseIf int_Pards_Suit(j) = Hrts Then
                            int_Pards_SuitLen(j) = Val("" & !HNum)

                        ElseIf int_Pards_Suit(j) = Spds Then
                            int_Pards_SuitLen(j) = Val("" & !SNum)
                        End If

                        If int_Pards_Suit(j) = int_Pards_Suit(j - 1) Then
                            int_Pards_RebidSuit(k) = _
                                  int_Pards_Suit(j)
                            int_Pards_Suit(j) = 0
                            int_Pards_SuitLen(j) = 0
                            k = k + 1
                        Else
                            j = j + 1
                        End If
                    End If
                ElseIf j = 1 Then
                    ' Get real suits.
                    q(1) = Val("" & !CNum)
                    q(2) = Val("" & !DNum)
                    q(3) = Val("" & !HNum)
                    q(4) = Val("" & !SNum)
                    For i = 1 To 4
                        z(i) = i
                    Next
                    Do
                        DoEvents
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
                        DoEvents
                    Loop
                    For i = 1 To 4
                        If q(z(i)) > 3 Then

                            MsgBox "Stop GetOtherPlayersData 3"
                            ' check the following:
                            int_Pards_SuitLen(i) = q(z(i))
                            int_Pards_Suit(i) = z(i)
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If
            End With
        End If
        n = n + 4
        DoEvents
    Loop

End With

    ' Get current player's suits: '
    m = int_BidNum Mod 4
    If m = 0 Then m = 4
    ' First bid: '
    k = 1
    c = " CDHS"
    Do While k < 4 And m < int_BidNum
        DoEvents
        s = Right$(" " & str_Bid(m), 1)
        n = InStr(c, s) - 1
        If n > 0 Then
            Mid$(c, n + 1) = " " ' Don't use same suit twice. '
            j = 1
            Do While int_Pards_Suit(j) > 0
                DoEvents
                If int_Pards_Suit(j) = n And _
                   int_SuitLen(n) + int_Pards_SuitLen(j) > 7 _
                Then
                    int_Fit_Suit(k) = n
                    int_Fit_Len(k) = _
                          int_SuitLen(n) + _
                          int_Pards_SuitLen(j)
                    bln_Fit_Confirmed = True
                    bln_Fit_Established = True
                    Exit Do
                End If
                j = j + 1
                If j > 3 Then Exit Do
                DoEvents
            Loop
            If int_SuitLen(n) > 3 Then
                int_My_Suit(k) = n
            End If
        End If
        k = k + 1
        m = m + 4
        DoEvents
    Loop

    bln_AlreadyHaveOPsData = True
End Sub


Private Function MakePB(j As Integer)
    Dim i As Integer
    Dim pb As String

    If j = 1 Then Exit Function
    pb = str_Bid(1)
    For i = 2 To j - 1
        pb = pb & "-" & str_Bid(i)
    Next
    MakePB = pb
End Function


Private Function xSuitEqualSuitBid(b As String) As Boolean

    With rs_BidsQuery
    '               '
    If .NoMatch Or .EOF Then
        Exit Function
    End If
    If LCase$(Right$(!Bid, 1)) <> "x" Then
        xSuitEqualSuitBid = True
        Exit Function
    End If

    Do
        DoEvents
        Select Case b
        Case Is = "C"
            xSuitEqualSuitBid = (LCase$(Left$(!CNum & " ", 1)) = "x")
            Exit Function
        Case Is = "D"
            xSuitEqualSuitBid = (LCase$(Left$(!DNum & " ", 1)) = "x")
            Exit Function
        Case Is = "H"
            xSuitEqualSuitBid = (LCase$(Left$(!HNum & " ", 1)) = "x")
            Exit Function
        Case Is = "S"
            xSuitEqualSuitBid = (LCase$(Left$(!SNum & " ", 1)) = "x")
            Exit Function
        End Select
        If .EOF Then Exit Function
        .MoveNext
        DoEvents
    Loop While LCase$(Right$(!Bid, 1)) = "x"
    xSuitEqualSuitBid = True
    '      '
    End With
End Function


'**************************************************************************
' Procedure : GetOPsBid                                                   *
' Called By : GetOtherPlayersData                                         *
' DateTime  : 12/29/2005 07:21                                            *
' Purpose   : Load other player's bids into rs_BidsQuery recordset for    *
'             use by GetOtherPlayersData sub.                             *
'**************************************************************************

Private Sub GetOPsBid(pb As String, _
                      b As String)
    Dim s As String
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    Dim qSel As String
    Dim bx As String

    On Error GoTo GetOPsBid_Error

    If Not bln_AlreadyInitialized Then
        InitializeBidBaseDLL
    End If

    If InStr(" CDHS", Right$(" " & b, 1)) > 1 Then
        bx = LCase$(Left$(b, 1)) & "x"
        ' e.g.: If b=1S, bx=1x     '
        ' Have to search for both. '
    End If

    s = "BidNames.Selected, " & _
        "BidNames.SubCat_Selected, " & _
        "Bids.Pct_Used, Bids.RoundOfBidding, " & _
        "Bids.Player, Bids.PB_Sort_Order, " & _
        "Bids.Prior_Bids, " & _
        "Bids.Bid_Order_Number, " & _
        "Bids.Bid, " & _
        "BidNames.Bid_Name, " & _
        "Bids.HCP, " & _
        "Bids.DistPts, " & _
        "Bids.Cust_Pts, "
    s = s & "Bids.A_Cust, " & _
        "Bids.K_Cust, " & _
        "Bids.Q_Cust, " & _
        "Bids.J_Cust, " & _
        "Bids.T_Cust, " & _
        "Bids.Winrs, " & _
        "Bids.Losers, " & _
        "Bids.QTs, " & _
        "Bids.Stopped_Suits, "
    s = s & "Bids.Shape, " & _
        "Bids.Intermed, " & _
        "Bids.xNum, " & _
        "Bids.xHCP, " & _
        "Bids.sNum, " & _
        "Bids.sHCP, " & _
        "Bids.hNum, " & _
        "Bids.hHCP, " & _
        "Bids.dNum, " & _
        "Bids.dHCP, " & _
        "Bids.cNum, " & _
        "Bids.cHCP, " & _
        "Bids.cmNum, " & _
        "Bids.cmHCP, "
    s = s & "Bids.Vul, " & _
        "Bids.Pos, " & _
        "Bids.Artificial, " & _
        "Bids.Alert, " & _
        "Bids.Explanation, " & _
        "Bids.Test_Hand, " & _
        "Bids.Asks, " & _
        "Bids.DateOfEntry, " & _
        "Bids.Best_Contract, " & _
        "Bids.Num_Used, " & _
        "Bids.Source, " & _
        "Bids.LHO, " & _
        "Bids.Pard, " & _
        "Bids.RHO, " & _
        "Bids.ID "
    s = s & "FROM BidNames RIGHT JOIN Bids " & _
            "ON BidNames.IDnum = Bids.ConvID "
    qSel = "SELECT "
    ' If Selected = "0" or "T" then "We" are not using it.'
    ' If _1.Selected (which is "Use Against") = "0" or "W",
    '   then "They" are not using it.                     '
    s = " WHERE (Bids.Prior_Bids='" & pb & "' OR " & _
                "Bids.Prior_Bids Is Null)" & _
          " AND (Bids.Bid='" & b & "' OR " & _
                "Bids.Bid='" & bx & "')" & _
          " AND (BidNames.Selected<>'0' OR " & _
                "BidNames.Selected Is Null)" & _
          " AND (BidNames.SubCat_Selected<>'0' OR " & _
                "BidNames.SubCat_Selected Is Null)" & _
          " AND (Pct_Used<>'0' OR Pct_Used Is Null);"

    Set rs_BidsQuery = db_BidBase.OpenRecordset(qSel & s, dbOpenDynaset)

Exit_GetOPsBid:
   On Error GoTo 0
   Exit Sub

GetOPsBid_Error:
'''''''''''''''
    MsgBox Err.Description & " getting entries for " & vbCr & _
           "Prior Bids of " & pb & vbCr & _
           "and Current Bid of " & b
    Stop
    Resume 'Next
    On Error GoTo 0
End Sub


Sub InitializeHandVal()
    ChDir App.Path
    Set ws_HandVal = DBEngine.Workspaces(0)
    Set db_HandVal = ws_HandVal.OpenDatabase("BB-HEP.accdb")
    Set rs_HandVal = db_HandVal.OpenRecordset("HandVal")
    Set rs_HandValBidReq = db_HandVal.OpenRecordset("HandVal_BidReq")
    bln_HandValInitialized = True
End Sub


'                       Hand Evaluation Routines                     '


'******************************************************************
' Name:  Eval_A_Hand                                              '
' Purpose: Calculate custom hand evaluation points based on specs '
'          in the HandVal table in BidBase.                       '
' See the programmer docs for details and explanations.           '
' All Subs used by this Sub start with "Eval_" to distinguish     '
' them from Subs used by the LookUpBid sub.                       '
'******************************************************************
'
Private Sub Eval_A_Hand(sysName As String)
    Dim n As Single
    Dim m As String
    Dim i As Integer
    Dim j As Integer

    Eval_Initialize sysName

    With rs_HandVal
    '              '
    Do While Not .EOF And Not .NoMatch
        Eval_SetUpFieldVariables

        DoEvents
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
        .MoveNext
        DoEvents
    Loop
    End With

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

    If Not bln_AlreadyInitialized Then
        InitializeBidBaseDLL
    End If

    If Not bln_HandValInitialized Then
        InitializeHandVal
    End If

    Set rs_HandVal = db_HandVal.OpenRecordset( _
                     "SELECT * FROM HandVal " & _
                     "WHERE Name='" & sysName & "' " & _
                     "ORDER BY ID;", _
                     dbOpenSnapshot)

    Eval_CalcHandShape
    Eval_SortSuitsByHCP
    If Not bln_AlreadyHaveOPsData Then
        GetOtherPlayersData
    End If
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
        DoEvents
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
        Exit Sub ' Already done. '
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
        DoEvents
    Loop
End Sub


Private Sub Eval_SetUpFieldVariables()
    With rs_HandVal
        str_ValSpec = Trim$("" & !Value)
        sng_ValSpec = Val(str_ValSpec)
        str_FlagSpec = Trim$("" & !Flag)
        str_PosSpec = Trim$("" & !position)
        str_HandSpec = Trim$("" & !hands)
        str_FitSpec = Trim$("" & !fit)
        str_SuitSpec = Trim$("" & !Suit)
        str_CardsSpec = Trim$("" & !Cards)
        str_LenSpec = Trim$("" & !Len)
    End With

    If Left$(str_LenSpec & " ", 1) = "<" Then
        int_LenSpec = Val(Mid$(str_LenSpec, 2))
    Else
        int_LenSpec = Val(str_LenSpec)
    End If
    int_SuitSpec = 0
    int_SuitNum = 0
    str_CardsSpecCombo = ""
    bln_ValAdded = False
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
        op = Mid$(f, n, 1)
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
        If fCompare <> "" Then
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
        ' This spec just means that the specified '
        ' player has shown/bid at least one suit. '
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
            MsgBox "Stop Eval_HandsSpecCalcOk 1"
            'Eval_HandsSpecCalcOk = (t_Str(i) <> Right$(s, 1)) Or t_Str(i - 1) = ""
        Else
            j = InStr("=<>", Mid$(s, 2, 1))
            If i > 0 And j > 0 And k > 0 Then
                i = i - 1
                Select Case Mid$(s, 2, 1)
                Case Is = "="
                    MsgBox "Stop Eval_HandsSpecCalcOk 2"
                    'Eval_HandsSpecCalcOk = (t_Str(i) = Right$(s, 1))
                Case Is = "-"
                    MsgBox "Stop Eval_HandsSpecCalcOk 2"
                    'Eval_HandsSpecCalcOk = (t_Str(i) <> Right$(s, 1)) Or t_Str(i) = ""
                Case Is = "<"
                    MsgBox "Stop Eval_HandsSpecCalcOk 2"
                    'Eval_HandsSpecCalcOk = (InStr("BWIOPSVX", t_Str(i)) < _
                                            InStr("BWIOPSVX", Right$(s, 1))) Or t_Str(i) = ""
                Case Is = ">"
                    MsgBox "Stop Eval_HandsSpecCalcOk 2"
                    'Eval_HandsSpecCalcOk = (InStr("BWIOPSVX", t_Str(i)) > _
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
                ' int_HCP(0) = total HCP for hand.  '
                ' int_HCP(1-4) = HCP for each suit. '
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
            MsgBox "Stop Eval_HandsSpecCalcOk 2"
            'Eval_HandsSpecCalcOk = (Val(str_LHOs_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: L<20 = LHO has <15 VPs '
            MsgBox "Stop Eval_HandsSpecCalcOk 2"
            'Eval_HandsSpecCalcOk = (Val(str_LHOs_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If

        If Len(s) = 2 Then ' LN, LS, LC, LR, LF -- alone. '
            Select Case r
            Case Is = "C"
                MsgBox "Stop Eval_HandsSpecCalcOk 2"
                'Eval_HandsSpecCalcOk = (int_LHOs_CueBid > 0) Or neg
            Case Is = "F"
                MsgBox "Stop Eval_HandsSpecCalcOk 2"
                'Eval_HandsSpecCalcOk = (int_LHOs_Fit(1) > 0) Or neg
            Case Is = "N"
                MsgBox "Stop Eval_HandsSpecCalcOk 2"
                'Eval_HandsSpecCalcOk = bln_LHO_Bid_NT Or neg
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
                    MsgBox "Stop Eval_HandsSpecCalcOk 2"
                    'Eval_HandsSpecCalcOk = (int_LHOs_CueBid = j)
                Case Is = "LF"
                    MsgBox "Stop Eval_HandsSpecCalcOk 2"
                    'Eval_HandsSpecCalcOk = (int_LHOs_Fit(1) = j Or _
                                            int_LHOs_Fit(2) = j)
                Case Is = "LH"
                    MsgBox "Stop Eval_HandsSpecCalcOk 2"
                    'Eval_HandsSpecCalcOk = (int_LHOs_HelpSuit = j)
                Case Is = "LR"
                    MsgBox "Stop Eval_HandsSpecCalcOk 2"
                    'Eval_HandsSpecCalcOk = (int_LHOs_RebidSuit(1) = j Or _
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
            'Stop
            'Eval_HandsSpecCalcOk = (Val(str_RHOs_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: R<20 = RHO has <15 VPs '
            'Stop
            'Eval_HandsSpecCalcOk = (Val(str_RHOs_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If

        If Len(s) = 2 Then ' RN, RS, RC, RR, RF -- alone. '
            Select Case r
            Case Is = "C"
                'Stop
                'Eval_HandsSpecCalcOk = (int_RHOs_CueBid > 0) Or neg
            Case Is = "F"
                'Stop
                'Eval_HandsSpecCalcOk = (int_RHOs_Fit(1) > 0) Or neg
            Case Is = "N"
                'Stop
                'Eval_HandsSpecCalcOk = bln_RHO_Bid_NT Or neg
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
                    'Stop
                    'Eval_HandsSpecCalcOk = (int_RHOs_CueBid = j)
                Case Is = "RF"
                    'Stop
                    'Eval_HandsSpecCalcOk = (int_RHOs_Fit(1) = j Or _
                                            int_RHOs_Fit(2) = j)
                Case Is = "RH"
                    'Stop
                    'Eval_HandsSpecCalcOk = (int_RHOs_HelpSuit = j)
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
            'Stop
            'Eval_HandsSpecCalcOk = (Val(str_Pards_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: P<20 = Pard has <15 VPs '
            'Stop
            'Eval_HandsSpecCalcOk = (Val(str_Pards_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If

        If Len(r) = 1 Then
            Select Case r  ' PC, PF, PN, PR, PS -- alone '
            Case Is = "C"
                Eval_HandsSpecCalcOk = (int_Pards_CueBid > 0) Or neg
            Case Is = "F"
                'Stop
                'Eval_HandsSpecCalcOk = (int_Pards_Fit(1) > 0) Or neg
            Case Is = "N"
                'Stop
                'Eval_HandsSpecCalcOk = bln_Pard_Bid_NT Or neg
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
                    'Stop
                    'Eval_HandsSpecCalcOk = (int_Pards_Fit(1) = j Or _
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
            MsgBox "Stop Eval_HandsSpecCalcOk 3"
            'Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = 0 Or _
                            (int_Pards_Suit(1) <> int_LHOs_CueBid And _
                             int_Pards_Suit(2) <> int_LHOs_CueBid))
        Case Is = "P<>RC"
            MsgBox "Stop Eval_HandsSpecCalcOk 3"
            'Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = 0 Or _
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
            'Stop
            'Eval_HandsSpecCalcOk = _
                            (int_Pards_Suit(1) = int_LHOs_CueBid And _
                             int_Pards_Suit(2) = int_LHOs_CueBid)
        Case Is = "P=RC"
            'Stop
            'Eval_HandsSpecCalcOk = _
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
                'Stop
                'Eval_HandsSpecCalcOk = (int_My_CueBid > 0) Or neg
            Case Is = "F"
                'Stop
                'Eval_HandsSpecCalcOk = (int_My_Fit(1) > 0) Or neg
            Case Is = "N"
                'Stop
                'Eval_HandsSpecCalcOk = bln_I_Bid_NT Or neg
            Case Is = "R"
                'Stop
                Eval_HandsSpecCalcOk = (int_My_RebidSuit(1) > 0) Or neg
            Case Is = "S"
                'Stop
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
        If int_RHOs_Suit(1) > 0 Or _
           int_LHOs_Suit(1) > 0 _
        Then
            Eval_SuitSpecOk = True
        End If
        Exit Function

    ElseIf s = "ER" Then
        If int_RHOs_RebidSuit(1) > 0 Or _
           int_LHOs_RebidSuit(1) > 0 _
        Then
            Eval_SuitSpecOk = True
        End If
        Exit Function

    ElseIf s = "-M" Or _
           s = "-P" Or _
           s = "-F" Or _
           s = "-E" Or _
           s = "-ER" _
    Then
        ' These are checked in Eval_CardsSpecOk. '
        Eval_SuitSpecOk = True
        Exit Function

    ElseIf s1 = "F" Then ' Fit suit. '
        int_SuitSpec = int_Fit_Suit(1)

    ElseIf s1 = "M" Then ' My suit. '
        int_SuitSpec = int_My_Suit(1)

    ElseIf s1 = "P" Then ' Pard's suit. '
        'a = Right$(" " & t_Len(4 + int_Pards_Suit(1)), 1)
        'Stop
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
            MsgBox "Stop Eval_SuitSpecCalcOk 1"
            'a = t_Len(4 + int_Pards_Suit(1))
        Else
            Exit Function
        End If

    ElseIf Left$(s & " ", 2) = "PS" Or _
           Mid$(s & "  ", 2) = "PS" _
    Then
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
        If int_Pards_SS = 0 And _
           Left$(s, 1) <> "-" _
        Then
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

        If Not (s = "PS<>L" And _
                (l = False Or l0)) _
        Then
            ' Spec is PSS <> LHO's long suit. '
            ' If L is true, get out:          '
            Exit Function

        ElseIf Not (s = "PS=L" And _
                    l = True) _
        Then
            ' Spec is that PSS must = LHO. '
            Exit Function

        ElseIf Not (s = "PS<>R" And _
                    (r = False Or r0)) _
        Then
            Exit Function

        ElseIf Not (s = "PS=R" And _
                    r = True) _
        Then
            Exit Function

        ElseIf Not (s = "PS<>E" And _
                    (n = False Or (l0 And r0))) _
        Then
            ' Either Pard's suit matches the opp's '
            ' or opp's don't have a suit.          '
            Exit Function

        ElseIf Not (s = "PS=E" And _
                    n = True) _
        Then
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
            ' All HCPs have been sorted were 4=most, then 3,2,1. '
            ' So if int_HCPn(1)=0 then all HCPs are in 2-4. '
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

    ElseIf s = "Clbs" Or _
           s = "Dias" Or _
           s = "Hrts" Or _
           s = "Spds" _
    Then
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

    If int_Pards_Suit(1) = 0 Or str_FitSpec = "" Then
        Eval_FitSpecOk = (str_FitSpec = "" Or _
                          str_FitSpec = "N")
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
                If int_Fit_Suit(1) = i Or _
                   int_Fit_Suit(2) = i Or _
                   int_Fit_Suit(3) = i _
                Then
                    ok = True
                End If
                ' Stand-alone pgm code:
                'If Right$(" " & t_Len(4 + i), 1) = "F" Then
                '    ok = True
                'End If

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
        str_CardsSpec = Trim$("" & rs_HandVal!Cards)

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
                DoEvents
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

            DoEvents
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
            DoEvents
        Loop

        If str_LenSpec <> "" And _
           Trim$("" & rs_HandVal!Cards) = "" _
        Then
            If Right$("  " & str_ValSpec, 2) = "*n" Then
                n = int_SuitLen(int_SuitNum)
                If rs_HandVal!fit = "Y" Or _
                    rs_HandVal!Suit = "F" _
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
            DoEvents
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


Private Function Eval_Pts_Ok() As Boolean
    ' This is no longer used. ' in the Bidder DLL '
    ' but not in the HandVal program.'
    Dim ok As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim k As Single
    Dim s As String
    Dim vs As String
    Dim vn As String

    ok = True
    With rs_BidTable
    '              '
    s = Trim$("" & !ValSystem)
    vs = Trim$("" & !SuitVal)
    vn = Trim$("" & !NT_Val)

    If s <> "" And (vs <> "" Or vn <> "") Then
        ' Val.System & points specified. '
        If str_ValSystem(1) = "" Then
            ' Load the system names: '
            InitializeHandVal
            With rs_HandValBidReq
                .MoveFirst
                i = 0
                Do While Not .EOF
                    i = i + 1
                    str_ValSystem(i) = !EvalName
                    .MoveNext
                    DoEvents
                Loop
            End With
        End If

        i = 1
        Do While str_ValSystem(i) <> s And i < 20
            i = i + 1
        Loop
        If str_ValSystem(i) <> s Then
            MsgBox "Hand evaluation system not found:" & vbCr & s
            bln_Cancel = True
            Exit Function
        End If

        If (vs <> "" And sng_ValPoints(i, 1) = 0) Or _
           (vn <> "" And sng_ValPoints(i, 2) = 0) _
        Then
            ' Hasn't been calculated yet.'
            Eval_A_Hand s
            If bln_Cancel Then Exit Function
        End If

        If vs <> "" Then
            ok = Eval_PtsOk(sng_ValPoints(i, 1), vs, 5) ' This is the next sub, not this one. '
            If Not ok Then Exit Function
        End If

        If vn <> "" Then
            ok = Eval_PtsOk(sng_ValPoints(i, 2), vn, 5)
            If Not ok Then Exit Function
        End If

    End If
    '              '
    End With

    Eval_Pts_Ok = ok
End Function


Private Function Eval_PtsOk(evalPts As Single, _
                            specs As String, _
                            analy_Line As Integer) _
                            As Boolean
    ' This is used in the Bidder DLL '
    ' but not in the HandVal program.'
    Dim intMin As Integer
    Dim intMax As Integer
    Dim intDash As Integer

    ' specs can have the following types of formats:
    '   >10   '
    '   6-8   '
    '   <20   '
    '   12    '
    '   12+   '

    intDash = InStr(specs, "-")

    ' Compute min and max for test: '

    If Left$(specs, 1) = "<" Then 'e.g.: <10
        ' Should this be 0? '
        intMax = Val(Mid$(specs, 2)) - 1

    ElseIf Left$(specs, 1) = ">" Then 'e.g.: >10
        intMax = 9999
        intMin = Val(Mid$(specs, 2))

    ElseIf intDash > 0 Then
        If intDash = Len(specs) Then
            'e.g.: 6- (normally use <7 instead)
            intMax = Val(specs)
            intMin = 0

        Else                 'e.g.: 6-10
            intMax = Val(Mid$(specs, intDash + 1))

        End If

    ElseIf Right$(specs, 1) = "+" Then 'e.g.: 6+
        intMax = 9999

    ElseIf Left$(specs, 1) <> "<" And _
           Left$(specs, 1) <> ">" _
    Then                                 'e.g.: 10
        intMax = intMin

    End If

    ' Do the comparison: '

    If evalPts >= intMin And _
       evalPts <= intMax _
    Then
        Eval_PtsOk = True

    ElseIf bln_ShowAnalysis Then
        With frm.lstAnalysis
            .ListIndex = analy_Line
            .List(analy_Line) = .List(analy_Line) & _
                        " -- NO MATCH:  Hand has " & _
                        evalPts
        End With
    End If

End Function


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

    If ("" & rs_HandVal!Value) = "" And int_FlagNum > 0 Then
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
            Select Case rs_HandVal!NT_or_Suit
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
        If int_FlagNum > 0 Then
            int_Flag(int_FlagNum) = int_Flag(int_FlagNum) + 1
        End If
    Else
        bln_ValAdded = False
    End If

End Sub


Private Sub SaveBidEntry()
    Dim b As String
    Dim n As Integer
    Dim s As String
    Dim sn As Integer
    Dim bon As String
    Dim pb As String

    'If int_OptAvg = 0 Then MsgBox "Stop SaveBidEntry 1"
    AnalyzeHand

    On Local Error GoTo SaveEntryError

    ' Compute Bid Order Number: '
    pb = str_PriorCalls
    Do While Left$(pb & "  ", 2) = "P-"
        pb = Mid$(pb, 3)
    Loop
    b = str_BidToSave
    If b = "P" Then
        bon = "999900"
    ElseIf b = "D" Or b = "R" Then
        bon = "980000"
    Else
        n = Val(b)
        s = Mid$(b, 2)
        sn = InStr("CDHSN", s)
        If n = 0 Or n > 7 Or sn = 0 Then
            bon = "999900"
        Else
            bon = "9" & sn & n & "000"
        End If
    End If

    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidTable = db_BidBase.OpenRecordset("Bids")
    With rs_BidTable
    ''''''''''''''''
    .AddNew
    ' Unused fields must have "" assigned to them: '
    !pct_Used = ""
    !RoundOfBidding = "" & int_BidRound
    !player = "" & int_PosNum
    !PB_Sort_Order = PB_To_SortOrd_Code(pb)
    !Prior_Bids = pb
    !Bid_Order_Number = bon
    !Bid = str_BidToSave
    
    !Cust_Pts = ""
    !A_Cust = ""
    !K_Cust = ""
    !Q_Cust = ""
    !J_Cust = ""
    !T_Cust = ""
    !Winrs = ""
    !Losers = ""
    !QTs = ""
    !Stopped_Suits = ""
    !Shape = ""
    !Intermed = ""
    !xNum = ""
    !xHCP = ""
    !cmNum = ""
    !cmHCP = ""
    !Vul = ""
    !pos = ""
    !Artificial = ""
    !Alert = ""
    !Explanation = ""
    
    If bon = "999900" Then ' "P" catch-all '
        !HCP = ""
        !Dist_Pts = ""
        !ConvID = 4  ' ID# for black Convention '
        !HCP = ""
        !Dist_Pts = ""
        !SNum = ""
        !SHCP = ""
        !HNum = ""
        !HHCP = ""
        !DNum = ""
        !DHCP = ""
        !CNum = ""
        !CHCP = ""
        !Test_Hand = str_PlayersHand
        !Asks = ""
        !DateOfEntry = Date$
        GoTo SaveEntry
    End If
    
    If Left$(b & "   ", 2) = "A-" Then
        !ConvID = 578 ' Ace asking '
        b = Mid$(b, 3)
    ElseIf Left$(b & "   ", 2) = "Q-" Then
        !ConvID = 173 ' Control asking bid '
        b = Mid$(b, 3)
    Else
        !ConvID = 4 ' ID# for blank Convention '
    End If
    
    !HCP = "" & int_HCP(0)
    !Dist_Pts = "T" & int_HCP(0) + int_DistribPts(0)
    !SNum = "" & int_SuitLen(4)
    !SHCP = "" & int_HCP(4)
    !HNum = "" & int_SuitLen(3)
    !HHCP = "" & int_HCP(3)
    !DNum = "" & int_SuitLen(2)
    !DHCP = "" & int_HCP(2)
    !CNum = "" & int_SuitLen(1)
    !CHCP = "" & int_HCP(1)
    !Test_Hand = str_PlayersHand
    !Asks = ""
    !DateOfEntry = Date$
    s = ""
    If int_OptAvg > 0 Then
        n = int_OptAvg Mod 10 ' e.g.: 53 mod 10 = 3 = hearts = 5H '
        s = Mid$("CDHSN", n, 1)
        n = int_OptAvg \ 10 ' 53 \ 10 = 5 '
        s = n & s
    End If
    
SaveEntry:
    
    !Best_Contract = s
    !Num_Used = 0
    !LHO = lng_LHO
    !pard = lng_Pard
    !RHO = lng_RHO
    !Source = "Deal #" & lng_DealNum
    .Update

    ' Make this entry the current record: '
    .Bookmark = .LastModified
    lng_ID = !id
    ' Note that if Editor is open, its data '
    ' control and grids will not be updated '
    ' even though if you enter this hand in '
    ' Find Hand, it will find this entry.   '
    ' The only way I've found to resolve    '
    ' this is to import the entry into the  '
    ' editor, put it into the Input Boxes,  '
    ' and save the entry there: '
    ChDir App.Path
    Dim fiNum As Integer
    fiNum = FreeFile
    Open "EditPractice.txt" For Output As #fiNum
    Print #fiNum, str_PlayersHand
    Print #fiNum, str_PriorCalls
    Close #fiNum
    '''''''''''''''''''''''''''''''''''''''''''

    If int_OptAvg = 0 Then
        'MsgBox "Entry added. Click Edit to see the full entry."
    Else
        str_Call(int_CallNum) = str_BidToSave
    End If
    ''''''''
    End With
    On Local Error GoTo 0
Exit Sub

SaveEntryError:
    If Err.Description <> "" Then
        MsgBox Err.Description & vbCr & _
               "in Sub SaveBidEntry."
    End If
    Stop
    Resume 'Next
    On Local Error GoTo 0
End Sub


Private Function PB_To_SortOrd_Code(ByVal pb As String) As String
    Dim v As String
    Dim v1 As String
    Dim i As Integer
    Dim n As Integer
    Dim a As String

    ' See a discussion of this code in        '
    ' CardShark_BidBase_Programming_Notes.htm '
    ' Note that this same Function is also in '
    ' the editor code with the same name, so  '
    ' any changes in one should be copied to  '
    ' the other. '
    If Val(pb) = 0 And pb <> "" Then
        PB_To_SortOrd_Code = LCase$(pb)
        Exit Function
    End If

On Local Error GoTo SortOrdCodeErr
    v = ""
    pb = UCase$(Trim$(pb))

    If pb = "" Then
        PB_To_SortOrd_Code = Space$(34) & "0"
        Exit Function
    End If

    If InStr(pb, "@") > 0 Then
        ' Slam bidding entries. '
        ' See the Notes file.   '
        PB_To_SortOrd_Code = pb
        Exit Function
    End If

    Do While pb <> ""
        i = InStr(pb, "-")
        If i > 0 Then
            ' Extract next bid from BidSequence:
            v1 = Left$(pb, i - 1)
            pb = Mid$(pb, i + 1)

        Else
            v1 = pb
            pb = ""
        End If

        v1 = Replace(v1, Chr$(13), "")
        v1 = Replace(v1, Chr$(10), "")
        If Len(v1) = 1 Then
            v = v & (InStr(" DRP", v1) - 1)

        ElseIf v1 <> "" Then
            'If Len(v1) > 2 Then ' Ex: 4N@30 '
            '    a = Mid$(v1, 3) '       @30 '
            '    v1 = Left$(v1, 2) '   4N    '
            'End If

            ' Find number of bid in the list of all bids:
            n = InStr(" 1C1D1H1S1N2C2D2H2S2N3C3D3H3S3N" & _
                      "4C4D4H4S4N5C5D5H5S5N6C6D6H6S6N" & _
                      "7C7D7H7S7N", v1) / 2
            If n > 36 Then
                MsgBox "Invalid prior bid."
                Exit Function
            End If
            ' Convert decimal value of bid # to code character:
            v = v & Mid$("456789ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]", n, 1)
            ' The following was not doing '
            ' what I wanted anyway. '
            'If Left$(v, 1) = "@" Then
            '    v = v & a
            'End If
        End If
        DoEvents
    Loop
    ' For BidBase purposes, I need the codes in alphanumeric '
    ' order, so I have to pad them on the left with spaces.  '
    ' For example, "413A" is less than "46" when comparing   '
    ' character strings, so I have to convert "46" to "  46" '
    ' so that "  46" will compare as being less than "413A": '
    PB_To_SortOrd_Code = Right$(Space$(35) + v, 35)

SortOrdCodeErr:
    If Err.Number > 0 Then
        MsgBox Err.Description & " creating Sort Order Code."
        Stop
        Resume ' next
    End If
    On Local Error GoTo 0
End Function


Private Sub GetDDA()
    ' Load array int_OptByStrainAvg(4, 5, 2)     '
    ' Where 4 = player #: S=1 E=2 N=3 W=4     '
    '       5 = strain #: C=1 D=2 H=3 S=4 N=5 '
    '       2 = # iterations: 1=11, 2=101,    '
    '       or 0 for the original DDA results '
    ' Example: Opt for North's NT for 11      '
    '          iterations: (3, 5, 1) = 2      '
    ' This was saved in the calling program   '
    ' just before calling this bidder when    '
    ' DDA has been calculated. The array is   '
    ' used in Sub ComputeOptBid below.        '
    Dim p As Integer
    Dim pn As Integer
    Dim s As Integer
    Dim i As Integer
    Dim fi As Integer
    
    On Local Error GoTo GetDDA_Error
    fi = FreeFile()
    Open "DDAresults.txt" For Input As #fi
        For i = 0 To 2
          For p = 1 To 4
            If p = 1 Then
                pn = 1 ' S '
            ElseIf p = 2 Then
                pn = 3 ' N '
            ElseIf p = 3 Then
                pn = 2 ' W '
            Else
                pn = 4 ' E '
            End If
                    
            'For s = 1 To 5
                    Input #fi, int_OptByStrainAvg(pn, 0, i)
            'Next
          Next
        Next

GetDDA_Error:
    If Err.Number > 0 Then
        MsgBox Err.Description & " in Sub GetDDA." & _
               ""
        Stop
        Resume ' next
    End If
    
    Close #fi
    On Local Error GoTo 0
End Sub


Private Sub ComputeOptBid()
    ' No bid was found in the database. Try computing one. '

    ' Us: '
    Dim optSuitNumUs As Integer ' My Optimum suit number. (CDHSN = 1 to 5) '
    Dim optSuitStrUs As String ' My Optimum suit string. (CDHSN) '
    Dim optSuit4Pard As String
    Dim optSuitNum4Pard As Integer
    Dim optMaxLvlUs As Integer ' Highest level we can make, game or more. '
    Dim optGameLvlUs As Integer ' Lowest Opt game level: 5(C/D), 4(H/S), 3(N) '
    Dim optSuitLen4Me As Integer
    Dim optMaxUs As Integer  ' 11-75 : Our maximum bid. Also int_OptBid. '
    Dim optGameUs As Integer ' 51-52,43-44,35 : Our optimum game bid. '
    Dim hiBidLvlUs As Integer ' 0-7 '
    Dim hiBidSuitUs As Integer ' 1-5 : C-N '
    Dim hiOptBidLvlUs As Integer
    Dim hiBidUs As Integer ' 11-75 : 1C-7N : Our hi bid so far. '
    Dim maxSac As Integer ' Maximum # of tricks we can sac profitably. '
    Dim oppOptSuitHigher As Integer ' =1 if opp's opt suit is higher than ours. '
    Dim nextHiBidLvlUs As Integer
    ' Opp: '
    Dim optSuitNumOpp As Integer
    Dim optSuitStrOpp As String
    Dim optMaxLvlOpp As Integer
    Dim optGameLvlOpp As Integer
    Dim optMaxOpp As Integer
    Dim optGameOpp As Integer
    Dim hiBidLvlOpp As Integer
    Dim hiBidSuitOpp As Integer
    Dim hiOptBidLvlOpp As Integer
    Dim hiBidOpp As Integer
    ' Both: '
    Dim hiBidLvl As Integer ' Highest level bid so far: 1-7 '
    Dim hiBidSuit As Integer ' Highest suit bid so far: 1-5 : club - notrump '
    Dim allBids(4) As String ' String of all bids made by each hand, numbered CCW.
    Dim pardsLastCall As String
    Dim RHOsLastCall As String
    Dim LHOsLastCall As String
    
    Dim diff As Integer ' diff between last bid and optimum '
    Dim bidPt As String ' Bid Point: Where the code is that made the bid. '
    Dim bidBlok As String ' Block of code where bid was made. '
    Dim redoBidPt As String ' Return to the bidPt. '
    Dim unbal As Boolean
    
    Dim p As Integer ' partner's position # '
    Dim m As Integer ' my position # '
    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim s As Integer ' suit pointer '
    Dim n As Integer ' bid level '
    Dim b As String ' new bid '
    Dim X As String
    'If rs_BidTable!id = 17864 Then Stop
    ' Module variables: '
    ' int_HCP(0) = current bidder's total HCP.   '
    ' int_BidNum = Current bidder's bid number.  '
    ' int_CallNum = the call num being computed. '
    ' str_Call() = an array of prior calls.      '
    ' str_Vul = vul for current bidder: F,U,B,N  '
    ' str_PriorBids = a string of prior calls.   '
    ' int_OptAvg = optimum bid (e.g.: 13=1H, 15=1N) '
    ' int_Opt4PardAvg = optimum for partner.        '
    ' int_Opt4OppAvg = optimum for opponents.       '
    '                                            '
    ' bln_Pard_Bid_NT As Boolean                 '
    ' int_Pards_CueBid As Integer                '
    ' int_Pards_Fit(2) As Integer                '
    ' int_Pards_HelpSuit As Integer              '
    ' int_Pards_Pos As Integer ' Partner's pos.  '
    ' int_Pards_RebidSuit(2) As Integer          '
    ' int_Pards_SS As Integer ' Pard's short suit'
    ' int_Pards_Suit(4) As Integer ' C=1...S=4   '
    ' int_Pards_SuitLen(4) As Integer            '
    ' int_Pards_Strength As String               '

'On Local Error GoTo ComputeBidError
On Local Error GoTo 0

    If int_BidNum = 3 Then
        pardsLastCall = str_Bid(int_BidNum - 2)
    ElseIf int_BidNum = 2 Then
        RHOsLastCall = str_Bid(int_BidNum - 1)
    ElseIf int_BidNum > 1 Then ' Means it's 4+ '
        LHOsLastCall = str_Bid(int_BidNum - 3)
        pardsLastCall = str_Bid(int_BidNum - 2)
        RHOsLastCall = str_Bid(int_BidNum - 1)
    End If

    ' Copy str_Call(1-50) to allBids(1-4)   '
    ' with P, D, R excluded. This makes it  '
    ' easier to search for a suit bid by a  '
    ' player using Instr.  Example:         '
    '                  S   W   N   E        '
    ' str_Call(1-6) =     1C,  D, 1H,       '
    '                 1S, 2H, 2S, ??        '
    ' 1st bidder = W   4   1   2   3        '
    '                  S   W   N   E        '
    ' allBids index =  2   3   4   1        '
    ' If I'm always allBids index 1 and the '
    ' indices go CW from there, it will be  '
    ' easier to figure the index # for any  '
    ' player. For this example, I'm East:   '
    ' allBids(1)=1H (2)=1S (3)=1C2H (4)=2S  '
    ' Note that North's D is not included.  '
    ' My pard is always allBids(3), RHO = 4 '
    ' and LHO = 2. Using Instr to chk bids: '
    ' (Instr(allBids(3), "C") > 0) = True   '
    ' (Instr(allBids(3), "H") > 0) = True   '
    ' My partner has bid hearts and clubs.  '
    ' Find the opening bidder's number: '
    j = (int_BidNum Mod 4)
    If j = 0 Then j = 4
    ' If j = 1, then I opened.    I'm always allBids(1)    '
    ' If j = 2, then RHO opened.  allBids(4) = str_Call(1) '
    '           RHO must = allBids(4) so that I'm 1.       '
    ' If j = 3, then pard opened. allBids(3) = str_Call(1) '
    '           Pard must be allBids(3), RHO=4, I'm 1.     '
    ' If j = 4, then LHO opened.  allBids(2) = str_Call(1) '
    '           LHO = allBids(2), pard=3, RHO=4, I'm 1.    '

    For i = 1 To 4 ' allBids() indexes '
        For n = j To int_BidNum Step 4
            If Len(str_Bid(n)) > 1 Then ' Don't include P, D, R. '
                allBids(i) = allBids(i) & str_Bid(n)
            End If
        Next
        j = j + 1
        If j = 5 Then j = 1
    Next

    ' I sometimes refer to the current bidder as "I", "me" '
    ' or "my" even though this sub is bidding all 4 hands. '

    GetDDA ' The Sub immediately before this one. '

    If int_OptAvg <> int_Opt4PardAvg Then
        ' Edit of below:                      '
        ' If my opt is 6D and 5S and pard's   '
        ' is 5D and 5S and he bids D first,   '
        ' then I can't be declarer in D and   '
        ' so S is our optimum. '
        ' Pard's optimum is higher than mine, '
        ' (meaning that he has to be declarer)'
        ' so if he bid the suit first then he '
        ' would be declarer,  so his optimum  '
        ' should be used, but if I bid it 1st '
        ' then his optimum won't matter since '
        ' he could not be the declarer in his '
        ' optimum suit. Example:  Optimum may '
        ' be 9 tricks in spades if I declare  '
        ' but 10 tricks if he is declarer.    '
        ' int_OptAvg & int_Opt4PardAvg are like 34  '
        ' and 44 where the first digit is the '
        ' level and the second is the strain. '
        ' int_OptAvg Mod 10 = the suit number.   '
        ' For 34, suit # is 4 = S, so 34 = 3S '
        ' More often, we will have the same   '
        ' opt suit but one of us will be able '
        ' to take an extra trick as declarer. '
        ' So if pard's optimum is 4S but mine '
        ' is 3S, then if I bid S first, our   '
        ' opt will be 3S. It's possible that  '
        ' we could have different optimums in '
        ' different suits. Say that pard can  '
        ' make 5D but only 3S while I can make'
        ' 4S and only 4D (as declarer).  If I '
        ' bid D first and he bids S first, our'
        ' optimums in the 2 suits are 4D (130 '
        ' points) and 3S )170 points) so our  '
        ' optimum contract would be 3S rather '
        ' than 4D. Highly unlikely to be bid  '
        ' that way, but still possible. I have'
        ' not bothered to account for that in '
        ' the code yet. '
        ' Here is a more complex deal:        '
        '     C   D      H  S      N          '
        ' E:  8  10(130) 6  9(170) 8(120)     '
        ' W:  8  10      6  9      9(480/680) '
        ' E: 1D-1S, ?? - BB bid 1N, but if E  '
        ' bids N first then opt is S, but E   '
        ' has AQ-AJ53-J765-Q74. Too weak for  '
        ' 2H. Can't bid anything but 1N. With '
        ' game no longer possible, West bid   '
        ' 2S to show an extra spade and with  '
        ' 3N by W no longer possible, the opt '
        ' becomes 2S making 3. E must pass 2S.'
        optSuitNumUs = int_OptAvg Mod 10 ' My opt suit number. '
        If optSuitNumUs > 0 Then
           optSuitStrUs = Mid$("CDHSN", optSuitNumUs, 1)
           optSuitNum4Pard = int_Opt4PardAvg Mod 10
        End If
        
        If optSuitNum4Pard > 0 Then
            optSuit4Pard = Mid$("CDHSN", optSuitNum4Pard, 1)
        End If
        
        If optSuitNumUs < optSuitNum4Pard Then
            ' My opt is lower than pard's.   '
            ' See who bid pard's suit first. '
            i = InStr(allBids(1), optSuit4Pard) ' I'm always allBids(1) '
            j = InStr(allBids(3), optSuit4Pard) ' and pard is always 3. '
            ' If my opt is 2N and 3S and pard's is 3N '
            ' and 3S and he bid NT first, then I have '
            ' to change my opt to his. '
            If i > 1 Then
                m = Val(Mid$(allBids(1), i - 1, 1)) ' My first bid level. '
            Else
                m = 99 ' We're looking for the lowest bid.     '
                       ' Since I didn't bid it, 99 is highest. '
            End If
            If j > 1 Then
                p = Val(Mid$(allBids(3), j - 1, 1)) ' Pard's first level. '
            Else
                p = 99
            End If
            If p > m Then
                ' Pard's opt is higher than mine. '
                ' Pard bid his opt suit before me '
                ' so change my opt to his. '
                int_OptAvg = int_Opt4PardAvg
                optSuitNumUs = int_OptAvg Mod 10
                optSuitStrUs = Mid$("CDHSN", optSuitNumUs, 1)
            Else
                ' I bid pard's opt first, so when it '
                ' is his next turn to bid, he will   '
                ' have to change his opt to mine. '
            End If
        
        ElseIf optSuitNumUs > optSuitNum4Pard Then
            ' My opt suit is higher. See who bid it first: '
            i = InStr(allBids(1), optSuitStrUs) ' I'm always allBids(1) '
            j = InStr(allBids(3), optSuitStrUs) ' and pard is always 3. '
            ' If my opt is 3N and 3S and pard's is  '
            ' 2N and 3S and pard bid NT first, then '
            ' NT can't be the opt. S is our opt. '
            If i > 1 Then
                m = Val(Mid$(allBids(1), i - 1, 1)) ' My first bid level. '
            Else
                m = 99 ' We're looking for the lowest bid.     '
                       ' Since I didn't bid it, 99 is highest. '
            End If
            If j > 1 Then
                p = Val(Mid$(allBids(3), j - 1, 1)) ' Pard's first level. '
            Else
                p = 99
            End If
            If p < m Then
                ' Pard's opt is higher than mine. '
                ' Pard bid his opt suit before me '
                ' so change my opt to his. '
                int_OptAvg = int_Opt4PardAvg
                optSuitNumUs = int_OptAvg Mod 10
                If optSuitNumUs = 0 Then
                    optSuitStrUs = ""
                Else
                    optSuitStrUs = Mid$("CDHSN", optSuitNumUs, 1)
                End If
            Else
                ' I bid pard's opt first, so when it '
                ' is his next turn to bid, he will   '
                ' have to change his opt to mine. '
            End If

        Else ' We have the same opt suit. '
            i = InStr(allBids(1), optSuit4Pard)
            j = InStr(allBids(3), optSuit4Pard)
            If j > 1 Then
                ' Pard bid it. If I bid it too, '
                ' see who bid first which would '
                ' be the lower bid. '
                If i > 1 Then
                    m = Val(Mid$(allBids(1), i - 1, 1)) ' My bid level. '
                Else
                    m = 99
                End If
                p = Val(Mid$(allBids(3), j - 1, 1)) ' Pard's level. '
                ' Note that if D(ouble) was included '
                ' in allBids, and D(iamonds) was the '
                ' suit, it would take more code to   '
                ' figure out which the D was for.    '
                If p < m Then
                    ' Pard's bid was lower than mine, meaning '
                    ' that he bid his opt suit first so his   '
                    ' optimum must be used since he will be   '
                    ' the declarer. '
                    int_OptAvg = int_Opt4PardAvg
                    optSuitNumUs = int_OptAvg Mod 10
                    optSuitStrUs = Mid$("CDHSN", optSuitNumUs, 1)
                Else
                    ' I bid his suit first, so '
                    ' no change is necessary.  '
                End If
            End If
        End If

    ElseIf int_OptAvg > int_Opt4PardAvg Then
        ' My optimum is higher than pard's, so change his to mine. '
        ' See if I bid the optimum first: '
        Stop
    End If

    ' Set up Optimum variables: '
    optSuitNumUs = int_OptAvg Mod 10
    If optSuitNumUs = 0 Then
        ' The partner and i cannot take >6 '
        ' tricks in any strain, so exit.'
        GoTo SaveBid
    End If
    optSuitStrUs = Mid$("CDHSN", optSuitNumUs, 1)
    If optSuitNumUs < 5 Then
        optSuitLen4Me = int_HCP(optSuitNumUs)
    Else ' My opt "suit" is NT which has no "length". '
        optSuitLen4Me = 99
    End If
    optMaxLvlUs = int_OptAvg \ 10
    ' Ex: int_OptAvg = 43: optMaxLvlUs = 43 \ 10 = 4 '

    ' Don't need to bid optimum w/o competition '
    ' if game is less than optimum. Ex: Opt=5S, '
    ' g=4S. But if Opt=6S or 7S, must bid it.   '
    ' If Opt < game, then we can't make game but'
    ' don't have to bid 2S/3S unless forced to. '
    ' Example: optimum=3S. Bidding: 1S-P-P      '
    ' If no competition, no need to bid the opt.'
    ' If opponents have bid their optimum which '
    ' is higher than yours, sac if profitable.  '
    optGameLvlUs = 0 ' Max is a partscore unless overridden below. '
    If optSuitNumUs < 3 Then ' C/D '
        If optMaxLvlUs > 4 Then
            optGameLvlUs = optMaxLvlUs ' 5C/D for game. 6-7 = slam '
        End If
    ElseIf optSuitNumUs < 5 Then ' H/S '
        ' Setting optGameLvlUs to optMaxLvlUs was    '
        ' causing BB to jump to slam inappropriately '
        ' such as with a worthless doubleton. '
        'If optMaxLvlUs > 5 Then
        '    optGameLvlUs = optMaxLvlUs ' Slam '
        If optMaxLvlUs > 3 Then
            optGameLvlUs = 4 ' Game minimum level; Max could be 5H/S. '
        End If
    Else ' NT: '
        'If optMaxLvlUs > 5 Then
        '    optGameLvlUs = optMaxLvlUs ' Slam '
        'ElseIf optMaxLvlUs > 2 Then
            optGameLvlUs = 3 ' Game minnimum level; Max could be 4 or 5. '
        'End If
    End If
    
    optMaxUs = int_OptAvg
    optGameUs = optGameLvlUs * 10 + optSuitNumUs

    ' Opponents: '
  If int_Opt4OppAvg > 0 Then
    optSuitNumOpp = int_Opt4OppAvg Mod 10
    optSuitStrOpp = Mid$("CDHSN", optSuitNumOpp, 1)
    optMaxLvlOpp = int_Opt4OppAvg \ 10
    If optSuitNumOpp > optSuitNumUs Then oppOptSuitHigher = 1

    optGameLvlOpp = 0
    If optSuitNumOpp < 3 Then ' C/D '
        If optMaxLvlOpp > 5 Then
            optGameLvlOpp = optMaxLvlOpp ' Slam '
        ElseIf optMaxLvlOpp > 4 Then
            optGameLvlOpp = 5 ' Game '
        End If
    ElseIf optSuitNumOpp < 5 Then ' H/S '
        If optMaxLvlOpp > 5 Then
            optGameLvlOpp = optMaxLvlOpp ' Slam '
        ElseIf optMaxLvlOpp > 3 Then
            optGameLvlOpp = 4 ' Game '
        End If
    Else ' NT: '
        If optMaxLvlOpp > 5 Then
            optGameLvlOpp = optMaxLvlOpp ' Slam '
        ElseIf optMaxLvlOpp > 2 Then
            optGameLvlOpp = 3 ' Game '
        End If
    End If
  End If
  
    ' Find the high bids in Prior Calls: '
    j = 0
    For i = int_CallNum - 1 To 1 Step -1
        ' Working backwards from the last call, the '
        ' first bid we come across will be the high '
        ' bid. Once we find our highest bid and the '
        ' high optimum suit bid, and the opponents' '
        ' highest bid, we can exit this loop.       '
        j = j + 1 ' Keeps track of which pair's call we're looking at. '
        ' If j is odd, the call is the opponents'; otherwise, it's us. '
        ' 1=RHO, 2=pard, 3=LHO, 4=me, 5=RHO, 6=pard, 7=LHO, 8=me, etc. '
        If Len(str_Call(i)) > 1 Then
            ' A 1-character call has to be P, D, or R. '
            ' We are just checking actual 2-char calls.'
            If j Mod 2 = 0 Then
                ' if j Mod 2 = 0 then j is us. '
                ' High bid in any suit: '
                If hiBidLvlUs = 0 Then
                    hiBidLvlUs = Val(str_Call(i)) ' ex: 2C : hiBid level is 2. '
                    hiBidSuitUs = InStr("CDHSN", Right$(str_Call(i), 1))
                End If
                ' High bid in the Optimum suit: '
                If hiOptBidLvlUs = 0 And _
                   optSuitStrUs = Right$(str_Call(i), 1) _
                Then
                    hiOptBidLvlUs = Val(str_Call(i))
                End If

            ElseIf j Mod 2 = 1 Then
                If hiBidLvlOpp = 0 Then
                    hiBidLvlOpp = Val(str_Call(i))
                    hiBidSuitOpp = InStr("CDHSN", Right$(str_Call(i), 1))
                End If
                ' High bid in the Optimum suit: '
                If hiOptBidLvlOpp = 0 And _
                   optSuitStrOpp = Right$(str_Call(i), 1) _
                Then
                    hiOptBidLvlOpp = Val(str_Call(i))
                End If
            End If

            If hiBidLvlUs > 0 And hiBidLvlOpp > 0 And _
               hiOptBidLvlUs = 0 _
            Then
                ' We started with the last bid '
                ' so when we have high bids for'
                ' both sides, we can stop. '
                Exit For
            End If
        End If
    Next

    hiBidUs = hiBidLvlUs * 10 + hiBidSuitUs  ' 11-75 : 1C-7N '
    hiBidOpp = hiBidLvlOpp * 10 + hiBidSuitOpp

    If hiBidLvlUs > hiBidLvlOpp Then
        hiBidLvl = hiBidLvlUs
        hiBidSuit = hiBidSuitUs
    ElseIf hiBidLvlUs = hiBidLvlOpp And _
           hiBidSuitUs > hiBidSuitOpp _
    Then
        hiBidLvl = hiBidLvlUs
        hiBidSuit = hiBidSuitUs
    Else
        hiBidLvl = hiBidLvlOpp
        hiBidSuit = hiBidSuitOpp
    End If

'GoTo QuickExits:
''''''''''''''''''''''''''New way to calculate'''''''''''''''''''''''''''''''

'    Stop
    ' Pard bid opt game. Pass.
    
    ' RHO opened 1N. E has semi-balanced 9 HCP. Pass

    ' 1N-P-P. W is 5-5 and 10 HCP. S passed - weak. Balance.

    ' 1N-P-P-overcall. Pard passed 1N originally, but with four S and 5 D, Dbl.

    ' N's Dbl shows S & minor. E's four spades means W could be short. E has H:AQ76.
    ' Compete with 3H.

    ' If DDA shows that we can make 4S but we haven't bid spades yet (see #8087),
    ' a player with fewer than 5 should not just bid them on the 3 or 4 level.

    ' If DDA show that we can make 4 but pard bid 3 in a competetive auction
    ' and hasn't showwn extras, we should not free-bid 4 without extras.
    
    
    

    ''''''''''Quick exits: '''''''''
QuickExits:
    If redoBidPt = "Quick exit." Then Stop
    bidPt = "Quick exit."
    
    If int_BidNum > 2 Then
    ' This If-End If has no active code.'
    
        If hiBidLvlOpp * 10 + hiBidSuitOpp > hiBidLvlUs * 10 + hiBidSuitUs > 0 Then
            ' Opp's have outbid us. (Ex: 4S vs 4H) '
            If hiBidLvlUs >= optMaxLvlUs And optSuitNumUs >= hiBidSuitUs Then
                ' We have already bid our max. '
                'If hiBidLvlOpp <= optMaxLvlOpp and optMaxLvlOpp >
            End If
        End If
    
        ' This was not doing Doubles. '
        'If Not (optMaxLvlUs > 5 And hiBidLvlUs = 6) Then
        '    If hiBidLvlUs * 10 + optSuitNumUs >= int_OptAvg Then
        '        GoTo SaveBid ' Pass
        '    End If
        'End If
    End If

    ' Lead-directing doubles: '
    If int_BidNum > 2 Then
      If str_Bid(int_BidNum - 3) = "1N" And _
        Len(str_Bid(int_BidNum - 1)) = 2 _
      Then
        ' 1N-any-2C/D/H/S '
        If int_SuitLen(hiBidSuitOpp) > 4 Then
            If int_HCP(hiBidSuitOpp) > 2 Then
                b = "D"
                GoTo SaveBid
            End If
        End If
        ' This only covers 1N artificial responses.         '
        ' BB has no way to recognize other artificial bids. '
      End If
    End If
    
SlamBidding:
    
    If optMaxLvlUs > 5 And hiBidLvl < 5 Then
    'Stop ' Add code as needed. '
    ' A5-AKT54-A85-T85 ' Want to know H:Q and C:control. '
                ' We've bid game but opt is slam. '
                ' Ace asking may not be correct,  '
                ' but that's where I'm starting.  '
                If redoBidPt = "1b1" Then Stop
                bidPt = "1b1"
                If optSuitNumUs < 5 Then
                    b = "4N@" & optSuitNumUs & "0"
                    GoTo SaveBid
                End If
    End If
    
    ' King-asking response (after 4N@ ace asking): '
    b = ""
    'Stop ' If the below works, remove the stop. '
    k = InStr(allBids(1), "5N@") ' See if 5N@ has been bid. '
    If k > 0 Then
        If Mid$(allBids(1), k + 4, 1) = "1" Then
            ' Some play to bid the lowest K. '
            ' The 1 on 5N@#1 means lowest K. '
            ' The # means the trump suit num.'
            For s = 1 To 4
                If bln_Suit(s, King) And _
                   s <> int_TrumpSuit _
                Then
                    b = "6" & Mid$("CDHS", s, 1)
                    Exit For
                End If
            Next
        Else
            ' Others play to show the # of K's. '
            ' If it's 5N@#0, 0 means show total '
            ' number of K's excluding trump K.  '
            n = 0
            For s = 1 To 4
                If bln_Suit(s, King) And _
                   s <> int_TrumpSuit _
                Then
                    n = n + 1
                End If
            Next
            b = "6" & Mid$("CDHS", n + 1, 1) ' 0-3 means C-S '
            ' Example: 2 K's outside of trumps = 6H. '
            '          No K's = 6C. '
        End If
        If b <> "" Then GoTo SaveBid

    ' Ace-asking response: '
    ElseIf Left$(allBids(3) & "    ", 3) = "4N@" Then
        MsgBox "Because Ace-asking is conventional/artificial," & vbCrLf & _
               "subsequent bidding must be manually entered."
        GoTo SaveBid
        Stop ' MsgBox "Break - ComputeOptBid 1"
        n = 0
        For s = 1 To 4
            If bln_Suit(s, Ace) Then
                n = n + 1
            End If
            ' RKCB: '
            If bln_Suit(s, King) And _
               s = int_TrumpSuit _
            Then
                n = n + 1
            End If
        Next
        ' RKCB 1430: '
        If n = 1 Or n = 4 Then
            b = "5C"
        ElseIf n = 3 Or n = 0 Then
            b = "5D"
        ElseIf n = 2 Or n = 5 Then
            If bln_Suit(int_TrumpSuit, Queen) Then
                b = "5S" ' 2 or 5 Keys AND the trump Queen. '
            Else
                b = "5H" ' 2 or 5 keys W/O the trump Queen. '
            End If
            If redoBidPt = "RKCB 1430" Then Stop
            bidPt = "RKCB 1430"
        End If
        GoTo SaveBid
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'2.
    If redoBidPt = "00000" Then
        'Stop
    Else
        redoBidPt = "00000"
    End If

    If int_HCP(0) < 5 Then
        ' allBids(1-4) = String of all bids made by each hand, numbered CCW '
        ' where 1 is me. '
'        Stop ' Check code '
        If allBids(1) = "" And _
           pardsLastCall = "D" And _
           int_SuitLen(optSuitNumUs) > 5 _
        Then
            If optGameLvlUs > 0 Then
                b = "4" & optSuitStrUs  ' Law of TT '
            ElseIf hiBidLvlOpp > optMaxLvlUs Then
                b = "4" & optSuitStrUs
            Else
                b = optMaxLvlUs & optSuitStrUs
            End If
            GoTo SaveBid

        ElseIf allBids(1) = "" And _
           Not Right$(" " & allBids(3), 1) <> "D" And _
           Not (InStr(allBids(3), optSuitStrUs) > 0 And _
                int_SuitLen(optSuitNumUs) > 2) _
        Then
            b = "P"
            GoTo SaveBid
'3.
        Else
            b = "P"
            GoTo SaveBid
        End If
        GoTo SaveBid
    End If

' Penalty double: '

    i = int_CallNum
    If i > 4 Then i = i - 4
    j = int_CallNum
    If j > 2 Then j = j - 2

    If hiBidLvlOpp > optMaxLvlOpp And _
       hiBidLvlOpp > hiBidLvlUs And _
       hiBidLvlOpp > optMaxLvlUs And _
       str_Call(i) <> "D" And _
       str_Call(j) <> "D" _
    Then
        ' Opponents' opt is higher than ours '
        ' and they have overbid it, so double.
        b = "D"
        GoTo SaveBid
    End If

    If int_HCP(0) < 10 And _
       allBids(3) = "" And _
       hiBidLvlOpp > 2 _
    Then
        ' Example: RHO opened Unusual 3NT and '
        ' I have xx-KQxxxx-xxx-xx, just pass.  '
        b = "P"
        'Stop ' check the math
        GoTo SaveBid
    End If

    If hiBidUs > hiBidOpp And optGameLvlUs > 0 And hiBidLvl > 1 Then
        ' Opt is a part score and we've overbid them. '
        If hiBidUs >= optGameUs And optMaxLvlUs < 6 Then
            GoTo SaveBid
        End If
    End If

    If hiBidLvl = hiBidLvlUs And _
       hiBidSuit = optSuitNumUs And _
       hiBidSuit = hiBidSuitUs And _
       hiBidLvl >= optGameLvlUs And _
       optGameLvlUs > 0 And _
       optMaxLvlUs < 6 _
    Then
'4.
        ' We've already bid our max and slam is out. '
        b = "P"
        GoTo SaveBid
    End If

    If int_BidNum > 10 Then
        If str_Bid(int_BidNum - 2) = "D" And _
           str_Bid(int_BidNum - 1) = "P" _
        Then
            b = "P"
            GoTo SaveBid
        End If
    End If

    If str_Call(int_CallNum) <> "" Then
        MsgBox "At this date, current bidder" & Chr$(13) & _
               "should not have bid yet" & Chr$(13) & _
               "but in the future, after a" & Chr$(13) & _
               "bid is found, the program" & Chr$(13) & _
               "could come here for verification."
        If redoBidPt = "1z" Then Stop
        bidPt = "1z"
        Stop
        GoTo SaveBid
    End If

    j = 0
    For i = 1 To 4
        If int_SuitLen(i) < 2 Then
            unbal = True
            Exit For
        End If
        If int_SuitLen(i) = 2 Then
            j = j + 1
        End If
    Next
    If j < 1 Then ' 2+ doubletons '
        unbal = True
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                      Bidding calculations begin                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If redoBidPt = "RKCB 1430" Then Stop: GoTo SlamBidding
    bidBlok = "100"
WeBid1:
    If allBids(3) <> "" And _
       allBids(2) = "" And _
       allBids(4) = "" _
    Then
        ' Pard has bid and the opponents have not. '
        ' This excludes, P, D, and R. '
WeBid:
'5.
        bidBlok = "100"
        i = InStrRev(allBids(3), optSuitStrUs)
        'Stop
        If i > 0 Then
            ' Pard has bid our optimum suit (ex: H) at some point. '
            
            If str_Bid(int_BidNum - 1) = "D" Then ' Ex: 1S-P-2H-D. '
                If int_HCP(0) > 15 Then ' My total HCP > 15. '
                    If int_SuitLen(optSuitNumUs) > 3 Then ' Ex: Len(H) is 4+. '
                        If hiBidLvlUs + 2 <= optMaxLvlUs Then
                            ' Ex: hiBidLvlUs = 2 +2 <= 4H (Example). '
                            If redoBidPt = "1a1" Then Stop
                            bidPt = "1a1"
                            b = (hiBidLvlUs + 2) & optSuitStrUs
                            GoTo SaveBid
                        Else
                            Stop ' MsgBox "Calculation needed."
                        End If
                    End If
                Else
                    Stop ' MsgBox "Calculation needed."
                End If
            
            ElseIf hiBidLvlUs = optGameLvlUs And _
                   optMaxLvlUs > 5 And hiBidLvlUs < 5 _
            Then
                GoTo SlamBidding
            End If
            
'5a.
            j = Val(Mid$(allBids(2), i - 1))
            
            If optGameLvlUs = 0 And (Val(bidPt) = 2 Or bidPt = "") Then
                ' We can't make game (optGameLvlUs = 0)  '
                ' so any part score level will do for us.'

                If i >= Len(allBids(3)) Then
                    ' Pard bid the part score in our opt   '
                    ' on his last bid, so no need to go on.'
                    If redoBidPt = "1c1" Then Stop
                    bidPt = "1c1"  ' Pass '
                    b = "P"
                    GoTo SaveBid
                Else
                    ' Pard bid our opt suit earlier,  '
                    ' but I need to bid it at above   '
                    ' any other suits we bid since.   '
                    If hiBidSuit > optSuitNumUs Then
                        b = (hiBidLvl + 1) & optSuitStrUs
                    Else
                        b = hiBidLvl & optSuitStrUs
                    End If
                    If redoBidPt = "1d1" Then Stop
                    bidPt = "1d1"
                    GoTo SaveBid
                End If
'5b.
            ElseIf int_SuitLen(optSuitNumUs) > 2 Then ' We can make game+. '
                ' Pard bid below game in our opt suit. '
                ' Options are to invite game, bid game, or explore slam. '
                ' Other options are conventions which shold be hard coded.
                'Stop ' check the code '
                
                
                If optMaxLvlUs > 5 And _
                   InStr(allBids(3), optSuitStrUs) > 0 And _
                   InStr(allBids(1), optSuitStrUs) > 0 _
                Then
                    ' We may have slam.
                    If redoBidPt = "1e1" Then Stop
                    bidPt = "1e1"
                    Dim askOK As Integer
                    askOK = 0
                    For i = 1 To 4 ' Clubs to Spades '
                        If str_Suit(i) = "" Then Exit For
                        If InStr(str_Suit(i), "A") = 0 And InStr(str_Suit(i), "K ") = 0 Then
                            Exit For
                        End If
                    Next
                    
                    If askOK = 1 Then
                    
                        If optSuitNumUs < 5 And hiBidLvl < 5 Then
                            b = "4N@" & optSuitNumUs & "0"
                        ElseIf hiBidLvl < 4 Then
                            b = "4C"
                        ElseIf hiBidUs < int_OptAvg Then
                            b = "6" & optSuitStrUs
                        Else
                            b = "P"
                        End If
                        GoTo SaveBid
                    End If
                End If

                If optGameLvlUs - 1 > hiBidLvl And _
                   int_BidNum > 2 And _
                   int_HCP(0) < 12 _
                Then
                    ' Pard had bid 2N and optGameLvlUs = 3N '
                    ' so I change >= hiBidLvl to > hiBidLvl '
                    b = (optGameLvlUs - 1) & optSuitStrUs
                ElseIf optGameLvlUs > hiBidLvl Then
                    If optGameLvlUs > 5 Then
                        GoTo SlamBidding
                    Else
                        b = optGameLvlUs & optSuitStrUs
                    End If
                ElseIf optSuitNumUs > hiBidSuit Then
                    b = hiBidLvl & optSuitStrUs
                End If
                If b <> "" Then
                    If redoBidPt = "2b2b" Then Stop
                    bidPt = "2b2b"
                    GoTo SaveBid
                End If
            End If
'5c.
        ElseIf InStr(allBids(3), optSuitStrUs) = 0 And _
               InStr(allBids(1), optSuitStrUs) > 0 _
        Then
            ' Pard has bid, but has never bid our opt suit.  '
            ' I have bid the opt suit. I need 6+ to rebid it '
            ' unless the opt is NT. '
            If optSuitNumUs = 5 Then
                If hiBidSuit > optSuitNumUs Then
                    b = (hiBidLvl + 1) & optSuitStrUs
                ElseIf optSuitNumUs >= hiBidSuitUs Then
                    b = (hiBidLvl + 1) & optSuitStrUs
                Else
                    b = hiBidLvl & optSuitStrUs
                End If
                ' Could jump with a strong enough hand: '
                ' 16-17 jump to 3, 18+ jump to 4. '
                If redoBidPt = "2b2c" Then Stop
                bidPt = "2c"
                GoTo SaveBid
            ElseIf hiBidLvlUs < 3 Then
                For i = 4 To 1 Step -1
                    X = Mid$("CDHS", i, 1)
                    If int_SuitLen(i) > 3 And _
                       InStr(allBids(1), X) = 0 _
                    Then
                        ' Jump with 16+ HCP. '
                        If hiBidSuit > i Then
                            b = (hiBidLvl + 1) & X
                        Else
                            b = hiBidLvl & X
                        End If
                        If redoBidPt = "2d" Then Stop
                        bidPt = "2d"
                        GoTo SaveBid
                    End If
                Next
                
                If int_SuitLen(optSuitNumUs) > 5 Then
                    b = (hiBidLvl + 1) & optSuitStrUs
                    If redoBidPt = "2d2" Then Stop
                    bidPt = "2d2"
                    GoTo SaveBid
                End If
                
                If optSuitNumUs < hiBidSuitUs And hiBidLvlUs > optGameLvlUs Then
                    b = "P"
                    GoTo SaveBid
                End If

                If Not unbal Then
                    If hiBidLvl * 10 + 5 < int_OptAvg Then
                        If redoBidPt = "2e" Then Stop
                        bidPt = "2e"
                        If hiBidSuit = 5 Then
                            b = (hiBidLvl + 1) & "N"
                        Else
                            b = hiBidLvl & "N"
                        End If
                        GoTo SaveBid
                    End If
                End If
'5d.
                If RHOsLastCall <> "D" Then
                    ' I don't have a bid, though if I've very '
                    ' strong I should make up one; otherwise, '
                    ' I can pass RHO's double around to pard. '
                    b = "P"
                    GoTo SaveBid
                Else
                    ' Can't let it pass out below opt. '
                    If hiBidSuit > optSuitNumUs Then
                        b = (hiBidLvl + 1) & optSuitStrUs
                    Else
                        b = hiBidLvl & optSuitStrUs
                    End If
                    If redoBidPt = "2f" Then Stop
                    bidPt = "2f"
                    GoTo SaveBid
                End If

            End If
'5e.
        ElseIf optSuitNumUs = 5 And hiBidLvl < 4 And _
               (1 = 1 Or int_SuitLen(hiBidSuitUs) = 2) _
        Then ' Our opt = NT. '
            If hiOptBidLvlUs < optMaxLvlUs Then
                If hiBidLvl > hiOptBidLvlUs And _
                   hiOptBidLvlUs > 0 _
                Then
                    b = (hiBidLvl + 1) & optSuitStrUs
                
                ElseIf Len(allBids(3)) \ 2 > Len(allBids(1)) \ 2 And _
                       int_HCP(0) > 12 And _
                       optGameLvlUs > 2 _
                Then
                    ' Pard has made more bids so must be stronger   '
                    ' and this hand has enough for game, so bid it. '
                    b = optGameLvlUs & optSuitStrUs
                Else
                    b = hiBidLvl & optSuitStrUs
                End If
            Else
                b = (hiOptBidLvlUs + 1) & optSuitStrUs
            End If
            If redoBidPt = "2f1" Then Stop
            bidPt = "2f1"
            GoTo SaveBid
'5f.
        ElseIf optSuitLen4Me > 3 And optSuitLen4Me < 99 Then
            ' Neither pard nor I have bid our opt.      '
            ' I have 4+ of the suit, so I should bid it.'
            If hiBidSuitUs > optSuitNumUs Then
                b = (hiBidLvlUs + 1) & optSuitStrUs
            Else
                b = (hiBidLvlUs) & optSuitStrUs
            End If

            If Val(b) > optMaxLvlUs Then
                n = MsgBox("Bid=" & b & _
                           " but Max level is " & optMaxLvlUs & _
                           ". Save the bid anyway?", vbYesNo)
                If n = vbNo Then
                    b = ""
                    ' Opps haven't bid, so we just bid '
                    ' past the max ourselves. '
                End If
            End If
            If redoBidPt = "2f3" Then Stop
            bidPt = "2f3"
            GoTo SaveBid

            ' It is barely possible the he has 5 of our '
            ' opt and I only have 3, in which case it   '
            ' would make no sense for me to open the    '
            ' suit and I must find another bid. '
            Stop ' MsgBox "Stop ComputeOptBid 2" ' 2d, e, f
        Else
            ' We haven't bid our opt. I '
            ' need to bid another suit. '
            n = 0
            j = 0
            k = 0
            For i = 1 To 4
                'If InStr(allBids(3), Mid$("CDHS", i, 1)) > 0 Then
                    If i < 3 Then
                        If int_SuitLen(i) > 3 Then
                            n = i
                        End If
                    Else
                        If int_SuitLen(i) > 2 Then
                            n = i
                        End If
                    End If
                    If int_SuitLen(n) >= int_SuitLen(j) And n > j Then
                        j = n
                    End If
                'End If
            Next
            n = 0 ' A variable like n should not '
                  ' retain value beyond a routine.
            If j > 0 Then
                If hiBidSuit >= j Then
                    b = (hiBidLvl + 1) & Mid$("CDHS", j, 1)
                Else
                    b = hiBidLvl & Mid$("CDHS", j, 1)
                End If
                If redoBidPt = "3" Then Stop
                bidPt = "3"
                GoTo SaveBid
            End If
        End If
'5g.
    ElseIf int_Opt4OppAvg > 0 Then
        ' Competitive bidding: '

        If hiBidLvl = hiBidLvlUs And _
           hiBidLvl > 0 And _
           hiBidSuit = hiBidSuitUs _
        Then
            ' We were the last to bid. '
            If (hiBidLvlUs * 10 + hiBidSuit) < int_OptAvg Then
                ' Our last bid was below the optimum. Keep going. '
                GoTo WeBid
            Else
                GoTo SaveBid ' Pass '
            End If

        Else ' Opps were last to bid (excl. P,D,R). '
            If hiBidSuit > optSuitNumUs Then
                n = hiBidLvl + 1
            Else
                n = hiBidLvl
            End If

            If InStr(allBids(1), optSuitStrUs) > 0 Then
                ' I bid the opt, I can either rebid it '
                ' to make or check into making a sac.  '
                If n <= optMaxLvlUs Then
                    b = n & optSuitStrUs
                    bidPt = ""
                    GoTo SaveBid
                ElseIf (pardsLastCall <> "D" Or _
                    Len(str_Call(int_BidNum - 1)) = 2) And int_Opt4OppAvg > 0 _
                Then
                    GoTo Check4Sac
                Else
                    GoTo SaveBid
                End If

            ElseIf InStr(allBids(3), optSuitStrUs) > 0 And _
                   int_SuitLen(optSuitNumUs) > 2 _
            Then
                ' Pard has bid the opt suit. '
                ' If my opt is 4C and pard's is 3N '
                ' then he actually has bid MY opt  '
                ' suit, not OUR's (which is N) and '
                ' if he has not bid N yet, then if '
                ' 1N would be the normal bid, I    '
                ' still have to make it, just like '
                ' if pard opens 1N and our opt is  '
                ' 2C by him, I can't just not make '
                ' a normal 2C Stayman bid. And once'
                ' I made the 1N bid, his opt will  '
                ' then be the same as mine - 4C    '
                ' unless he has some other opt bid.'
                If hiBidSuit >= optSuitNumUs Then
                    n = hiBidLvl + 1
                Else
                    n = hiBidLvl
                End If

                If n > optMaxLvlUs Then
                    GoTo Check4Sac
                End If

                If int_HCP(0) > 17 And _
                   n < optGameLvlUs _
                Then
                    n = optGameLvlUs
                End If

                b = n & optSuitStrUs
                If redoBidPt = "3a" Then Stop
                bidPt = "3a"
                GoTo SaveBid

            ElseIf hiBidLvlUs = 0 And _
                   hiBidLvlOpp > 0 And _
                   pardsLastCall <> "D" _
            Then
                ' We haven't bid and they have.  '
                ' BidBase should really have the '
                ' overcalls & TOX in it, so I'm  '
                ' passing for now and will add   '
                ' entries when asked about pass  '
                ' if needed. '
                b = "P"
                GoTo SaveBid

            ElseIf allBids(1) = "" And _
                   int_HCP(0) < 6 And _
                   pardsLastCall <> "D" _
            Then
                GoTo SaveBid

            Else ' Neither pard nor I have bid our opt. '
                ' I should bid it with 4+ if I've bid;  '
                ' otherwise, I should rebid a 6+ card   '
                ' suit or else a second 4+ card suit if '
                ' none of these take us past our opt.   '
                If optSuitNumUs < 5 Then
                    i = int_SuitLen(optSuitNumUs)
                Else
                    i = 0
                End If

                If i > 3 Then
                    If hiBidSuit > optSuitNumUs Then
                        n = hiBidLvl + 1
                    Else
                        n = hiBidLvl
                    End If

                    If n < optMaxLvlUs + 1 And n > 0 And _
                       int_SuitLen(optSuitNumUs) > 5 _
                    Then ' Sac 1 if needed. '
                        b = n & optSuitStrUs
                        bidPt = "3a1"
                        GoTo SaveBid
                    Else
                        GoTo Check4Sac
                    End If
                Else ' Check for 6+ card suit to rebid. '
                    If hiBidSuitUs < 5 And _
                       hiBidSuitUs > 0 _
                    Then
                        k = int_SuitLen(hiBidSuitUs)
                    Else
                        k = 0
                    End If
                    If k > 5 Then
                        If hiBidSuit > optSuitNumUs Then
                            n = hiBidLvl + 1
                        Else
                            n = hiBidLvl
                        End If
                        If n < optMaxLvlUs + 1 Then ' Sac 1 if needed. '
                            b = (n + 1) & Mid$("CDHS", hiBidSuitUs, 1)
                            If redoBidPt = "D1" Then Stop
                            bidPt = "D1"
                            GoTo SaveBid
                        Else
                            'Stop ' MsgBox "Stop ComputeOptBid 3"
                        End If
                    ' Else a 2nd 4+ card suit. '
                    ' Else bid NT with a stopper else Dbl? '
                    ElseIf int_HCP(0) > 17 Then
                        If redoBidPt = "D2" Then Stop
                        bidPt = "D2"
                        b = "D"
                        GoTo SaveBid
                    ElseIf int_BidNum < 2 Then
                        b = "P"
                        bidPt = "Can't open."
                        GoTo SaveBid
                    Else
                        GoTo Check4Sac
                    End If

                    Stop
                End If
            End If
        End If
'5h.
        X = allBids(2) & allBids(3) & allBids(4)
        j = 0
        For i = 4 To 1 Step -1
            ' Assume if we could have raised pard  '
            ' we would have done it above, so just '
            ' look for a long suit to bid. '
            If int_SuitLen(i) > 3 Then
                If int_SuitLen(1) > j Then
                    b = Mid$("CDHS", i, 1)
                    If InStr(X, b) = 0 Then
                        ' Suit has not been bid by anyone. '
                        j = i
                    End If
                End If
            End If
        Next
        b = (hiBidLvl + 1) & Mid$("CDHS", j, 1)
        If redoBidPt = "4a" Then Stop
        bidPt = "4a"
        GoTo SaveBid
    End If

'MsgBox "Break - ComputeOptBid 4" ' Add entry above. '


If int_OptAvg = hiBidUs And bidPt <> "" Then
    'Stop
    'Exit Sub
End If

' The following was the original mish-mash of code before organizing it. '
    If int_BidNum = 2 And _
       hiBidLvlOpp > 0 And _
       hiBidSuitOpp > optSuitNumUs _
    Then
        ' 2nd player to bid must adhere to '
        ' normal overcall requirements if  '
        ' he must overcall 1N or 2x. '
        If (int_HCP(optSuitNumUs) > 2 And _
              int_SuitLen(optSuitNumUs) > 5) Or _
           (int_HCP(optSuitNumUs) > 4 And _
              int_SuitLen(optSuitNumUs) = 5) _
        Then
            b = (hiBidLvlOpp + 1) & optSuitStrUs
            GoTo SaveBid
        ElseIf int_HCP(0) < 13 Then ' My HCP < 13. '
            If hiBidSuitOpp > 2 And _
               int_SuitLen(3) < 5 And _
               int_SuitLen(4) < 5 _
            Then
                ' With < 13 HCP and no 5+ card major '
                ' and RHO bid 1C/D, I can't bid. '
                GoTo SaveBid
            End If
        Else
            GoTo SaveBid
        End If
    End If


    bidPt = 11
    If redoBidPt = "1b1" Then Stop
    
    If hiOptBidLvlOpp = 0 And _
       RHOsLastCall <> "D" _
    Then
        ' Opponents have not bid. '
        'If int_BidNum < 9 Then
        '    i = MsgBox("The first 2 rounds of non-competitive bids" & vbCrLf & _
        '               "should be manually entered. Compute bid anyway?", vbYesNo)
        '    If i = vbNo Then GoTo SaveBid
        'End If
    End If

    If optSuitNumUs < 5 Then
        i = Len(str_Suit(optSuitNumUs))
    Else
        i = 0
    End If

    If int_OptAvg > (hiBidLvlOpp * 10 + hiBidSuitOpp) And _
       int_CallNum > 3 And _
       (i > 4 Or _
         i = 0 Or _
         hiOptBidLvlUs > 0) _
    Then
        If optSuitStrUs = "N" Then
            If hiBidSuitOpp = 0 Or _
               hiBidSuitOpp = 5 _
            Then
                i = 0
            Else
                i = int_HCP(hiBidSuitOpp)
            End If
            
            If i > 2 Then
                ' Have opp's suit stopped, so bid NT: '
                If int_BidNum Mod 4 = 3 And _
                   int_HCP(0) > 13 And _
                   optGameLvlUs > hiBidLvl _
                Then
                    ' I'm in 3rd chair with >13 HCP; pard opened. '
                    b = optGameLvlUs & "N"
                ElseIf hiBidSuit = 5 Then
                    b = hiBidLvl + 1 & "N"
                ElseIf (int_HCP(0) > 24 Or _
                          sng_NumWinners(0) > 8) _
                       And hiBidLvl < 4 _
                Then
                    b = "3N"
                ElseIf int_HCP(0) > 20 And _
                       hiBidLvl < 3 _
                Then
                    b = "2N"
                Else
                    b = hiBidLvl & "N"
                End If
            
            ElseIf optGameLvlUs - hiBidLvl > 0 Then
                b = hiBidLvl + 1 & optSuitStrUs
                
            ElseIf optMaxLvlUs > 5 And hiBidUs >= optGameUs Then
                ' Slam is possible but only game has been bid. '
                If redoBidPt = "12a" Then Stop
                bidPt = "12a"
                b = optMaxLvlUs & optSuitStrUs
                GoTo SaveBid
            
            ElseIf optGameLvlUs <= hiOptBidLvlUs Then
                GoTo SaveBid
            ElseIf hiBidLvlOpp > hiBidLvl Then
                b = hiBidLvlOpp & Mid$("CDHSN", hiBidSuitOpp + 1, 1)
            Else
                ' Deal 8081-S was a control bid. '
            End If
        ElseIf optGameLvlUs - hiOptBidLvlUs = 1 Or _
               (optGameLvlUs - hiOptBidLvlUs = 2 And _
                 int_HCP(0) > 15) _
        Then
            b = optGameLvlUs & optSuitStrUs
        ElseIf optGameLvlUs - hiBidLvl > 0 And _
               InStr(pardsLastCall, optSuitStrUs) > 0 _
        Then
            i = hiBidLvl
            If int_HCP(0) > 15 And _
               int_HCP(0) < 18 _
            Then ' My HCP = 16-17. '
                i = hiBidLvl + 1
            
            ElseIf int_HCP(0) > 17 And _
                   int_HCP(0) < 22 _
            Then ' My HCP = 18-21. '
                i = hiBidLvl + 2
            
            Else
                i = hiBidLvl + 3 ' May be above optimum. '
                ' Slam may be on. '
            End If
            b = i & optSuitStrUs
        
        ElseIf optGameLvlUs - hiBidLvlOpp < 3 Then
            'If optGameLvlUs = hiBidLvl And optSuitNumUs = hiBidSuit Then GoTo SaveBid
            nextHiBidLvlUs = hiBidLvl
            If optSuitNumUs <= hiBidSuit Then
                nextHiBidLvlUs = nextHiBidLvlUs + 1
            End If
            b = nextHiBidLvlUs & optSuitStrUs

        ElseIf hiBidLvlUs = hiBidLvl Then
            If hiBidSuit = optSuitNumUs Then
                b = hiBidLvlUs + 1 & optSuitStrUs
            ElseIf hiBidSuitOpp = hiBidSuit Then
                If optSuitNumUs > hiBidSuit Then
                    b = hiBidLvlUs & optSuitStrUs
                Else
                    b = hiBidLvlUs + 1 & optSuitStrUs
                End If
            End If
        ElseIf optMaxLvlUs - hiBidLvlUs > 2 Then
            If int_SuitLen(optSuitNumUs) < 4 Then
                b = "D"
            Else
                b = (hiBidLvl + oppOptSuitHigher) & optSuitStrUs
            End If
        Else
            b = optMaxLvlUs & optSuitStrUs
        End If
        If b <> "" Then GoTo SaveBid
    End If

    If optGameLvlUs = 0 Then ' optimum is part score '
        j = hiOptBidLvlUs * 10
    Else
        j = optGameLvlUs * 10
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' We check for sac below:
'6.
    If hiBidLvlUs * 10 + hiBidSuitUs > _
       hiBidLvlOpp * 10 + hiBidSuitOpp _
    Then
        hiBidLvl = hiBidLvlUs
        hiBidSuit = hiBidSuitUs
    ElseIf hiBidLvlOpp > hiBidLvlUs Then
        hiBidLvl = hiBidLvlOpp
        hiBidSuit = hiBidSuitOpp
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'7.
    ' See if our optimum suit has been bid yet: '
    i = Len(str_PriorCalls) + 1
    If i > 1 Then
        Do
            i = InStrRev(str_PriorCalls, optSuitStrUs, i - 1)
            If i = 0 Then Exit Do
        Loop While Val(Mid$(str_PriorCalls, i - 1, 1)) = 0
        ' Ex: 1C-P-1D-P and Opt suit=D, then i=7, but    '
        '     1C-1D-D-P and Opt suit=D, then i=5, not 7. '
    End If

    Do While i > 1
        hiOptBidLvlUs = Val(Mid$(str_PriorCalls, i - 1, 1))
        ' If 1C-P-1D-P and i=7 then hiOptBid=i-1="1".'
        If hiOptBidLvlUs > 0 Then Exit Do
        ' If optimum suit is D then D without '
        ' a number before it is Dbl, not Dia. '
        ' Look for another D: '
        i = InStrRev(str_PriorCalls, "D", i - 1)
    Loop

    If redoBidPt = "12b" Then Stop
    bidPt = "12b"

    ' See if next lowest Opt bid would be over our Max '
    ' & if so, see if we should sac or double or pass: '
Check4Sac:
    If redoBidPt = "Sac" Then
        bidPt = ""
        Stop
    End If
    
    diff = hiBidLvlOpp + oppOptSuitHigher - optMaxLvlUs
    ' diff is the number of tricks we would go set if we sac. '
    ' Ex: Their Max = 4S.  Our Max = 4H.                      '
    '     opp bid 4S. Higher suit=1. We would have to bid 5H. '
    'If optSuitNumUs < 5 Then
    '    j = int_SuitLen(optSuitNumUs)
    'Else
    '    j = 5
    'End If
    'If diff = 0 Or (diff > 0 And j < 4) Then
    '    GoTo SaveBid
    'Else
    
    If int_Opt4OppAvg > 0 Then
        If hiBidLvlUs > 0 And (j > 4 Or _
           hiOptBidLvlUs > 0) _
        Then
            ' We can't bid our opt suit without going over Max. '
            If hiBidLvlOpp * 10 + hiBidSuitOpp > _
                 int_Opt4OppAvg _
            Then
                ' PROBLEM: Deal #130023 W opened 1D with D:AKQ952 C:A76 '
                ' and E raised with D:84 C:QJT95 and eventually bid 5D. '
                ' Both 5C and 5D make but Opt was given as 5C. Since EW '
                ' bid 5D which is > 5C, S doubled 5D. '
                If RHOsLastCall <> "P" And _
                   pardsLastCall <> "D" And _
                   RHOsLastCall <> "D" _
                Then
                 ' Opp have overbid, so... '
                    b = "D"
                    GoTo SaveBid
                End If
            End If
        End If

        ' They didn't overbid. See if we should sac: '
        If str_Vul = "F" Then ' Favorable '
            If optGameLvlOpp = 0 Then
                ' Opp bid partscore: 70-150. Our -1=100. '
                If diff > 1 Then GoTo SaveBid
                ' When we go to SaveBid with b="", we pass. '
            
            ElseIf optGameLvlOpp < 6 Then ' Opp bid Game: 600-620. Our -3=500. '
                If diff > 3 Then GoTo SaveBid
            
            ElseIf optGameLvlOpp = 6 Then
                If optSuitNumOpp > 2 Then ' Opp bid 6C/D: 1370. Our -5=1100. '
                    If diff > 5 Then GoTo SaveBid
                Else                      ' Opp bid 6H/S/N: 1430-1440. '
                    If diff > 6 Then GoTo SaveBid ' Our -6 = 1400 '
                End If
            
            ElseIf optGameLvlOpp = 7 Then ' Opp bid 7any: 2140-2210. Our -8=2000. '
                If diff > 8 Then GoTo SaveBid
            End If

        ElseIf str_Vul = "U" Then ' Unfavorable '
            If optGameLvlOpp = 0 Then ' Opp bid partscore: 70-150. Our -1=200. '
                If diff > 0 Then GoTo SaveBid
            
            ElseIf optGameLvlOpp < 6 Then ' Opp bid Game: 400-420. Our -1=200. '
                If diff > 3 Then GoTo SaveBid
            
            ElseIf optGameLvlOpp = 6 Then ' Opp bid 6any: 920-990. Our -3=800. '
                If diff > 3 Then GoTo SaveBid
            
            ElseIf optGameLvlOpp = 7 Then ' Opp bid 7any: 1440-1510. Our -4=1100. '
                If diff > 4 Then GoTo SaveBid
            End If

        ElseIf str_Vul = "B" Or _
               str_Vul = "N " Or str_Vul = "" _
        Then ' Both/Neither vul. '
            If optGameLvlOpp = 0 Then ' Opp bid partscore: 70-150. Our -1=100. '
                If diff > 1 Then GoTo SaveBid
            
            ElseIf optMaxLvlOpp < 6 Then ' Both NV Game: 400-420. Our -2=300. '
                If diff > 2 Then GoTo SaveBid ' Both vul Game: 600-620, Our -2=500.'
            
            ElseIf optGameLvlOpp = 6 Then ' Both NV 6any: 920-990. Our -4=800. '
                If optSuitNumOpp > 2 Then ' Opp bid 6C/D: 1370. Our -4=1100. '
                    If diff > 4 Then GoTo SaveBid
                Else                      ' Opp bid 6H/S/N: 1430-1440. '
                    If diff > 5 Then GoTo SaveBid ' Our -5 = 1400 '
                End If
            
            ElseIf optGameLvlOpp = 7 Then ' Both NV 7any: 1440-1510. Our -6=1400. '
                If diff > 7 Then GoTo SaveBid ' Both vul 7any: 2140-2210. Our -7=2000. '
            End If
        Else
            ' V = vul. bidder, X = non-vul. bidder. Opponents can be either, '
            ' therefore we can't calculate overs/unders. But 1 over should be'
            ' safe if I have a long opt suit or if pard bid it. '
            If redoBidPt = "Sac" Then Stop
            bidPt = "Sac"
            If diff = 1 And _
               (int_SuitLen(optSuitNumUs) > 5 Or _
                 InStr(allBids(3), optSuitStrUs) > 0) _
            Then
                b = (optMaxLvlUs + 1) & optSuitStrUs
            Else
                b = "P"
            End If
            GoTo SaveBid
        End If

        ' Per the tests above, we can afford to bid higher. '

        If int_BidNum > 1 And _
           hiBidSuitOpp <> optSuitNumUs _
        Then
            If redoBidPt = "13" Then Stop
            bidPt = "13"

            If hiBidLvlOpp > optMaxLvlOpp And _
               hiBidLvlUs >= optMaxLvlUs _
            Then
                ' They overbid and we've bid as high as we can. '
                If pardsLastCall <> "D" And _
                   RHOsLastCall <> "P" _
                Then
                    If hiBidLvl & hiBidSuit = _
                       hiBidLvlOpp & hiBidSuitOpp _
                    Then
                        ' Opponents were last to bid. '
                        b = "D"
                    End If
                End If
                GoTo SaveBid
            End If

            If hiBidLvl + _
               oppOptSuitHigher - _
               optMaxLvlUs > 2 _
            Then
                ' Our sac would be too high. '
                GoTo SaveBid

            ElseIf optSuitLen4Me > 4 Or _
                   optSuitStrUs = "N" Or _
                   int_BidNum > 2 Or _
                   pardsLastCall = "D" _
            Then
                ' Above: Was overcalling with S:97542. '
                ' Below: When opp's high bid wasn't in '
                ' opt suit, we still have to bid higher'
                ' if we can afford to. '
                If hiBidSuitOpp >= optSuitNumUs Then
                    oppOptSuitHigher = 1
                Else
                    oppOptSuitHigher = 0
                End If
                
                If optSuitNumUs < 5 And _
                   hiBidLvl + oppOptSuitHigher < optMaxLvlUs _
                Then
                    If int_SuitLen(optSuitNumUs) = 4 And _
                       pardsLastCall = "D" _
                    Then
                        b = (hiBidLvl + oppOptSuitHigher) & optSuitStrUs
                    
                    ElseIf int_SuitLen(optSuitNumUs) > 4 Then
                        If hiBidSuitUs > optSuitNumUs Then
                            b = (hiBidLvl + oppOptSuitHigher + 1) & optSuitStrUs
                            ' If this is too high, make oppOptSuitHigher = 0. '
                        Else
                            b = (hiBidLvl + oppOptSuitHigher) & optSuitStrUs
                        End If
                    End If
                End If
            End If
            If b <> "" Then GoTo SaveBid
        End If

        ' Start of bidPt 20: '

        If redoBidPt = "20" Then Stop ' Edit this section. '
        bidPt = "20"

        If hiOptBidLvlUs < optSuitNumUs And _
           int_BidNum > 1 _
        Then
            If hiBidSuitOpp = optSuitNumUs Then
                ' An opponent has bid our suit. '
                ' If we already bid our suit, then '
                ' he is cue bidding... '

                MsgBox "Stop - Add cue bidding code."
            End If
            
            If hiBidOpp - hiBidLvlUs > 1 Then
                GoTo Ck4Double
            End If

            If Len(RHOsLastCall) > 1 Then
                i = (hiBidLvlOpp - (optSuitNumUs < hiBidSuitOpp))
                If i - optMaxLvlUs > 1 And _
                   optGameLvlOpp = 0 _
                Then
                    ' Can't go down 2 doubled when opps only have part score. '
                    GoTo SaveBid
                End If
                
                If optSuitNumUs < 3 And _
                   int_SuitLen(optSuitNumUs) > 3 _
                Then
                    ' Ex: 1D-1H-?? We need >3 Diamonds to raise. '
                    b = (hiBidLvlOpp - _
                        (optSuitNumUs < hiBidSuitOpp)) & _
                        optSuitStrUs
                
                ElseIf optSuitNumUs > 2 And _
                       int_SuitLen(optSuitNumUs) > 2 And _
                       (1 = 1 Or pardsLastCall <> "P") _
                Then
                    If pardsLastCall = "P" Then
                        If Right$(RHOsLastCall, 1) = "S" Then
                            
                        End If
                    End If
                
                    If hiBidSuitUs > hiBidSuitOpp Then
                        ' Ex: 1H-2C-?? We can raise a higher suit with the same bid level. '
                        b = hiBidLvlOpp & optSuitStrUs
                    Else
                        b = hiBidLvlOpp + 1 & optSuitStrUs
                    End If
                Else
                    GoTo SaveBid
                End If
            ElseIf hiBidLvlOpp = hiBidLvl And _
                   optMaxLvlUs > hiBidLvl And hiBidLvl > 0 _
            Then
                b = (hiBidLvlOpp + oppOptSuitHigher) & optSuitStrUs
            ElseIf hiBidSuitUs <= optSuitNumUs Then
                b = hiBidLvl & optSuitStrUs
            Else
                b = (hiBidLvl + 1) & optSuitStrUs
            End If
            If b = "" Then GoTo Ck4Double
            GoTo SaveBid
        End If
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Calculate an opening bid: '
    ' (Should rarely be needed because the calling program should handle it.)'
    If int_BidNum = 1 Or _
       int_CallNum = 1 _
    Then
        b = str_Call(int_CallNum)
        If b = "P" Or _
           b = "" _
        Then
            b = ""
            If int_HCP(0) > 11 Then
                ' Hand has 12+ HCP. Must bid in 1st chair. '
                If int_SuitLen(4) > 4 Then
                    b = "1S"
                ElseIf int_SuitLen(3) > 4 Then
                    b = "1H"
                ElseIf int_SuitLen(2) > 3 Then
                    b = "1D"
                ElseIf int_SuitLen(1) > 3 Then
                    b = "1C"
                ElseIf int_HCP(0) > 14 And _
                       int_HCP(0) < 18 _
                Then ' My HCP = 15-17. '
                    b = "1N"
                Else ' Have to bid something with 12+ HCP. '
                    b = "1C"
                End If
            ' The following bids are only if '
            ' this hand has <12 HCP, so these'
            ' bids are preemptive. '
            ElseIf int_HCP(0) > 4 And _
                   optSuitNumUs < 5 _
            Then
                If int_SuitLen(optSuitNumUs) = 6 And _
                   optSuitNumUs > 1 And int_HCP(optSuitNumUs) > 4 _
                Then
                    b = "2" & optSuitStrUs ' Weak 2 '
                
                ElseIf int_SuitLen(optSuitNumUs) = 7 Then
                    b = "3" & optSuitStrUs
                
                ElseIf int_SuitLen(optSuitNumUs) = 8 Then
                    b = "4" & optSuitStrUs
                End If
            End If
        End If

        If redoBidPt = "30" Then Stop
        bidPt = "30"

        GoTo SaveBid

    ElseIf int_BidNum = 2 Then
        'MsgBox ("2nd chair bid calcs not done yet.")
        'GoTo SaveBid
    End If

    ''''''''''''''''''''''''''''''''''''''
    ' From here down, bid number is > 2: '
    ''''''''''''''''''''''''''''''''''''''
    ' If pard's last bid len > 1 then '
    ' his bid must be hiBidLvlUs.     '
    ' Ditto for RHO and hiBidLvlOpp.  '
Ck4Double:
    
    If hiBidOpp > int_Opt4OppAvg Then
        ' Opps have overbid. '
        If hiBidLvlUs + 1 + optSuitNumUs <= int_OptAvg Then
            ' We have already bid at or over our max. '
            b = "D"
            GoTo SaveBid
        End If
    End If

    If Len(pardsLastCall) > 1 And _
       hiBidLvlUs * 10 + hiBidSuitUs >= int_OptAvg _
    Then
        ' Pard bid optGameLvlUs or more. '
        ' See if RHO overbid him:   '
        If Len(RHOsLastCall) = 1 Then
            ' 4a. RHO passed or doubled so we can pass. '
            GoTo SaveBid
        Else ' RHO bid. '
            ' 4b. See if RHO overbid his Optimum.  '
            '     If so, and we've already bid our '
            '     Optimum, we should double. '
            If hiBidLvlOpp * 10 + hiBidSuitOpp > _
               int_Opt4OppAvg _
            Then
                b = "D"

                If redoBidPt = "40" Then Stop
                bidPt = "40"

                GoTo SaveBid
            Else
                ' If we can overbid RHO in our optimum '
                ' by 1 then we should do it. '
                If optMaxLvlUs > hiBidLvlOpp Then
                    If optSuitNumUs > hiBidLvlOpp Then
                        ' Bidder has higher suit: '
                        b = hiBidLvlOpp & optSuitStrUs
                    Else ' Opps have bid a higher suit. '
                        b = hiBidLvlOpp & _
                            oppOptSuitHigher & _
                            optSuitStrUs
                    End If

                    If redoBidPt = "50" Then Stop
                    bidPt = "50"

                    GoTo SaveBid
                End If
            End If
        End If

    Else ' We haven't bid int_OptAvg/game/slam yet. '
        n = 0
        If Len(pardsLastCall) > 1 And _
           Len(RHOsLastCall) = 1 _
        Then
            ' Pard bid and RHO passed or doubled, so'
            ' I should bid the optimum suit (again).'
            ' If optGameLvlUs = 0 then int_OptAvg is a '
            ' part score so we don't need to bid    '
            ' unless passing would be unrealistic.  '
            If optGameLvlUs = 0 Then
                If int_HCP(0) > 17 Then
                    If hiBidLvlUs >= optSuitNumUs Then
                        b = (hiOptBidLvlUs + 1) & optSuitStrUs
                    
                    ElseIf hiOptBidLvlUs > 0 Then
                        b = (hiBidLvlUs + 1) & optSuitStrUs
                    Else
                        b = hiBidLvlUs & optSuitStrUs
                    End If
                End If

                If redoBidPt = "60" Then Stop
                bidPt = "60"

                GoTo SaveBid
            ElseIf hiOptBidLvlUs < optGameLvlUs Then
                ' We haven't bid at optimum level yet. '
                ' Also note optGameLvlUs is >0 here.   '
                If optGameLvlUs - hiOptBidLvlUs < 4 Then
                    ' We're not going to jump to slam nor '
                    ' from 1 level to game without HCP.   '
                    If int_HCP(0) > 11 Then
                        b = optGameLvlUs & optSuitStrUs
                    
                    ElseIf hiOptBidLvlUs >= hiBidLvlUs Then
                        If optGameLvlUs > hiBidSuitUs Then
                            b = hiOptBidLvlUs & optSuitStrUs
                        Else
                            b = (hiOptBidLvlUs + 1) & optSuitStrUs
                        End If
                    
                    ElseIf optMaxLvlUs > hiOptBidLvlUs Then
                        b = optGameLvlUs & optSuitStrUs
                    
                    ElseIf optSuitStrUs = "N" And _
                           hiBidSuitOpp > 0 And _
                           int_HCP(hiBidSuitOpp) < 3 _
                    Then
                        ' Bids 3N without a stopper in opp's suit. '
                        ' Find a 4-card unbid suit. '
                    ElseIf int_CallNum > 4 Then
                        ' See about rebidding my suit:  Ex: KQJ864 & Opt = 5D (52) '
                        X = str_Call(int_CallNum - 4) ' Ex: 1D & pard responded 1S '
                        If Len(X) > 1 Then            '     Want to rebid 3D w/10+ '
                            i = InStr("CDHS", Right$(X, 1))
                            
                            If int_SuitLen(i) > 5 Then
                                If int_HCP(0) > 9 Then
                                    b = (hiBidLvl + 2) & Right$(X, 1)
                                Else
                                    b = (hiBidLvl + 1) & Right$(X, 1)
                                End If
                            End If
                        End If
                    Else
                        b = (hiBidLvlUs + 1) & optSuitStrUs
                    End If
                
                ElseIf optGameLvlUs > 5 And _
                       hiBidLvlUs < 4 _
                Then
                    ' Game optimum is slam and bidding is below the 4 level. '
                    ' Example: 1H-1S, 2S and opt = 6S. We can bid 4N ace ask '
                    ' or make a control bid with a void or worthless dbltn.  '
                    ' In this example, bidder has AKQ9853-T92--AQ9 with both '
                    ' worthless hearts and a diamond void, so a cue bid is   '
                    ' called for, but it takes a lot of code to test for this.
                    Stop
                    
                    For s = 1 To 4
                        If int_SuitLen(s) > 1 And _
                           int_HCP(s) < 3 And _
                           Right$(pardsLastCall, 1) <> _
                              Mid$("CDHS", s, 1) _
                        Then
                            ' Qx(x) or worse. '
                            ' Find the next lower non-opt suit '
                            ' with a 1st or 2nd round control: '
                            i = s - 1
                            If i = 0 Then i = 4
                            Do While int_HCP(i) < 4 And _
                                     int_SuitLen(i) > 1
                                i = i - 1
                                If i = 0 Then i = 4
                            Loop
                            ' Q (control) bid: '
                            b = (hiBidLvlUs + 2) & _
                                 Mid$("CDHS", i, 1) & "@Q"
                            Exit For
                        End If
                    Next
                    If b = "" Then
                        b = "4N@A" & Right$(pardsLastCall, 1)
                    End If

                    If redoBidPt = "70" Then Stop
                    bidPt = "70'"

                    GoTo SaveBid
                End If

            ElseIf hiBidLvlOpp > 0 Then
                ' If we can overbid RHO in our optimum '
                ' by 1 then we should do it. '
                If int_OptAvg > hiBidLvlOpp Then
                    ' Can't raise pard with 1. (Deal 120104) '
                    ' Can't bid NT because NT's pard's slam. '
                    ' Need to rebid 6-card club suit. '
                    Stop ' MsgBox "Stop ComputeOptBid 6" ' not right
                    If optSuitNumUs > hiBidLvlOpp Then
                        ' Bidder has higher suit: '
                        b = hiBidLvlOpp & optSuitStrUs
                    Else ' Opps have bid a higher suit. '
                        b = hiBidLvlOpp + 1 & optSuitStrUs
                    End If

                    If redoBidPt = "80" Then Stop
                    bidPt = "80"

                    GoTo SaveBid
                End If
            End If
        Else ' hiBidLvlUs <= hiBidLvlOpp '
            'Stop ' MsgBox "Stop ComputeOptBid 6" ' The following may not be right: '
            If optSuitNumUs = 5 Then
                j = 0
            Else
                j = int_SuitLen(optSuitNumUs) ' j = len of opt suit. '
            End If

            If hiBidLvlUs = 0 And _
               int_HCP(0) < 11 And _
               j < 5 _
            Then
                ' KQ3-63-9764-KQ85 Over 1H, can't bid anything. '
                GoTo SaveBid
            ElseIf optSuitNumUs = 5 Then
                'MsgBox "Stop ComputeOptBid 7"
                b = hiBidLvlOpp & "N"

                If redoBidPt = "89" Then Stop
                bidPt = "89"

                GoTo SaveBid
            ElseIf hiBidLvlUs < hiBidLvlOpp And _
                   int_SuitLen(optSuitNumUs) > 5 And _
                   int_HCP(0) > 11 _
            Then
                ' Opps have outbid us but we have '
                ' points and a good suit. '
                b = (hiBidLvlOpp + 1) & _
                    optSuitStrUs

                If redoBidPt = "90" Then Stop
                bidPt = "90"

                GoTo SaveBid
            ElseIf optMaxLvlUs >= hiBidLvlOpp And _
                   int_CallNum > 2 _
            Then
                If redoBidPt = "100" Then
                    redoBidPt = ""
                    Stop
                End If
                
                If str_Bid(1) = "1N" And str_Bid(2) = "D" And str_Bid(3) = "R" Then
                    If str_Bid(5) = "P" And (str_Bid(4) <> "" Or str_Bid(6) <> "") Then
                        If int_BidNum = 9 And str_Bid(7) = "D" And str_Bid(8) = "P" Then
                            b = "P"
                        Else
                            b = "D"
                        End If
                        bidPt = "100"
                        GoTo SaveBid
                    End If
                End If
                
                If optSuitNumUs - hiBidLvl < 3 And _
                   InStr(allBids(3), optSuitStrUs) > 0 _
                Then
                    ' Don't make big jump to game if '
                    ' pard has not bid the suit yet. '
                    If j = 4 Or _
                       (j = 3 And _
                        int_HCP(optSuitNumUs) > 5) _
                    Then
                        b = "D"
                    ElseIf j > 4 Then
                        b = optMaxLvlUs & optSuitStrUs
                    ElseIf hiOptBidLvlUs > 0 And _
                           hiOptBidLvlUs < 4 And _
                           j >= 3 _
                    Then
                        b = hiBidLvlOpp + 1 & _
                            optSuitStrUs ' same as below
                    End If
                ElseIf optSuitNumUs <> hiBidSuitOpp Then
                    j = hiBidLvlOpp + oppOptSuitHigher
                    If j - optMaxLvlUs < 2 Then
                        b = j & _
                        optSuitStrUs
                    End If
                End If
                If b <> "" Then
                    bidPt = "100"
                    GoTo SaveBid
                End If
            End If
        End If
    End If

    If redoBidPt = "110" Then Stop
    bidPt = "110"
    
'Stop ' If n = 0 then Opt suit has not been bid yet. '
    
    
    If hiOptBidLvlUs > 0 And _
       hiBidLvlOpp - optGameLvlUs < 2 _
    Then
        ' Optimum suit has been previously bid. '
        ' 3 possibilities now exist: '
        ' 1. Our hiBid was in the Optimum suit and
        '    a. Our hiBid was > oppHiBid, then
        '       our next bid should be hiBid + 10
        '    b. Our hiBid as < oppHiBid, then
        '       1).

        ' opt suit was previously bid '
        If optGameLvlUs = 0 Then
            ' Can only make part score. '
            If (hiOptBidLvlUs * 10 + optSuitNumUs) > (hiBidLvlUs * 10 + hiBidSuitUs) Then
                If optSuitNumUs > hiBidLvlUs And _
                   hiBidLvlUs > 0 _
                Then ' D > C '
                    b = (hiBidLvlUs) & Mid$("CDHSN", optSuitNumUs, 1)
                
                ElseIf hiBidLvlUs > hiOptBidLvlUs Then
                    Stop ' MsgBox "Stop ComputeOptBid 8"
                
                ElseIf optSuitNumUs > hiBidSuitOpp Then
                    ' Ex: hi bid = 2D and int_OptAvg = 3C  '
                    b = (hiBidLvlUs + 1) & Mid$("CDHSN", optSuitNumUs, 1)
                End If
                ' If hiBid is > int_OptAvg, we have a problem. '
            Else
                ' We've bid high part score in the Opt suit  '
                ' and we can't make game, so pass it out: '
                GoTo SaveBid
            End If
        ElseIf (optMaxLvlUs - hiBidLvlUs = 1 And _
               optGameLvlUs > 0) Or int_Opt4OppAvg = 0 _
        Then
            ' Last bid of opt suit was 1 below optGameLvlUs '
            ' and optGameLvlUs
            b = optGameLvlUs & optSuitStrUs
        ElseIf int_HCP(0) > 9 Then
            k = Val(hiBidLvlUs)
            If hiBidLvlOpp > 0 Then
                n = hiBidLvlOpp
            End If
            If k = 2 Then
                b = optGameLvlUs & optSuitStrUs
            ElseIf hiOptBidLvlUs > 34 And _
                   optGameLvlUs > 5 _
            Then
                Stop ' MsgBox "Stop ComputeOptBid 9" ' I don't get this. '
                b = optMaxLvlUs & optSuitStrUs
            Else
                If (optSuitNumUs > hiBidLvlOpp) And k >= n Then
                    ' We have the higher strain so '
                    ' bid the same level as them.  '
                    b = n & optSuitStrUs
                ElseIf n > k Then
                    ' They have the higher strain/level '
                    ' so we have to bid up a level. '
                    b = (n + 1) & optSuitStrUs
                End If
            End If

        Else
            GoTo SaveBid
        End If
    ElseIf hiBidLvlOpp - optGameLvlUs < 2 Then
        'If hiBidLvlOpp > hiBidLvlUs Then hiBidLvlUs = hiBidLvlOpp
        j = hiBidLvl ' Level of prior high bid of any suit by anyone. '
        ' Example: optBid = 4H, hiBid so far = 2C by pard. RHO passed. '
        ' I opened 1D. Let's see if I can bid 2H: '
        If hiBidLvlUs Mod 10 > optSuitNumUs Then
            ' Ex: if hiBidUs=2S then S=4 and optSuitNumUs = 2, '
            ' so we have to move up a level to 3s. '
            j = j + 1 ' up a bidding level '
        End If

        If j * 10 + optSuitNumUs > hiBidLvlUs And j > 0 Then
            ' j=2 + H=2 = 22.  hiBid = 21 (2C). '
            ' Now let's see if I have enough H's.
            If optSuitNumUs = 5 Then
                k = 5
            Else
                k = int_SuitLen(optSuitNumUs)
            End If
            If k >= 4 And optSuitNumUs < 5 Then
                If int_HCP(0) < 12 And _
                   (k < 6 Or _
                      int_HCP(optSuitNumUs) < 5) _
                Then
                    GoTo SaveBid
                End If
                ' With 4+ of Optimum suit, bid it. '
                b = j & optSuitStrUs
            ElseIf optSuitNumUs = 5 And _
                   hiBidLvlUs > 0 And _
                   int_HCP(0) > 17 _
            Then
                ' Optimum is NT and pard has bid and hand has 17+ HCP. '
                b = "3N"
            ElseIf optSuitNumUs - hiBidLvlUs = 1 Then
                b = hiBidLvlUs & Mid$("CDHSN", k, 1)
            Else
                ' Find another suit to bid that hasn't been bid yet. '
                ' Not interested in raising pard's non-optimum suit. '
                ' Want to bid a new suit/NT to give pard incentive   '
                ' to bid the optimum suit/NT. '
                n = optSuitNumUs + 1
                Do
                    b = ""
                    If n > 5 Then
                        n = 1
                        j = j + 1
                    End If
                    If n = optSuitNumUs Then Exit Do
                    b = j & Mid$("CDHSN", n, 1)
                    If n = 5 Then GoTo SaveBid
                    If InStr(str_PriorBids, b) = 0 Then
                        If int_SuitLen(n) > 3 Then Exit Do
                    End If
                    n = n + 1
                    DoEvents
                Loop
                If optSuitNumUs = 5 Then ' NT '
                    If hiBidLvlUs > 1 Then
                        Stop ' MsgBox "Break - ComputeOptBid 10"
                        ' If 2N is biddable, see if 3N is reasonable. '
                        b = "3N"
                    End If
                Else
                    Stop ' MsgBox "Break - ComputeOptBid 11"
                    ' May not be higher than hiBid if unbid suit was lower. '
                    b = hiBidLvlUs & Mid$("CDHSN", n, 1)
                End If
            End If
        End If
    End If
    GoTo SaveBid

SaveBid:
    If b = "" Then b = "P"
    'Stop
    Dim dp(2) As Integer

    'If b <> "P" Then
        redoBidPt = ""
        X = b
        'If bln_ShowBidPosSet = False Then
            ' PlayScr.Left + 2385 = 4500 -- PlayScr.Top + 3960 = 11800
        Open App.Path & "DealerPosition.txt" For Input As #5
            Input #5, dp(0)
            Input #5, dp(1)
        Close #5
            
        lng_ShowBid(0) = dp(0): lng_ShowBid(0) = dp(0) + 2500 ' 4500 ' left '
        lng_ShowBid(1) = dp(1): lng_ShowBid(1) = dp(1) + 7850 ' 11800 ' top '
        bln_ShowBidPosSet = True
        'End If
        b = InputBox("Click Ok to save computed call" & vbCr & vbCr & _
                     "or enter a different call including P" & vbCr & vbCr & _
                     "or enter x to show P but not save it" & vbCr & vbCr & _
                     "or enter e to edit computation.", "Changing and/or Saving Calls", b, lng_ShowBid(0), lng_ShowBid(1))
        If b = "X" Or b = "E" Then b = LCase(b)
        If b = "p" Then b = "P"
        If b = "e" Then
            b = ""
            redoBidPt = bidPt
            bidPt = ""
            If bidPt = "Sac" Then
                GoTo Check4Sac
            Else
                GoTo WeBid1 ' It stops at redoBidPt. '
            End If
        ElseIf b = "x" Or b = "" Then
            b = "x"
        Else
            ' The bid entered above will be saved below. '
        End If
    'End If

    If b <> "x" Then ' Deal 120126 East bid 4H over 4S.
        str_BidToSave = b
        str_Call(int_CallNum) = b
        SaveBidEntry
    End If
Exit Sub

ComputeBidError:
    MsgBox Err.Description & vbCr & _
           "- error in ComputeOptBid"
    Stop
    Resume 'Next
    On Local Error GoTo 0
End Sub

