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
'           In VB6 under Tools, Options, Editor Format * So I'm using a black background now:         '
'           I use the following screen colors:         *
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

Private lng_HandVal_Left As Long
Private lng_HandVal_Top As Long
Private bln_ShowBefore As Boolean
Private bln_ShowAfter As Boolean
Private bln_ShowBySuit As Boolean
Private bln_ShowNever As Boolean
'Client input variables:
Private bln_AlreadySearching As Boolean
Private str_AnalyzeBid As String 'Bid(s) to show analysis for; e.g.: 1S
Private bln_Cancel As Integer   ' Emergency exit from this module
Private bln_CheckAllPcts As Boolean
Private str_DBPathAndName As String
Private bln_GetDisclosure As Boolean  ' Flag: Client wants disclosures?
Private str_PlayersHand As String
Private str_PriorCalls As String
Private str_ScoreMethod As String 'I=IMPs, M=MPs
Private lng_TestID As Long 'ID# of entry that test hand should match.
Private bln_TestEntry As Boolean ' Test only the entry for lng_TestID.
Private str_Vul As String
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
Private bln_AnalyzeBid As Boolean 'Flag: Has bid analysis been requested?
Private bln_AnalysisShowing As Boolean ' Flag: Analysis form showing
Private int_BidNum As Integer   ' Current bid #, starting with first non-pass
Private int_BidRound As Integer ' Bidding round, starting with 1st seat
Private str_Bid(50) As String ' Bids, starting with first non-pass
Private str_Call(50) As String  ' Array of prior bids/calls including starting passes
Private int_CallNum As Integer  ' Current call #, starting with 1st seat
Private str_Disclosure As String  ' Give client all possible bid meanings
Private int_DistribPts(4) As Integer  ' Distribution points by suit/hand
Private bln_GettingDiscl As Boolean  ' Entry found - getting disclosures
Private int_HCP(4) As Integer   ' High Card Points by suit/hand
Private lng_ID As Long  'ID# of entry matching hand's specs
Private int_Intermediates(4) As Integer ' # of intermediates by suit/hand
Private int_1NT_Min As Integer ' Minimum HCPs for opening 1N.
Private int_1NT_Max As Integer ' Maximum HCPs for opening 1N.
Private int_2NT_Min As Integer ' Minimum HCPs for opening 2N.
Private int_2NT_Max As Integer ' Maximum HCPs for opening 2N.
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
Private int_SuitLen(4) As Integer  ' # of cards in each suit
Private str_SuitNames(4) As String ' Clubs, Diamonds, Hearts, Spades
Private bln_TrumpQueen As Boolean ' for RKCB '
Private bln_TrumpKing As Boolean ' ditto '
Private bln_TrumpAce As Boolean ' for Grand Slam Force [
Private int_TrumpSuit As Integer ' C=1, D=2, H=3, S=4 - For Roman KC Blackwood, etc. '
Private bln_xField As Boolean      ' Suit bid matches "x" specs (T/F)
Private str_xSuit As String        ' Suit bid when matching "x" specs.
' BidBase Hand Valuation variables:                   '
Private str_CardsSpec As String
Private str_CardsSpecCombo As String
Private str_CardsSpecExtracted As String
Private str_CardsSpecOperator As String
Private str_FlagSpec As String
Private str_FitSpec As String
Private str_HandShape As String
Private str_HandSpec As String
Private int_HCPn(4) As Integer
'Private int_Intermediates(4) As Integer
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
    On Local Error Resume Next
    rs_BidNames.Close
    rs_BidTable.Close
    rs_BidsQuery.Close
    db_BidBase.Close
    ws_BidBase.Close
    On Local Error GoTo 0
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

Public Property Let PriorCalls(ByVal strClientInput As String)
    ' Prior Calls, including starting Passes. '
    ' Format: "P-1C-1N-2C-D-R-2S", etc.       '
    If Not bln_AlreadySearching Then
        str_PriorCalls = UCase$(Trim$(strClientInput))
    End If
End Property

Public Property Let Vul(ByVal strClientInput As String)
    ' Vulnerability set by client:           '
    '     WE, THEY, NEITHER, BOTH            '
    ' Vulnerability set by editor:           '
    '     Y, N, Fav, Unfav, Eq, Eq+          '
    ' If testing an entry, vul may be blank. '
    If Not bln_AlreadySearching Then
        strClientInput = UCase$(Trim$(strClientInput))

        Select Case strClientInput
            Case "WE"
                str_Vul = " Y Unfav "
            Case "THEY"
                str_Vul = " N Fav "
            Case "NEITHER"
                str_Vul = " Eq N Eq+ "
            Case "BOTH"
                str_Vul = " Eq Y Eq+ "

            Case "Y"
                str_Vul = " Y "
            Case "N"
                str_Vul = " N "
            Case "FAV"
                str_Vul = " Fav "
            Case "UNFAV"
                str_Vul = " Unfav "
            Case "EQ"
                str_Vul = " Eq "
            Case "EQ+"
                str_Vul = " Eq+ "

            Case ""
                str_Vul = ""

            Case Else
                MsgBox Chr$(34) & strClientInput & Chr$(34) & _
                " is not a valid 'Vulnerability' specification."
        End Select
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
    int_TrumpSuit = intClientInput
End Property


'                     Pass data back to Client                          '

Public Property Get Bid() As String
    If Not bln_AlreadySearching Then
        bln_AlreadySearching = True
        int_ErrCode = 0
        str_Disclosure = ""
        LookUpBid
        If str_AnalyzeBid <> "" Then Unload frm
        bln_AlreadySearching = False
        Bid = str_Call(int_CallNum)
        ' str_Call(int_CallNum) comes from '
        ' Sub GetDataFromMatchingEntry     '
    End If
End Property

Public Property Get id() As Long
    ' The ID of the entry for the bid, if any. '
    If Not bln_AlreadySearching Then
        id = lng_ID
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
    If Not bln_AlreadySearching Then
        ErrMsg = str_ErrMsg
    Else
        ErrMsg = "Already searching for bid."
    End If
    str_ErrMsg = "" ' Clear for next try. '
    'Unload frm
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
    '97: "Invalid vulnerability specified."               '
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
                             PosNum As Integer, _
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
    PosNum = int_PosNum
    ' Hand evaluatorion: '

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

    If int_ErrCode > 0 Then
        ' Input error from client. Abort. '
        Exit Sub
    End If
    bln_ShowNever = True

    If Not bln_AlreadyInitialized Then
        InitializeBidBaseDLL
    End If

    ' Analyze hand before testing entries: '
    AnalyzeHand
    If int_ErrCode Then Exit Sub

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
        Else
            str_ErrMsg = _
                "No entries found matching the prior bids: " & vbCrLf & _
                str_PriorCalls & vbCrLf & _
                " (Note: opening passes are ignored.)"
        End If
        int_ErrCode = 1
        Exit Sub
    End If

    ' Loop through entries with matching Prior Bids. '
    ' ("ZZZZZZ" is a dividing line entry that marks  '
    '   the end of a Prior Bids section. It may not  '
    '   always be used, but it should be treated as  '
    '   a stopping point when it has been used.)     '

    Do While rs_BidTable!Bid_Order_Number <> "ZZZZZZ" And _
             rs_BidTable!prior_bids = str_PriorBids

      ' See if this bid was specified by client to '
      ' be analyzed, and if so, show the details:  '
      If str_AnalyzeBid = "ALL" Or _
         InStr(str_AnalyzeBid, rs_BidTable!Bid) > 0 _
      Then
          ShowAnalysis

      ElseIf str_AnalyzeBid <> "" And _
             Right$(rs_BidTable!Bid, 1) = "x" And _
             Left$(str_AnalyzeBid & " ", 1) = Left$(rs_BidTable!Bid, 1) _
      Then
          s = Right$(str_AnalyzeBid, 1) ' Suit
          If (s = "C" And InStr(rs_BidTable!cNum, "x")) Or _
             (s = "D" And InStr(rs_BidTable!dNum, "x")) Or _
             (s = "H" And InStr(rs_BidTable!hNum, "x")) Or _
             (s = "S" And InStr(rs_BidTable!sNum, "x")) _
          Then
              ShowAnalysis
          Else
              Unload frm
              bln_AnalyzeBid = False
          End If

      Else
          Unload frm
          bln_AnalyzeBid = False
      End If

      ' Compare hand to entry specifications: '
      ' (Suits_Ok is last because it is the   '
      '  longest and most complicated test.)  '
      If PctUsed_Ok Then
       If TotalHCPs_Ok Then
        If CustPoints_Ok Then
         If DistribPts_Ok Then
          If Winners_Ok Then
           If Losers_Ok Then
            If QuickTricks_Ok Then
             If Stoppers_Ok Then
              If Shape_Ok Then
               If Intermed_Ok Then
                If Vul_Ok Then
                 If Position_Ok Then
                  If Suits_Ok Then
                   If Eval_Pts_Ok Then
                       '     MATCHING ENTRY FOUND                  '
                       GetDataFromMatchingEntry
                       If str_ErrMsg = "Keep testing" Then
                           str_ErrMsg = ""
                       Else
                           If bln_AnalyzeBid Then
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
                                       Exit Sub
                                   End If
                                End With
                           End If
                           Exit Sub
                       End If
                       '                                           '
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
      End If

      xxx = DoEvents()
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

      If bln_AnalyzeBid Then
          With frm
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
              If lng_TestID = rs_BidTable!id Then
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
          End With
      End If

      If lng_TestID = rs_BidTable!id Then
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
    Loop

    int_ErrCode = 2
    str_ErrMsg = "No (additional) entry found for this hand."
    lng_ID = 0

End Sub


Private Sub InitializeBidBaseDLL()
    If bln_AlreadyInitialized Then Exit Sub
        ' To avoid a time-out when loading the DLL,     '
        ' initialization does not take place until the  '
        ' first time the DLL is called by the Client,   '
        ' then a flag is set below to keep this routine '
        ' from executing again on subsequent calls.     '
        Set ws_BidBase = DBEngine.Workspaces(0) ' Not sure if wsp is needed.
        Set db_BidBase = ws_BidBase.OpenDatabase("BidBase.accdb")
        Set rs_BidTable = db_BidBase.OpenRecordset("Bids") ' Table in the db to use.
        Set rs_BidNames = db_BidBase.OpenRecordset("BidNames")
        Set rs_BidsQuery = db_BidBase.OpenRecordset("GridQuery") 'Set rstProducts = dbsNorthwind.OpenRecordset("Current Product List")


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
    If int_ErrCode > 0 Then Exit Sub

    ' PctUsed is calculated in PctUsed_Ok. '
    ' TotalHCPs are calculated in the Sub  '
    '     ConvertHandStringToArrays, above.'
    ' CustPoints calculated in CustPoints_Ok.
    ' DistribPts calc below in Calculate_Shape.'
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
    GetMinNTpts
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
    '*************************************************

    Dim pb As String
    Dim pc As String
    Dim i As Integer

    Erase str_Bid()
    Erase str_Call()

    str_PriorBids = ""
    
    If str_PriorCalls = "" Then
        int_BidNum = 1
        int_CallNum = 1
        int_BidRound = 1
        int_PosNum = 1
        str_Position = "12NFU"
        ' 1=1st seat              '
        ' 2=2nd seat              '
        ' N=unpassed,             '
        ' F=First/Second seat,    '
        ' U=responder is unpassed '
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
    Dim v As Integer   ' void       '

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
            int_DistribPts(s) = int_DistribPts(s) + 1

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
        str_Shortness = "v"

    ElseIf s1 > 0 Then
        ' We have a singleton.'
        str_Shortness = "s"

    End If

End Sub


Private Sub Calculate_Position()
    ' Calculate "position" variable where '
    '   Y = bidder is a passed hand       '
    '   N = bidder is an unpassed hand    '
    '   1 = bidder is in 1st seat         '
    '   2 = bidder is in 2nd seat         '
    '   3 = bidder is in 3rd seat         '
    '   4 = bidder is in 4th seat         '
    '   F = opener is 1st or 2nd seat     '
    '   T = opener is 3rd or 4th seat     '
    '   P = partner is a passed hand      '
    '   U = partner is an unpassed hand   '
    ' The variable str_Position can have  '
    ' more than one variable, because the '
    ' codes are not mutually exclusive and'
    ' an entry may specify 1 code rather  '
    ' than another, both of which match   '
    ' the current bidder's position.      '
    ' Example: Y3TP - see the codes above '
    ' That's why you see --               '
    '  str_Position = str_Position & ...  '

    Dim p As String

    ' See if bidder unpassed:  '
    If int_CallNum < 5 Then
        str_Position = "N" ' = unpassed hand '

    ElseIf str_Call(int_PosNum) <> "P" Then
        ' If more than 4 calls, then  '
        ' this player's position num  '
        ' (e.g.: 3) is also the array '
        ' pointer to his first call in'
        ' the array, like str_Call(3).'
        str_Position = "N" ' = unpassed hand '

    Else
        str_Position = "Y" ' = passed hand '
    End If

    ' Get seat number, 1-4: '
    If int_CallNum < 5 Then
        str_Position = str_Position & int_CallNum
    End If

    ' Add F, T: '
    If int_PosNum < 3 Then
         ' In 1st/2nd chair: '
        str_Position = str_Position & "F"
    Else ' In 3rd/4th chair: '
        str_Position = str_Position & "T"
        If int_BidNum = 1 Then
            ' Opener '
            str_Position = str_Position & int_PosNum
        End If
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


Sub GetMinNTpts()
    ' If partner opened 1N, get minimum expected '
    ' HCPs for his NT range. If 15-17, then 15.  '
    ' If 15-17 but sometimes a good 14, then the '
    ' normal expected minimum is still 15.       '
    ' This is used to compute point ranges for   '
    ' responding; otherwise, BidBase would have  '
    ' to have duplicate sets of entries for each '
    ' range of NT HCPs. The idea is used for 2N  '
    ' opening ranges as well.                    '

    Dim n As String
    Dim s As String
    Dim b As String
    Dim i As Integer
    Dim triedOnce As Boolean

    b = Left$(str_PriorBids & "  ", 2)

    If Not _
        (((b = "1N" Or _
          b = "2N") And _
         (int_BidNum = 3 Or _
          int_BidNum = 5)) Or (str_PriorBids = "")) _
    Then
        Exit Sub ' Must be responder or opener rebidding.
    End If

    If b = "1N" Then
        n = "130000"

    ElseIf b = "2N" Then
        n = "110000"

    End If

'Stop ' Don't understand. When S is opener/dealer, b=1N, n=13000, 1NT_min=25 '
     ' but when S is responder, b=2N, n=11000, and 1NT_min=15 ??? '

    For i = 1 To 2
        If b = "1N" Then
            n = "130000"
        ElseIf b = "2N" Then
            n = "110000"
            i = 2
        ElseIf i = 1 Then
            n = "130000"
        Else
            n = "110000"
        End If

On Local Error GoTo DBClosed

        s = "SELECT * FROM Bids " & _
            "WHERE Prior_Bids = '' " & _
            "AND Bid_Order_Number = '" & n & "' " & _
            "ORDER BY Bid_Order_Number;"
        Set rs_BidTable = db_BidBase.OpenRecordset(s, dbOpenSnapshot)
        If rs_BidTable.NoMatch Or rs_BidTable.EOF Then
            str_ErrMsg = _
                "Opening " & b & _
                " entry #" & n & _
                " not found." & vbCrLf & _
                "Cannot get opening " & b & " range."
            int_ErrCode = 1
            Exit Sub
        End If

        With rs_BidTable
            If Not (.NoMatch Or .EOF) Then
                Do While !Pct_Used <> "" And _
                         !Bid_Order_Number = n And _
                         Not .EOF
                    .MoveNext
                    If .EOF Then Exit Do
                Loop
                If i = 1 Then
                    int_1NT_Min = Val(!HCP)
                    int_1NT_Max = Val(Right$("  " & !HCP, 2))
                Else
                    int_2NT_Min = Val(!HCP)
                    int_2NT_Max = Val(Right$("  " & !HCP, 2))
                End If
            End If
        End With
        If b = "1N" Then Exit For
    Next

ExitGetMinNTpts:
    On Local Error GoTo 0
    Exit Sub

DBClosed:
    If triedOnce Then
        Resume ExitGetMinNTpts
    End If
    triedOnce = True
    InitializeBidBaseDLL
    Resume

End Sub


'                      Start entry analysis routines                          '


Private Function PctUsed_Ok() As Boolean
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
    Dim BON As Long 'Bid Order Number
    Dim num As Long
    Dim sel As String
    Dim sel2 As String
    Dim nam As String

    If bln_TestEntry Then
        ' Must accept this entry despite Pct setting.'
        PctUsed_Ok = True
        Exit Function
    End If

    With rs_BidTable
        ' Check convention's "BidNames.Selected" code:
        num = Val("" & !ConvID)
        LookUpBidName num, sel, sel2, nam ' sel = "selected"
        If Not (sel = "" Or !id = lng_TestID Or _
               (sel = "W" And (int_PosNum = 1 Or int_PosNum = 3)) Or _
               (sel = "T" And (int_PosNum = 2 Or int_PosNum = 4))) Or _
           Not (sel2 = "" Or !id = lng_TestID Or _
               (sel2 = "W" And (int_PosNum = 1 Or int_PosNum = 3)) Or _
               (sel2 = "T" And (int_PosNum = 2 Or int_PosNum = 4))) _
        Then   ' "W"=We  "T"=They  ""=both '
            If bln_AnalyzeBid Then
              If sel <> "0" Or sel2 <> "" Then
                With frm.lstAnalysis
                    .ListIndex = 6
                    If sel = "0" Then
                        .List(6) = Space$(30) & "Convention not activated by user."
                    ElseIf sel2 = "0" Then
                        .List(6) = Space$(25) & "This Convention's sub not activated by user."
                    Else
                        .List(6) = Space$(20) & "Cnv = " & sel & _
                                   Space$(10) & "Sub = " & sel2
                    End If
                End With
              Else
                frm.lstAnalysis.List(6) = Space$(30) & "Entry is not activated by user. Skip to next entry."
              End If
            End If
            Exit Function
        End If

        If lng_TestID > 0 And Val(!Pct_Used) > 0 Then
            If lng_TestID = Val(!id) Then
                PctUsed_Ok = True
            End If
            Exit Function
        End If

        If Trim$("" & !Pct_Used) = "" Or _
           bln_CheckAllPcts Or _
           Trim$("" & !Pct_Used) = str_ScoreMethod _
        Then

            PctUsed_Ok = True
            Exit Function

        ElseIf Val(!Pct_Used) > 0 Then
            ' BidBase entry has a Pct_Used number.    '
            ' Get a random number and see if it falls '
            ' within the specified range (i.e.:       '
            ' random number <= specified Pct          '
            ' Example: If Pct = 20 and random # = 13, '
            '          then Ok, but not if # = 32.    '

            BON = !Bid_Order_Number
            Randomize Timer
            pct = Fix(Rnd * 100 + 1)

            Do

                If (pct - ttlPct) <= Val(!Pct_Used) Or _
                   .EOF Or _
                   lng_TestID > 0 _
                Then
                    ' If end of file, last entry must '
                    ' be Okayed for further testing:  '
                    ' Or when testing, accept any Pct.'
                    PctUsed_Ok = True
                    GoTo PctOk
                End If

                ttlPct = ttlPct + Val(!Pct_Used)

                .MoveNext
                If BON <> !Bid_Order_Number Or _
                   Val(!Pct_Used) = 0 _
                Then
                    ' Alternative entries must have the '
                    ' same BON. This entry doesn't, so  '
                    ' go back to prior entry and OK it: '
                    On Local Error Resume Next
                    .MovePrevious
                    On Local Error GoTo 0
                    PctUsed_Ok = True
                    GoTo PctOk
                End If
            Loop
            ' No Do-Loop limits because exits are made   '
            ' if a Pct matches or if EOF or if Bid Order '
            ' Number changes or if next PctUsed field no '
            ' longer has a number in it.                 '


        ElseIf (Left$(!Pct_Used & " ", 1) = "M" Or _
                Left$(!Pct_Used & " ", 1) = "I") And _
               (Left$(!Pct_Used & " ", 1) = str_ScoreMethod Or _
                lng_TestID > 0) _
        Then  ' "M"=matchpoints   "I"=IMPs   "D"=Double Dummy Analysis '
              ' A number of 1-99 in the Pct field allows BB to randomly select a bid the given % of the time. '
              ' x followed by a number disables the % field but allows you to reinstate by by removing the x. '
            PctUsed_Ok = True
            Exit Function
        End If

    End With


    If bln_AnalyzeBid Then
        With frm.lstAnalysis
            .ListIndex = 6
          If rs_BidTable!Pct_Used = "0" Then
            .List(6) = Space$(30) & "Entry is not activated by user."
          Else
            .List(6) = Space$(40) & "Pct Used = " & _
                       rs_BidTable!Pct_Used
          End If
        End With
    End If
    Exit Function

PctOk:
    ' Redo analysis form for new entry: '
    If bln_AnalyzeBid Then
        ShowAnalysis
    End If

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

    strHCP = LCase$(Trim$("" & rs_BidTable!HCP)) ' Entry's specs '

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
            Left$(strHCP, 3) = "min" Or _
           Left$(rs_BidTable!Bid & "     ", 5) = "hcp2n" _
    Then
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

    ElseIf Left$(strHCP & " ", 1) = "P" Then
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
                int_Intermediates(0) = int_Intermediates(0) + adj
            End If
        End If

    ElseIf bln_AnalyzeBid Then
        With frm.lstAnalysis
            .ListIndex = 4
            .List(4) = .List(4) & _
                "            NO MATCH:  Hand has " & int_HCP(0)
        End With
    End If

End Function


Private Function CustPoints_Ok()
    '*****************************************************
    ' Purpose:  Custom point assignments can be made for *
    '           A-K-Q-J-T.  Example: A=2, K=1.           *
    '                                                    *
    ' Unlike other tests, Custom Points for the hand can *
    ' not be computed in advance. They change with each  *
    ' entry depending on what values have been assigned  *
    ' to the honors by each entry.                       *
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

    ' Using  _"" &_  before DAO field '
    ' prevents a null error from DAO. '
    ' Could use On Error Resume Next, '
    ' but that can hide other errors.)'

With rs_BidTable

    ' Convert entry's CP specs to array: '
    cp(1) = Val("" & !A_Cust)
    cp(2) = Val("" & !K_Cust)
    cp(3) = Val("" & !Q_Cust)
    cp(4) = Val("" & !J_Cust)
    cp(5) = Val("" & !T_Cust)
    cp(0) = cp(1) + cp(2) + cp(3) + cp(4) + cp(5)

    For s = 1 To 5
        If cp(s) > 0 Then
            cardsSpecified = True
            Exit For
        End If
    Next

    ' Required (total) custom points: '
    rp = "" & !Cust_Pts
    ReqdPts = Val(rp)

    ' Identify math sign for later use: '
    If ReqdPts = 0 And _
       Len(rp) > 1 _
    Then 'ex: <4, >4

        ReqdPts = Val(Mid$(rp, 2))
        If Left$(rp, 1) = "<" Then
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

    If bln_AnalyzeBid Then
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

    'Set distibution points required:
    tdp = Trim$("" & rs_BidTable!Dist_Pts)

    If tdp = "" Then
        ' None required. '
        DistribPts_Ok = True
        Exit Function
    End If

    ' Compute tdpn and see if there '
    ' is a math sign:
    tdpn = Val(tdp)
    If tdpn = 0 And Len(tdp) > 1 Then
        ' Must be a math sign in front '
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
                    int_Intermediates(0) = int_Intermediates(0) + adj
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

    If bln_AnalyzeBid Then
        With frm.lstAnalysis
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

    If Trim$("" & rs_BidTable!Winrs) = "" Or _
       sng_NumWinners(0) >= Val("" & rs_BidTable!Winrs) _
    Then

        Winners_Ok = True

    ElseIf bln_AnalyzeBid Then
        ' Failed test. '
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

    Losers = Trim$("" & rs_BidTable!Losers)
    If Trim$(Losers) = "" Or _
       (int_NumLosers(0) >= Val(Losers) And _
        Val(Losers) <> 0) _
    Then

        Losers_Ok = True
        Exit Function

    ElseIf Left$(Losers & " ", 1) = "<" Then
        If int_NumLosers(0) < Val(Mid$(Losers & " ", 2)) Then
            Losers_Ok = True
            Exit Function
        End If

    End If

    If bln_AnalyzeBid Then
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

    ElseIf bln_AnalyzeBid Then
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

    ElseIf bln_AnalyzeBid Then
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
       bln_AnalyzeBid _
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

    With rs_BidTable

        If "" & !Intermed = "" Then
            ' None specified '
            Intermed_Ok = True
            Exit Function
        End If

        i = Val(!Intermed)

    End With

    ' Test: '

    If int_Intermediates(0) >= i Then
        ' Note: int_Intermediates(0) is adjusted '
        ' for extra HCPs in Sub TotalHCPs_OK     '
        Intermed_Ok = True

    ElseIf bln_AnalyzeBid Then
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
    '     Y = vulnerable        '
    '     N = not vulnerable    '
    '   N  = neither fav.       '
    '   F  Fav = favorable      '
    '   U  Unfav = unfavorable  '
    '   B  Eq = equal/both vul. '
    '     Eq+ = equal or fav.   '

    Dim v As String ' vul spec from BB '

    v = rs_BidTable!Vul
    
    If v = "Fav" Or v = "T" Then v = "F"
    If v = "Eq" Then v = "B"
    If v = "Unfav" Then v = "U"
    If v = "Eq+" Then v = "E"
    If v = "Y" Or v = "W" Then v = "U" ' Yes, we are vul and not them, so U = unfavorable vul. '


    If (InStr(str_Vul, v) > 0 And v <> "") Or _
       (v = "" And (str_Vul <> " N " Or lng_TestID = 0)) Or _
       (v = "N" And str_Vul = "") Or (v = "Eq+" And str_Vul <> "W") _
    Then
        ' If playing and the Vul spec is "", then   '
        ' any vulnerability is okay, but if testing '
        ' an entry, use that entry's specified vul. '
        Vul_Ok = True

    ElseIf bln_AnalyzeBid Then
        ' Test failed. '
        With frm.lstAnalysis
            .ListIndex = 13
            If str_Vul = " N" Then ' Eq '
                v = "none"
            ElseIf str_Vul = " B" Then ' Eq '
                v = "both"
            ElseIf str_Vul = " W" Then ' Unfav '
                v = "We"
            ElseIf str_Vul = " T" Then ' Fav '
                v = "They"
            Else
                v = "Eq+"
            End If
            .List(13) = .List(13) & _
                "            NO MATCH:  Vul is:  " & v
        End With
    End If
End Function


Private Function Position_Ok()
    ' Check position of current player. '
    ' (E.g.: 1st seat, 2nd, 3, 4, etc.) '

    Dim position As String

    position = Trim$(UCase$("" & rs_BidTable!POS))

    If position = "" Or _
       (lng_TestID > 0 And _
        lng_TestID = rs_BidTable!id) _
    Then
    '   When testing, ignore position if this '
    '   is the entry for the hand being tested.
        Position_Ok = True
        Exit Function
    End If

    If InStr(str_Position, position) > 0 Then
        Position_Ok = True

    ElseIf bln_AnalyzeBid Then

        With frm.lstAnalysis
            .ListIndex = 14
            .List(14) = .List(14) & _
                        "            NO MATCH:  Position number = " & _
                        str_Position
        End With
    End If
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

    Dim num1 As String
    Dim num2 As String
    Dim pts1 As String
    Dim pts2 As String
    Dim sOk(4) As Boolean
    Dim dbNum(4) As String
    Dim dbHCP(4) As String
    Dim s As Integer
    Dim bln_xField As Boolean
    Dim i As Integer
    Dim longest_xSuit As Integer
    Dim compareXsuit As Integer
    Dim compareXnum As String
    Dim compareXpts As String
    Dim cmpr As String

    With rs_BidTable
        dbNum(4) = !sNum
        dbHCP(4) = !sHCP
        dbNum(3) = !hNum
        dbHCP(3) = !hHCP
        dbNum(2) = !dNum
        dbHCP(2) = !dHCP
        dbNum(1) = !cNum
        dbHCP(1) = !cHCP
    End With

    ' Get rid of blanks: '
    For s = Clbs To Spds
        dbNum(s) = Replace(dbNum(s), " ", "", , 99)
        dbHCP(s) = Replace(dbHCP(s), " ", "", , 99)
    Next

    str_xSuit = ""

    Dim s1 As Integer
    Dim s2 As Integer
    Dim s3 As Integer
    Dim xNum As String

    xNum = Trim$("" & rs_BidTable!xNum)
    ' If xNum as "U" in it, go up the suits '
    ' from Clubs to Spades, else go down.   '
    If Right$(" " & xNum, 1) = "U" Then
        s1 = Clbs
        s2 = Spds
        s3 = 1
        xNum = Left$(xNum, Len(xNum) - 1)
    Else
        s1 = Spds
        s2 = Clbs
        s3 = -1
    End If

    For s = s1 To s2 Step s3

        num1 = dbNum(s)
        num2 = ""
        pts1 = dbHCP(s)
        sOk(s) = False
        bln_xField = False

        ' Split alternatives, if any: '
        i = InStr(num1, ",")
        If i > 0 Then
            num2 = Trim$(Mid$(num1, i + 1))
            num1 = Trim$(Left$(num1, i - 1))
        End If

        i = InStr(pts1, ",")
        If i > 0 Then
            pts2 = Trim$(Mid$(pts1, i + 1))
            pts1 = Trim$(Left$(pts1, i - 1))
        End If

        If num1 = "" And pts1 = "" Then
            sOk(s) = True

        Else

            If str_xSuit = "" Or _
               (int_SuitLen(s) > int_SuitLen(longest_xSuit) And _
                Right$(" " & xNum, 1) = "+") _
            Then
                ' Have not found a match '
                ' for "x" specs yet.     '
                If num1 = "x" Then
                    ' Change num1 from "x" '
                    ' to xNum's value:     '
                    num1 = xNum
                    bln_xField = True
                End If

                If pts1 = "x" Then
                    ' Change from 'x' to xHCP value: '
                    pts1 = Trim$("" & rs_BidTable!xHCP)
                    bln_xField = True
                End If

            Else ' Match for "x" already found.'
                If num1 = "x" Or pts1 = "x" Then
                    ' Normally will not have '
                    ' x for HCP but not Num. '
                    num1 = num2
                    pts1 = pts2
                    num2 = ""
                    pts2 = ""
                End If
            End If

            Do
                sOk(s) = Suit_Ok(s, num1, pts1)

                If sOk(s) = False Then
                    bln_xField = False
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

                        If pts2 <> "" Then
                            pts1 = pts2
                            i = InStr(pts1, ",")
                            If i > 0 Then
                                pts2 = Trim$(Mid$(pts1, i + 1))
                                pts1 = Trim$(Left$(pts1, i - 1))
                            Else
                                pts2 = ""
                            End If
                        ElseIf Left$(dbHCP(s), 1) = "x" Then
                            pts1 = ""
                        End If

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
                        If dbNum(s) = "x" And _
                           (dbHCP(s) = "" Or _
                            dbHCP(s) = "x") _
                        Then
                            ' If x, it is not an error '
                            ' unless ALL suits fail.   '
                            sOk(s) = True

                        ElseIf bln_AnalyzeBid Then
                            With frm.lstAnalysis
                                i = 23 - s
                                .ListIndex = i
                                .List(i) = .List(i) & _
                                    "            NO MATCH:  Suit = " & _
                                    Chr$(34) & str_Suit(s) & Chr$(34)
                            End With
                        End If
                        Exit Do
                    End If

                Else
                    If bln_xField Then
                        str_xSuit = Mid$("CDHS", s, 1)
                        If longest_xSuit = 0 Or _
                           int_SuitLen(s) > int_SuitLen(longest_xSuit) _
                        Then
                            longest_xSuit = s
                        End If

                    ElseIf Right$(" " & num1, 1) = "x" Or _
                           Right$(" " & pts1, 1) = "x" Then
                        compareXsuit = s
                        compareXnum = num1
                        compareXpts = pts1

                    End If

                    Exit Do

                End If

            Loop

        End If

    Next

    bln_xField = (str_xSuit <> "")

    'See if xNum or xHCP had a value not met by any suit
    If str_xSuit = "" And _
       (Trim$(rs_BidTable!xNum) <> "" Or _
        Trim$(rs_BidTable!xHCP) <> "") _
    Then
        If bln_AnalyzeBid Then
            With frm.lstAnalysis
                .ListIndex = 18
                .List(18) = .List(18) & _
                "            NO MATCH:  Doesn't match 'x' specs."
            End With
        End If
        Exit Function
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

    If Combo_Ok Then

        If sOk(1) And sOk(2) And sOk(3) And sOk(4) Then
            Suits_Ok = True
        End If

    End If

End Function


Private Function Combo_Ok() As Boolean
    ' Check combo specs, if any. '

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
    Dim num1 As String
    Dim pts1 As String
    Dim numOk As Boolean
    Dim numKC As Integer
    Dim gerber As Boolean
    Dim numKeys As Integer
    Dim ref As String
    Dim t As String
    Dim i As Long
    Dim signK As String
    Dim signQ As String
    Dim c As Long ' comma position '
    Dim p As String ' p = a temp variable for pts1 '
    Dim s As String ' s = an individual spec within pts1 '
    
    ' Check combined suits specs: '
    num1 = Trim$("" & rs_BidTable!cmNum) ' "" &  prevents an error if cmNum is null. '
    pts1 = Trim$("" & rs_BidTable!cmHCP)
    
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

    ' Slam Bidding specs: '
    
    If Left$(pts1 & "    ", 3) <> "SB:" Then
        GoTo SlamBiddingDone
    End If
    
    ' Specs are in the Suits - Combo-HCP box. '
    ' Like other entries, each entry has a    '
    ' test hand for verifying that the specs  '
    ' do what is expected. For some ace-asks, '
    ' we need to know the trump suit. For the '
    ' test hand, the trump suit is indicated  '
    ' in the Source box to the right of the   '
    ' test hand, like "<H...".  For non-test  '
    ' looking up of bids, the suit is either  '
    ' passed by the calling program or this   '
    ' program asks for it below. Even without '
    ' the trump suit, some ace-asking checks  '
    ' can be done. '
    
    ' Find trump suit: '
    ref = Left$(rs_BidTable!Source & "  ", 2)
    int_TrumpSuit = (InStr("<C<D<H<S", ref) + 1) / 2
    If int_TrumpSuit = 0 Then
        t = UCase$(InputBox("Enter trump suit (C, D, H, or S)", "Entry #" & rs_BidTable!id & ".  Test Hand: " & rs_BidTable!Test_Hand))
        int_TrumpSuit = InStr("CDHS", t)
    End If
    
    bln_TrumpAce = bln_Suit(int_TrumpSuit, Ace)
    bln_TrumpKing = bln_Suit(int_TrumpSuit, King)
    
    
    ' If basic Grand Slam Force is being used: '
    ' Suits - Comb.-HCP should have "TA+TK,TA+TQ,TK+TQ". '
    ' This specifies 2 of the top 3 trumps. '
    
    Debug.Print "Entry ID# = "; rs_BidTable!id
    
    If InStr(pts1, "A") Then
        int_NumAces = 0
        For s1 = 1 To 4
            If bln_Suit(s1, Ace) Then
                int_NumAces = int_NumAces + 1
            End If
        Next
    End If
    
    
    'if bln_TrumpKing
    
    ' Count Kings: '
    i = InStr(pts1, "K")
    ' Two types of specs: "2K-TK" to exclude TK or omit '
    ' "-TK" for just "1K", "2K", "3K, etc., including TK.
    ' Don't think King asking is done for anything except
    ' after Ace asking. '
    If i > 1 Then
        If Val(Mid$(s, i - 1, 1)) > 0 Then
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
    
    ' Each spec separated by a comma has to be tested. '
    ' If one spec passes, then Combo is ok; otherwise, '
    ' try the next spec until done. '
    ' Example: pts1 = 2A+TQ,1A+TK+TQ,4A+TK+TQ  This is '
    ' for the bid 5S:  2 or 5 key cards plus trump Q.  '
    ' There are 3 specs: 2A+TQ - 1A+TK+TQ - 4A+TK+TQ.  '
    p = pts1 & ","  ' Example: 2A+TQ,1A+TK+TQ,4A+TK+TQ, '
    Do
        c = InStr(p, ",")   ' Example: c = 6 '
        s = Left$(p, c - 1) ' 1st spec from example is 2A+TQ. '
        If c = Len(p) Then
            p = ""
        Else
            p = Mid$(p, c + 1) ' 1A+TK+TQ,4A+TK+TQ, is left '
        End If
        
        If InStr(s, int_NumAces & "A") = 0 And InStr(s, "A") > 0 Then ' e.g.: "1A-TK" '
            GoTo NextLoop
        ElseIf InStr(s, "A") > 0 And _
               InStr(s, int_NumAces & "A") = 0 _
        Then
            ' Example: "4A+TK+TQ" and # Aces=2. '
            ' Actual # Aces not in the spec. '
            GoTo NextLoop
        End If
        
        ' Look for trump King in this spec: '
        i = InStr(s, "TK")
        If i > 1 Then
            signK = Mid$(s, i - 1, 1)
        End If
        If i > 0 Then
            If Not bln_Suit(int_TrumpSuit, King) Then
                ' TK specified but not held. '
                GoTo NextLoop
            ElseIf s = "TK" Or s = int_NumAces & "A+TK" Then
                comboOk = True
                GoTo ExitCombo
            End If
            ' We now know that
            bln_TrumpKing = True
        Else
        End If
        
        ' Look for trump Queen: '
        i = InStr(s, "TQ")
        If i > 1 Then
            signQ = Mid$(s, i - 1, 1)
            If signQ <> "+" And signQ <> "-" Then
                MsgBox "TQ does not have a + or - in front of it."
            End If
            If bln_Suit(int_TrumpSuit, Queen) Then
                bln_TrumpQueen = True
            End If
        End If
        
        ' Test for the current spec (example: 2A+TQ): '
        If bln_TrumpKing Then
            If bln_TrumpAce And s = "TA+TK" Then
                ' Grand Slam Force '
                comboOk = True
                GoTo ExitCombo
            ElseIf bln_TrumpQueen Then ' It has to be in the spec. '
                If Right$(s, 4) = "K+TQ" Then
                    ' s = "TK+TQ" or "#A+TK+TQ" '
                    comboOk = True
                    GoTo ExitCombo
                Else
                    ' s = "TK-TQ" or "#A+TK-TQ" '
                    ' not ok '
                End If
            ElseIf Right$(s, 4) = "K-TQ" Then
               ' s = "TK-TQ" or "#A+TK-TQ" '
                comboOk = True
                GoTo ExitCombo
            End If
        End If
        
NextLoop:
    Loop While p <> ""
    
SlamBiddingDone:

    s1 = InStr("CDHSx", Left$(pts1, 1))
    If Len(pts1) > 3 And s1 <> 0 Then
        ' Get x spec, if any: '
        If s1 = 5 Then
            s1 = InStr(" CDHS", str_xSuit) - 1
            If s1 < 1 Then s1 = 0
        End If

        sPts = 0
        ttl = 0
        s2 = 0
        s3 = 0

        If s1 > 0 Then
            s2 = InStr("CDHS", Mid$(pts1, 2, 1))
            s3 = InStr("CDHS", Mid$(pts1, 3, 1))
            If s3 > 0 Then ' CDH=10-12, DHS>9, etc.
                sign = Mid$(pts1, 4, 1)
                sPts = int_HCP(s1) + _
                       int_HCP(s2) + _
                       int_HCP(s3)
                ttl = Val(Mid$(pts1, 5))

            Else ' CD=8-9, SH>8, etc.
                sign = Mid$(pts1, 3, 1)
                sPts = int_HCP(s1) + _
                       int_HCP(s2)
                ttl = Val(Mid$(pts1, 4))
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
    Else ' No matches anywhere above should exit the function, so to get here, it must be ok. '
        comboOk = True
    End If

ExitCombo:
'''''''''
    If comboOk Then
        Combo_Ok = True

    ElseIf bln_AnalyzeBid Then
        With frm.lstAnalysis
            .ListIndex = 23
            ' This doesn't properly analyze why '
            ' Ace-asking didn't match. Fix it sometime. '
            'If InStr(pts1, "A") > 0 Then
            '    If InStr(pts1, "TK") > 0 Then
            '        ref = "# Keys = " & numKeys
            '    Else
            '        ref = "# Aces = " & int_NumAces
            '    End If
            'ElseIf InStr(pts1, "K") > 0 Then
            '    ref = "# Kings = " & int_NumKings
            'Else
            '    ref = "Suit quantity and HCPs = " & sTtl & ", " & sPts
            'End If
            ref = "    NO Match" ' & ":   " & ref
            .List(23) = .List(23) & _
            "" & _
            ref
        End With
    End If

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

    Dim qOk As Boolean ' Quantity passes? '
    Dim pOk As Boolean ' HCPs passes?     '

    qOk = True
    pOk = True
    bln_xField = False

    If num <> "" And Right$(num, 1) <> "x" Then
        ' Check suit quantity: '
        qOk = SuitQty_Ok(s, num)
    End If

    If HCPs <> "" And qOk And Right$(HCPs, 1) <> "x" Then
        ' Check suit HCPs:  '
        pOk = SuitHCP_Ok(s, HCPs)
    End If

    If qOk And pOk Then
        Suit_Ok = True
    End If

End Function


Private Function SuitQty_Ok( _
            ByVal s As Integer, _
            ByVal num As String) _
        As Boolean
    'Check number of cards in suit:

    Dim nm As Integer
    Dim nm2 As Integer
    Dim s2 As Integer
    Dim signR As String
    Dim signL As String
    Dim x As String
    Dim i As Integer

    nm = Val(num)
    signR = Right$(num, 1)
    signL = Left$(num, 1)

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
            x = Mid$(num, InStr(num, "-") + 1)
            i = InStr("CDHS", x)
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

End Function


Private Function SuitHCP_Ok( _
            ByVal s As Integer, _
            ByVal HCPs As String) _
        As Boolean
    ' Check strength of suit. '
    ' Format of HCPs can be:  '
    '   3 -- specific number  '
    '   3+                    '
    '   >2 -- same as 3+      '
    '   <4                    '
    '   2-4 -- range          '
    '   A -- Ace              '
    '   K/k -- King           '
    '   Q/q -- Queen          '
    '   A,K or A,K,Q or K,Q   '
    '      are also allowed.  '
    '   a -- AK, KQ, AQ       '
    '   b -- AJT, KJT, QJT    '
    '   c -- KQJ, AQJ, AKQ    '
    '   2stop -- 2+ stoppers  '
    '   1stop -- 1+ stopper(s)'
    '   hstop -- half+ stop(s)'
    '   0stop -- no stoppers  '
    '   FRC -- 1st rnd. ctrl. '
    '   SRC -- 2nd rnd. ctrl. '
    '   i2 -- 2 intermediates '
    '   4+i2 -- 4+ HCPs PLUS  '
    '        2 intermediates  '
    '   P -- Protected honors '
    ' In Comb. Points field:  '
    '   kc1,4 -- RCKB 1 or 4  '
    '   kc0,3 --      0 or 3  '
    '   kc2,5+Q with trump Q  '
    '   kc2,5-Q w/o trump Q   '

    Dim ok As Boolean
    Dim originalHCPs As String
    Dim ptsNeeded As Integer
    Dim left1 As String
    Dim left2 As String
    Dim left3 As String
    Dim iPtr As Integer
    Dim intermedReqd As Integer
    Dim honorsList As String
    Dim honCombo As String
    Dim signL As String
    Dim signR As String
    Dim extra As Boolean
    
    'Get rid of blanks:
    HCPs = Replace(HCPs, " ", "", , 99)

    If HCPs = "" Then
        SuitHCP_Ok = True
        Exit Function
    End If

    ok = False
    originalHCPs = HCPs
    HCPs = LCase$(HCPs)


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
    signL = Left$(HCPs, 1)

    If HCPs = "p" Then
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

        left1 = Trim$(Left$(str_Suit(s) & " ", 1))
        left2 = Trim$(Left$(str_Suit(s) & "  ", 2))
        left3 = Trim$(Left$(str_Suit(s) & "   ", 3))
        ok = True
        ' True except for the following:
        If left2 = "AQ" Or _
           left1 = "K" Or _
           left1 = "Q" _
        Then

            If left3 <> "KQJ" And _
               left3 <> "QJT" And _
               left3 <> "AQJ" _
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


    '  Check for Ace or King specified: '

    ElseIf InStr(originalHCPs, "A") > 0 Then
        ' Since case matters with "A"/"a", '
        ' and HCPs has been lowercased, we '
        ' must go back to originalHCPs to  '
        ' check for "A" or "A,K".          '
        
        ok = bln_Suit(s, Ace)

    ElseIf InStr(HCPs, "k") > 0 Then
        ' Case doesn't matter with "k".
        ok = bln_Suit(s, King)

    ElseIf InStr(HCPs, "q") > 0 Then
        ok = bln_Suit(s, Queen)

    '  Check for Controls specified: '

    ElseIf HCPs = "frc" Then
        'First Round Control (void or Ace)
        ok = int_SuitLen(s) = 0 Or _
              bln_Suit(s, Ace)

    ElseIf HCPs = "src" Then
        'Second Round Control (singleton or King)
        ok = int_SuitLen(s) = 1 Or _
              bln_Suit(s, King)


    '  Do the math:  '

    ElseIf signL = "<" Then ' e.g.: <5 '
        ok = int_HCP(s) < Val(Mid$(HCPs, 2))

    ElseIf Val(HCPs) > 0 Or HCPs = "0" Then
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

        Else ' No sign or range, so must be '
             ' a specific number, like "4". '
            ok = int_HCP(s) = ptsNeeded

        End If

    ElseIf Len(HCPs) > 0 Then
        ' At this point, HCPs must be a '
        ' code like a, b, c, ab, etc.   '
        ' Convert into actual cards:    '

        left1 = Left$(HCPs, 1)
        left2 = Left$(HCPs & " ", 2)

        If left2 = "ab" Then
            honorsList = "AK,AQ,KQ,AJT,KJT,QJT,AKQ"

        ElseIf left2 = "bc" Then
            honorsList = "AJT,KJT,QJT,KQJ,AQJ,AKQ,AKJ,AKQJ"

        ElseIf left1 = "a" Or _
               left2 = "ac" _
        Then ' "c" is redundant with "a" '
            honorsList = "AK,AQ,KQ"

        ElseIf left1 = "b" Then
            honorsList = "AJT,KJT,QJT"

        ElseIf left1 = "c" Then
            honorsList = "AKQ,AKJ,AQJ,KQJ"

        ElseIf left1 = "d" Then
            honorsList = "AKQJ,KQJT,AKQT,AKQJT"
        End If

        ' Work thru all honor combos in honorsList, '
        ' such as "AK,AQ,KQ" to see if the hand has '
        ' the cards specified.                      '
        Do While honorsList <> "" And ok = False
            ' Get the first combo, then '
            ' remove it from HCPs:      '
            iPtr = InStr(honorsList, ",")

            If iPtr > 0 Then ' e.g.:  AK,AQ,KQ '
                honCombo = Left$(honorsList, iPtr - 1) ' e.g.: AK '
                honorsList = Mid$(honorsList, iPtr + 1)   ' e.g.: AQ,KQ '
                If Left$(honCombo, 1) = "A" Then
                    extra = True
                End If
            Else
                honCombo = honorsList ' e.g.: KQ '
                honorsList = ""

            End If

            ok = (honCombo = Left$(str_Suit(s), Len(honCombo)))
                ' Note that if only "c" is specified '
                ' ("AJT,KJT,QJT") the entry will be  '
                ' be rejected if the Hand is better  '
                ' ("AQJT", "KQJT", etc.). If person  '
                ' creating BidBase entry wants to    '
                ' include better holdings, put "ab"  '
                ' or "bc".                           '

        Loop

    End If

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
            ' if K82 is okay.             '
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
        Stop
    End If


    SuitHCP_Ok = ok

End Function


Private Sub ShowAnalysis()
    Dim num As Long
    Dim sel As String
    Dim sel2 As String
    Dim nam As String

    bln_AnalyzeBid = True
    With frm.lstAnalysis
        Load frm ' frmBidBaseAnalysis
        frm.Show
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
        .AddItem "Bid Name:  " & nam
        .AddItem "Requirements for call--                (Hand being bid: " & str_PlayersHand & ")"
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
        .AddItem "Position:     " & rs_BidTable!POS
        .AddItem "Shape:        " & rs_BidTable!Shape
        .AddItem ""
        .AddItem "             Num    HCP"
        .AddItem "'x' :     " & _
                  Space$(7 - Len(rs_BidTable!xNum)) & rs_BidTable!xNum & _
                  Space$(8 - Len(rs_BidTable!xHCP)) & rs_BidTable!xHCP
        .AddItem "Spds:" & _
                  Space$(7 - Len(rs_BidTable!sNum)) & rs_BidTable!sNum & _
                  Space$(10 - Len(rs_BidTable!sHCP)) & rs_BidTable!sHCP
        .AddItem "Hrts:  " & _
                  Space$(7 - Len(rs_BidTable!hNum)) & rs_BidTable!hNum & _
                  Space$(10 - Len(rs_BidTable!hHCP)) & rs_BidTable!hHCP
        .AddItem "Dias: " & _
                  Space$(7 - Len(rs_BidTable!dNum)) & rs_BidTable!dNum & _
                  Space$(10 - Len(rs_BidTable!dHCP)) & rs_BidTable!dHCP
        .AddItem "Clbs: " & _
                  Space$(7 - Len(rs_BidTable!cNum)) & rs_BidTable!cNum & _
                  Space$(10 - Len(rs_BidTable!cHCP)) & rs_BidTable!cHCP
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

    With rs_BidTable
        If lng_TestID > 0 And !id <> lng_TestID Then
            m = "Test Hand for Bid Order Number " & lng_TestID & vbCr & "also matches the specs for entry ID#" & !id & vbCr
            If Trim(!prior_bids) <> "" Then
                m = m & "Prior bids: " & !prior_bids & vbCr
            End If
            m = m & "Entry #" & !Bid_Order_Number & " ID#:" & !id
            m = m & "      Pct: " & !Pct_Used
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
            m = m & vbCr & vbCr + "Keep testing?"
            ans = MsgBox(m, vbYesNo)
            If ans = vbNo Then
                int_ErrCode = 98
                str_ErrMsg = "Search cancelled by user."
            ElseIf ans = vbYes Then
                str_ErrMsg = "Keep testing"
            End If

        Else
            If InStr(!Bid, "x") > 0 Then
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
End Function


Private Sub GetDiscl()
    ' Loop thru all entries matching Prior Calls '
    ' and selected bid to get their disclosures  '
    ' to give to client. (As a reminder: you can '
    ' not just give the user the disclosure for  '
    ' the selected entry because that gives him  '
    ' more specific information than to which he '
    ' is entitled. For example, if there are two '
    ' entries for "1H-P-4H" where one is weak and'
    ' preemptive and the other is expecting to   '
    ' make, the user is not entitled to know     '
    ' which bid was chosen and why, just that it '
    ' could be either-or.                        '
    
    ' 2081-08-06: There is a disclosure field in '
    ' the database which is supposed to be filled'
    ' out for each entry made. Why do this too?  '

    Dim s As String

    If str_Call(int_CallNum) = "Pass" Then
        Exit Sub
    End If

    ' Get the first entry matching the Prior Bids: '
    s = "SELECT * FROM Bids"
    s = s + " WHERE Prior_Bids='" & str_PriorBids & "'"
    s = s + " ORDER BY Bid_Order_Number;"
    Set rs_BidTable = db_BidBase.OpenRecordset(s, dbOpenSnapshot)

    ' Look for entries matching the bid made: '

    Do While rs_BidTable!Bid_Order_Number <> "ZZZZZZ" And _
             rs_BidTable!prior_bids = str_PriorBids

        If rs_BidTable!Bid = str_Call(int_CallNum) And _
           (rs_BidTable!Pct_Used = "" Or _
            Val(rs_BidTable!Pct_Used) <> 0) _
        Then
            AddToDisclosure
        End If

        ' Get next record, if any: '
        If Not rs_BidTable.EOF Then
            rs_BidTable.MoveNext
            If rs_BidTable.EOF Then
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop

End Sub


Private Sub AddToDisclosure()

    Dim s As String
    Dim a As String
    Dim i As Integer
    Dim nb As String
    Dim nn As Integer
    Dim num As Long
    Dim sel As String
    Dim sel2 As String
    Dim nam As String

With rs_BidTable
    num = Val("" & !ConvID)
    LookUpBidName num, sel, sel2, nam
    If Trim(nam) <> "" Then
        If sel <> "" And sel <> "0" Then
            nam = nam & "  Cvn: " & sel & "  Sub: " & sel2
        End If
    End If
    
    If sel = "0" Or sel2 = "0" Then
        ' Bid is deselected. '
        ' This should have been picked up before now. '
        Exit Sub
    End If
    
    str_Disclosure = str_Disclosure & !Bid & ":  BON:" & !Bid_Order_Number & ". "
    If nam <> "" Then
      str_Disclosure = str_Disclosure & Chr$(34) & nam & Chr$(34) & ". "
    End If
    If "" & !POS <> "" Then str_Disclosure = str_Disclosure & "Position: " & !POS & ". "
    
    ' Strength:                       '
    If !Strength_Discl = "V" Then
        s = "Very Weak"
    ElseIf !Strength_Discl = "W" Then
        s = "Weak"
    ElseIf !Strength_Discl = "I" Then
        s = "Intermediate"
    ElseIf !Strength_Discl = "O" Then
        s = "Opening"
    ElseIf !Strength_Discl = "S" Then
        s = "Strong"
    ElseIf !Strength_Discl = "WS" Then
        s = "Weak-Strong"
    ElseIf !Strength_Discl = "WI" Then
        s = "Weak-Intermediate"
    ElseIf !Strength_Discl = "IS" Or _
           !Strength_Discl = "I+" _
    Then
        s = "Intermediate-Strong"
    ElseIf !Strength_Discl = "W+" Then
        s = "Weak-up"
    ElseIf !Strength_Discl = "S+" Then
        s = "Strong+"
    ElseIf !Strength_Discl = "VS" Then
        s = "Very Strong"
    End If

    str_Disclosure = _
        str_Disclosure & vbCrLf & _
        "Strength=" & s & ".  "

    ' Suit lengths & HCPs:               '
    s = ".  x:" & !xHCP & " S:" & !sHCP & " H:" & !hHCP & " D:" & !dHCP & " C:" & !cHCP
    str_Disclosure = str_Disclosure & s & vbCrLf
    
    str_Disclosure = str_Disclosure & "Shape: " & !Shape & ".  "
    s = "x:" & !xLen & " S:" & !Slen & " H:" & !Hlen & " D:" & !Dlen & " C:" & !Clen
    If Len(s) = 14 Then
        ' No suit length disclosures; give bid suit length requirements: '
        s = "x:" & !xNum & " S:" & !sNum & " H:" & !hNum & " D:" & !dNum & " C:" & !cNum
    End If
    str_Disclosure = str_Disclosure & s & vbCrLf
    
    ' Asks pard for/to bid:             '
    a = "" & LCase$(!Asks)

    If a = "" Then
        'skip it'
    ElseIf a = "g" Then
        nb = "game force"
    ElseIf a = "1" Then
        nb = "bid again"
    ElseIf a = "poc" Then
        nb = "pass or correct"
    ElseIf a = "p" Then
        nb = "pick a suit/slam"
    ElseIf a = "a" Then
        nb = "show Aces"
    ElseIf a = "k" Then
        nb = "show Kings"
    ElseIf a = "kc" Then
        nb = "show # of Key-Cards"
    ElseIf Left$(a, 1) = "h" Then
        s = Mid$(a, 2, 1)
        i = InStr("cdhs", s)
        s = str_SuitNames(i)
        nb = "for help in " & s
    ElseIf Left$(a, 1) = "c" Then
        s = Mid$(a, 2, 1)
        i = InStr("cdhs", s)
        s = str_SuitNames(i)
        nb = "for controls in " & s
    ElseIf Left$(LCase$(a), 1) = "b" Then
        s = LCase$(Mid$(a + " ", 2, 1))
        Select Case Mid$(a + "  ", 3, 1)
            Case "m"
                nb = "a minor"
            Case "M"
                nb = "a major"
            Case Else
                i = InStr("cdhs", s)
                s = str_SuitNames(i)
        End Select
        nb = "to bid " & nb & s
    ElseIf Trim$(!Asks) <> "" Then
        nb = "to " & !Asks
    End If
    If nb <> "" Then
        str_Disclosure = _
            str_Disclosure & _
            "Asks partner " & nb & "."
    End If
    ' Pct Used:                         '
    If Val(!Pct_Used) > 0 Then
      str_Disclosure = str_Disclosure & _
        "   Pct Used: " & !Pct_Used & "%" & vbCrLf
    End If

    ' Explanation:                      '
    a = ""
    'nb = "Notes: "
    If !Alert <> "D" And !Explanation <> "" Then
        a = Trim$(!Explanation)
        nn = InStr(a, "[")
        i = InStr(a, "]")
        If nn > 0 Then
            If i > nn Then
                If nn > 1 Then
                    ' "Show user this. [Not this.] Show this"
                    a = Trim$(Left$(a, nn - 1) _
                        & Mid$(a, i + 1))
                Else ' "[Hide this] Show user this."
                    a = Trim$(Mid$(a, i + 1))
                End If
            Else ' "[Don't show user any of this"
                a = ""
            End If
        End If
        If !Alert = "A" Then
            nb = "ALERT: "
        Else
            nb = "NOTE:  "
        End If
    End If
    If a <> "" Then 'a = "(none)"
        str_Disclosure = str_Disclosure _
                         & nb & a & vbCrLf
    End If
    
    ' Sample/test hand:                 '
    nb = Trim$(!Source)
    If nb <> "" Then ' nb = "(none)"
        str_Disclosure = str_Disclosure _
            & "Source: " & "'" & nb & "'"
    End If
    str_Disclosure = str_Disclosure & vbCrLf & vbCrLf
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
                    Stop
                    ' Get real suits.
                End If
            End If
        End If
        n = n + 4
    Loop

    ' Get RHO's suits: '
    j = 1
    n = int_RHOs_Pos
    Do While j < 3 And _
             str_Bid(n) <> "" And _
             str_Bid(n) <> "P"

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
                    Stop
                    ' Get real suits.
                End If
            End If
        End If
        n = n + 4
    Loop

    ' Get Pard's suits: '
    j = 1
    k = 1
    n = int_Pards_Pos
    Do While j < 3 And _
             str_Bid(n) <> "" And _
             str_Bid(n) <> "P"

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
                            int_Pards_SuitLen(j) = Val("" & !cNum)

                        ElseIf int_Pards_Suit(j) = Dias Then
                            int_Pards_SuitLen(j) = Val("" & !dNum)

                        ElseIf int_Pards_Suit(j) = Hrts Then
                            int_Pards_SuitLen(j) = Val("" & !hNum)

                        ElseIf int_Pards_Suit(j) = Spds Then
                            int_Pards_SuitLen(j) = Val("" & !sNum)
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
                    q(1) = Val("" & !cNum)
                    q(2) = Val("" & !dNum)
                    q(3) = Val("" & !hNum)
                    q(4) = Val("" & !sNum)
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
                    For i = 1 To 4
                        If q(z(i)) > 3 Then

                            Stop ' check the following:
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
    Loop

End With

    ' Get current player's suits: '
    m = int_BidNum Mod 4
    If m = 0 Then m = 4
    ' First bid: '
    k = 1
    c = " CDHS"
    Do While k < 4 And m < int_BidNum
        s = Right$(" " & str_Bid(m), 1)
        n = InStr(c, s) - 1
        If n > 0 Then
            Mid$(c, n + 1) = " " ' Don't use same suit twice. '
            j = 1
            Do While int_Pards_Suit(j) > 0
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
            Loop
            If int_SuitLen(n) > 3 Then
                int_My_Suit(k) = n
            End If
        End If
        k = k + 1
        m = m + 4
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
    If Right$(!Bid, 1) <> "x" Then
        xSuitEqualSuitBid = True
        Exit Function
    End If

    Do
        Select Case b
        Case Is = "C"
            xSuitEqualSuitBid = (Left$(!cNum & " ", 1) = "x")
            Exit Function
        Case Is = "D"
            xSuitEqualSuitBid = (Left$(!dNum & " ", 1) = "x")
            Exit Function
        Case Is = "H"
            xSuitEqualSuitBid = (Left$(!hNum & " ", 1) = "x")
            Exit Function
        Case Is = "S"
            xSuitEqualSuitBid = (Left$(!sNum & " ", 1) = "x")
            Exit Function
        End Select
        If .EOF Then Exit Function
        .MoveNext
    Loop While Right$(!Bid, 1) = "x"
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
    Dim qSel As String
    Dim bx As String

    On Error GoTo GetOPsBid_Error

    If Not bln_AlreadyInitialized Then
        InitializeBidBaseDLL
    End If

    If InStr(" CDHS", Right$(" " & b, 1)) > 1 Then
        bx = Left$(b, 1) & "x"
        ' e.g.: If b=1S, bx=1x     '
        ' Have to search for both. '
    End If

    qSel = "SELECT BidNames.Selected, BidNames.SubCat_Selected, Bids.Pct_Used, Bids.RoundOfBidding, Bids.Player, Bids.PB_Sort_Order, Bids.Prior_Bids, Bids.Bid_Order_Number, Bids.Bid, BidNames.Bid_Name, UseAgainst.Bid_Name, Bids.HCP, Bids.Dist_Pts, Bids.ValSystem, Bids.SuitVal, Bids.NT_Val, Bids.Cust_Pts, Bids.A_Cust, Bids.K_Cust, Bids.Q_Cust, Bids.J_Cust, Bids.T_Cust, Bids.Winrs, Bids.Losers, Bids.QTs, Bids.Stopped_Suits, Bids.Shape, Bids.Intermed, Bids.xNum, Bids.xHCP, Bids.sNum, Bids.sHCP, Bids.hNum, Bids.hHCP, Bids.dNum, Bids.dHCP, Bids.cNum, Bids.cHCP, Bids.cmNum, Bids.cmHCP, Bids.Vul, Bids.Pos, Bids.Artificial, Bids.Alert, Bids.Explanation, Bids.Test_Hand, Bids.Strength_Discl, Bids.xLen, Bids.SLen, Bids.HLen, Bids.DLen, Bids.CLen, Bids.Asks, Bids.DateOfEntry, Bids.Best_Contract, Bids.Num_Used, Bids.Source, Bids.ID FROM BidNames RIGHT JOIN (BidNames AS UseAgainst RIGHT JOIN Bids ON UseAgainst.IDNum = Bids.AgainstID) ON BidNames.IDnum = Bids.ConvID "


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

    Set rs_BidsQuery = db_BidBase.OpenRecordset(qSel & s, dbOpenDynaset) ' dbOpenForwardOnly)

Exit_GetOPsBid:
   On Error GoTo 0
   Exit Sub

GetOPsBid_Error:
'''''''''''''''
    MsgBox Err.Description & " getting entries for " & vbCr & _
           "Prior Bids of " & pb & vbCr & _
           "and Current Bid of " & b
    Resume Exit_GetOPsBid

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
    With rs_HandVal
        str_ValSpec = Trim$("" & !Value)
        sng_ValSpec = Val(str_ValSpec)
        str_FlagSpec = Trim$("" & !Flag)
        str_PosSpec = Trim$("" & !position)
        str_HandSpec = Trim$("" & !Hands)
        str_FitSpec = Trim$("" & !Fit)
        str_SuitSpec = Trim$("" & !suit)
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
            Stop
            'Eval_HandsSpecCalcOk = (t_Str(i) <> Right$(s, 1)) Or t_Str(i - 1) = ""
        Else
            j = InStr("=<>", Mid$(s, 2, 1))
            If i > 0 And j > 0 And k > 0 Then
                i = i - 1
                Select Case Mid$(s, 2, 1)
                Case Is = "="
                    Stop
                    'Eval_HandsSpecCalcOk = (t_Str(i) = Right$(s, 1))
                Case Is = "-"
                    Stop
                    'Eval_HandsSpecCalcOk = (t_Str(i) <> Right$(s, 1)) Or t_Str(i) = ""
                Case Is = "<"
                    Stop
                    'Eval_HandsSpecCalcOk = (InStr("BWIOPSVX", t_Str(i)) < _
                                            InStr("BWIOPSVX", Right$(s, 1))) Or t_Str(i) = ""
                Case Is = ">"
                    Stop
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
            Stop
            'Eval_HandsSpecCalcOk = (Val(str_LHOs_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: L<20 = LHO has <15 VPs '
            Stop
            'Eval_HandsSpecCalcOk = (Val(str_LHOs_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If

        If Len(s) = 2 Then ' LN, LS, LC, LR, LF -- alone. '
            Select Case r
            Case Is = "C"
                Stop
                'Eval_HandsSpecCalcOk = (int_LHOs_CueBid > 0) Or neg
            Case Is = "F"
                Stop
                'Eval_HandsSpecCalcOk = (int_LHOs_Fit(1) > 0) Or neg
            Case Is = "N"
                Stop
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
                    Stop
                    'Eval_HandsSpecCalcOk = (int_LHOs_CueBid = j)
                Case Is = "LF"
                    Stop
                    'Eval_HandsSpecCalcOk = (int_LHOs_Fit(1) = j Or _
                                            int_LHOs_Fit(2) = j)
                Case Is = "LH"
                    Stop
                    'Eval_HandsSpecCalcOk = (int_LHOs_HelpSuit = j)
                Case Is = "LR"
                    Stop
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
            Stop
            'Eval_HandsSpecCalcOk = (Val(str_RHOs_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: R<20 = RHO has <15 VPs '
            Stop
            'Eval_HandsSpecCalcOk = (Val(str_RHOs_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If

        If Len(s) = 2 Then ' RN, RS, RC, RR, RF -- alone. '
            Select Case r
            Case Is = "C"
                Stop
                'Eval_HandsSpecCalcOk = (int_RHOs_CueBid > 0) Or neg
            Case Is = "F"
                Stop
                'Eval_HandsSpecCalcOk = (int_RHOs_Fit(1) > 0) Or neg
            Case Is = "N"
                Stop
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
                    Stop
                    'Eval_HandsSpecCalcOk = (int_RHOs_CueBid = j)
                Case Is = "RF"
                    Stop
                    'Eval_HandsSpecCalcOk = (int_RHOs_Fit(1) = j Or _
                                            int_RHOs_Fit(2) = j)
                Case Is = "RH"
                    Stop
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
            Stop
            'Eval_HandsSpecCalcOk = (Val(str_Pards_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: P<20 = Pard has <15 VPs '
            Stop
            'Eval_HandsSpecCalcOk = (Val(str_Pards_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If

        If Len(r) = 1 Then
            Select Case r  ' PC, PF, PN, PR, PS -- alone '
            Case Is = "C"
                Eval_HandsSpecCalcOk = (int_Pards_CueBid > 0) Or neg
            Case Is = "F"
                Stop
                'Eval_HandsSpecCalcOk = (int_Pards_Fit(1) > 0) Or neg
            Case Is = "N"
                Stop
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
                    Stop
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
            Stop
            'Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = 0 Or _
                            (int_Pards_Suit(1) <> int_LHOs_CueBid And _
                             int_Pards_Suit(2) <> int_LHOs_CueBid))
        Case Is = "P<>RC"
            Stop
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
            Stop
            'Eval_HandsSpecCalcOk = _
                            (int_Pards_Suit(1) = int_LHOs_CueBid And _
                             int_Pards_Suit(2) = int_LHOs_CueBid)
        Case Is = "P=RC"
            Stop
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
                Stop
                'Eval_HandsSpecCalcOk = (int_My_CueBid > 0) Or neg
            Case Is = "F"
                Stop
                'Eval_HandsSpecCalcOk = (int_My_Fit(1) > 0) Or neg
            Case Is = "N"
                Stop
                'Eval_HandsSpecCalcOk = bln_I_Bid_NT Or neg
            Case Is = "R"
                Stop
                Eval_HandsSpecCalcOk = (int_My_RebidSuit(1) > 0) Or neg
            Case Is = "S"
                Stop
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
            Stop
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
           Trim$("" & rs_HandVal!Cards) = "" _
        Then
            If Right$("  " & str_ValSpec, 2) = "*n" Then
                n = int_SuitLen(int_SuitNum)
                If rs_HandVal!Fit = "Y" Or _
                    rs_HandVal!suit = "F" _
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
    Dim x As String

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
        x = Left$(str_CardsSpecExtracted, 1)
        If x = "?" Then
            str_CardsSpecOperator = x
            x = "-"
        End If
    Else
        x = ""
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
            x = Mid$(str_Suit(int_SuitNum), j, 1)
            If InStr(" AKQJT", x) > 1 Then
                Eval_AddVal
            End If

        ElseIf Mid$(str_Suit(int_SuitNum), _
                j + 1, _
                Len(str_CardsSpecExtracted) - j + 1) _
           = Mid$(str_CardsSpecExtracted, j) _
        Then
            If x <> "-" Then
                Eval_AddVal
            End If

        ElseIf x = "-" Then
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
    ' This is used in the Bidder DLL '
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
            ok = Eval_PtsOk(sng_ValPoints(i, 1), vs, 5)
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

    intMin = specs
    intDash = InStr(specs, "-")

    ' Compute min and max for test: '

    If Left$(specs, 1) = "<" Then 'e.g.: <10
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

    ElseIf bln_AnalyzeBid Then
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
    Dim x As Boolean
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




