Attribute VB_Name = "DealerRoutines"
Option Explicit

' This module contains routines for '
' shuffling, dealing, and sorting 4 '
' hands of 13 cards each.           '

Public bln_Abort As Boolean
Public bln_AbortSim As Boolean
Public int_Cards(52) As Integer ' The shuffled cards. '
Public int_Dealer As Integer
Public str_Dealer(4) As String
Public int_DealerSpecified As Integer
Public lng_DealNum As Long
Public lng_DealNumSpecified As Long
Public str_Hands(4) As String ' Hands in xx-xxx-xxx format
Public bln_HideHands As Boolean
Public bln_Rebid As Boolean
Public bln_SameDeal As Boolean
Public int_Suit(4, 4, 13) As Integer ' player, suit, card '
Public int_SuitQty(5, 4) As Integer ' player, suit '
Public str_Suits(4) As String ' Suit names - Spades, Hearts, etc.
Public int_Turn As Integer ' Player whose turn it is to bid or to whom cards are being assigned. '
Public bln_SelectedCards(52) As Boolean
Public int_HandQty(4)
Public int_Hand(4, 13) ' The Qcards.DLL # for each card in each player's hand. '
Public bln_PickedHand(4) As Boolean
Public int_NumPickedHands As Integer

' The following are variables used by Qcards32.DLL  '
' Most of them are not used in the bidding program, '
' but might be used in a full bridge-playing program.

Global CARDWIDTH As Long
Global CARDHEIGHT As Long
' Constants: '
Global Const Clb = 1
Global Const Dia = 2
Global Const Hrt = 3
Global Const Spd = 4

Global Const Ace = 1
Global Const King = 13

Global Const FACEDOWN = 0
Global Const FACEUP = 1
Global Const CARDWID = 71
Global Const CARDHT = 96
Global Const OFFSET = 16
' Card Backs: '
Global Const CrossHatch = 53
Global Const Plaid = 54
Global Const Weave = 55
Global Const Robot = 56
Global Const Roses = 57
Global Const IvyBlack = 58
Global Const IvyBlue = 59
Global Const FishCyan = 60
Global Const FishBlue = 61
Global Const ShellDesign = 62
Global Const Castle = 63
Global Const Beach = 64
Global Const CardHand = 65
' I don't know where the above came from, '
' but here is what I'm getting now: '
' 1 = big bubbles
' 2 = dark blue
' 3 = red
' 4 = blue sky with cloud
' 5 = light & dark purple diamonds
' 6 = MIDI keyboard

' Required in Form Load to use dll
Declare Function InitializeDeck Lib "qcard32.dll" (ByVal hwnd As Long) As Long
' Reset all card values to defaults
Declare Sub SetDefaultValues Lib "qcard32.dll" ()
' Set the currently used card back design for cards 105 to 109
Declare Sub SetCurrentBack Lib "qcard32.dll" (ByVal nIndex As Long)

' Card drawing subs                Card control=hwnd      Card value=1-52    Position: left=x          top=y        '
Declare Sub DrawSymbol Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nValue As Long, ByVal x As Long, ByVal Y As Long)
Declare Sub DrawCard Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nCard As Long, ByVal x As Long, ByVal Y As Long)
Declare Sub DrawBack Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nValue As Long, ByVal x As Long, ByVal Y As Long)
Declare Sub DealCard Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nCard As Long, ByVal x As Long, ByVal Y As Long)
Declare Sub RemoveCard Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nCard As Long)

' Get card information functions
Declare Function GetCardColor Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardSuit Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardValue Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardStatus Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardBlocked Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function IsCardDisabled Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardX Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardY Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetUser1 Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetUser2 Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetUser3 Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetUser4 Lib "qcard32.dll" (ByVal nCard As Long) As Long

' Set card information subs
Declare Sub SetCardStatus Lib "qcard32.dll" (ByVal nCard As Long, ByVal bValue As Long)
Declare Sub AdjustCardBlocked Lib "qcard32.dll" (ByVal nCard As Long, ByVal bValue As Long)
Declare Sub SetCardDisabled Lib "qcard32.dll" (ByVal nCard As Long, ByVal bValue As Long)
Declare Sub SetCardX Lib "qcard32.dll" (ByVal nCard As Long, ByVal x As Long)
Declare Sub SetCardY Lib "qcard32.dll" (ByVal nCard As Long, ByVal Y As Long)
Declare Sub SetUser1 Lib "qcard32.dll" (ByVal nCard As Long, ByVal bValue As Long)
Declare Sub SetUser2 Lib "qcard32.dll" (ByVal nCard As Long, ByVal nValue As Long)
Declare Sub SetUser3 Lib "qcard32.dll" (ByVal nCard As Long, ByVal nValue As Long)
Declare Sub SetUser4 Lib "qcard32.dll" (ByVal nCard As Long, ByVal nValue As Long)
Declare Sub SetOffSet Lib "qcard32.dll" (ByVal nValue As Long)

' Dragging subs and functions
Declare Function InitDrag Lib "qcard32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Sub AbortDrag Lib "qcard32.dll" ()
Declare Sub DoDrag Lib "qcard32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long)
Declare Function EndDrag Lib "qcard32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Sub ReturnDrag Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nCard As Long, ByVal nOldX As Long, ByVal nOldY As Long)
Declare Sub BlockDrag Lib "qcard32.dll" (ByVal hwnd As Long, nFirst As Long, ByVal nNumCards As Long, ByVal x As Long, ByVal Y As Long)
Declare Function EndBlockDrag Lib "qcard32.dll" (ByVal hwnd As Long, nFirst As Long, ByVal nNumCards As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Sub ReturnBlockDrag Lib "qcard32.dll" (ByVal hwnd As Long, nFirst As Long, ByVal nNumCards As Long, ByVal x As Long, ByVal Y As Long)
Declare Function GetFreeDestination Lib "qcard32.dll" (ByVal nSource As Long) As Long

' Undocumented functions, generally not used

' returns the number of any unblocked card which lies beneath the mouse coordinates x, y
Declare Function PointInFreeCard Lib "qcard32.dll" (ByVal x As Long, ByVal Y As Long) As Long

' returns the number of any card whose top 16 (or OffSet) pixels lie beneath the mouse coordinates x, y
Declare Function PointInCardTop Lib "qcard32.dll" (ByVal x As Long, ByVal Y As Long) As Long

' manually sets the active drag card for a subsequent DoDrag or BlockDrag call
Declare Sub SetActiveCard Lib "qcard32.dll" (ByVal nCard As Long)


Public Sub Deal()
  ' int_Cards() is filled with the numbers 1-52 to   '
  ' represent the 52 different cards.  For purposes  '
  ' of shuffling and dealing, it doesn't matter what '
  ' playing cards are assigned to each value from    '
  ' 1-52, but to keep things simple, 1-13 are Clubs, '
  ' 14-26 are Diamonds, 27-39 are Hearts, and 40-52  '
  ' are Spades.                                      '
  '                                                  '
  ' In Qcards32.DLL, which is used to display the    '
  ' cards on screen, 1=Ace, 2='2', 3='3', ..., 11=J, '
  ' 12=Q, 13=K.  The display values are stored in the'
  ' array int_Suit(h,s,c).                           '
  
  ShuffleCards
  If bln_Abort Then
      Exit Sub
  End If
  SortHands
  FillSuits
End Sub


Public Sub DisplayHand(ByVal h)
    ' In the PlayScr form, each hand is displayed  '
    ' in a picture control ("p_Hand(h)") where the '
    ' top and left of each card relates to the     '
    ' inside of the picture control. That is, for  '
    ' the first card, both top and left edges start'
    ' with a value of zero.                        '
    '                                              '
    ' For North and South, the left edge of each   '
    ' card must be incremented from the prior card '
    ' in order to fan across the screen.           '
    '                                              '
    ' For East and West, the top edge of each card '
    ' is incremented from the prior card in order  '
    ' to fan down the screen.                      '
    '                                              '
    ' h in the Sub name above is the hand number   '
    ' where 1=S, 2=W, 3=N, 4=E. '
  
    Dim s As Integer
    Dim s1 As Integer
    Dim c As Integer
    Dim leftc As Integer
    Dim topc As Integer
  
    PlayScr.p_Hand(h).Cls
    PlayScr.p_Hand(h).Refresh

    For s = Spd To Clb Step -1
        'Show Clubs before Diamonds for color contrast:
        If s = Dia Then
            s1 = Clb
        ElseIf s = Clb Then
            s1 = Dia
        Else
            s1 = s
        End If
    
        If int_Suit(h, s1, Ace) > 0 Then
            If bln_HideHands And h > 1 Then
                DrawBack PlayScr.p_Hand(h).hwnd, _
                  3, leftc, topc
            Else
                DrawCard PlayScr.p_Hand(h).hwnd, _
                  int_Suit(h, s1, Ace), _
                  leftc, topc
            End If
            If h = 1 Or h = 3 Then
                leftc = leftc + 20
            Else
                topc = topc + 27
            End If
        End If
    
        For c = 13 To 2 Step -1
            If int_Suit(h, s1, c) > 0 Then
                If bln_HideHands And h > 1 Then
                    DrawBack PlayScr.p_Hand(h).hwnd, _
                      3, leftc, topc
                Else
                    DrawCard PlayScr.p_Hand(h).hwnd, _
                      int_Suit(h, s1, c), _
                      leftc, topc
                End If
                If h = 1 Or h = 3 Then
                    leftc = leftc + 20
                Else
                    topc = topc + 27
                End If
            End If
        Next
    Next
End Sub


Public Sub LayOutDeck()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    ' There is a row of cards for each suit. '
    ' The Card controls are numbered 1 - 52. '
    ' qCard.DLL numbers Aces below 2's then  '
    ' then 2-K are numbered consecutively.   '
    ' We want to show the cards as AKQJT...  '
    ' so the Card controls are laid out in   '
    ' the following Index order:             '
    ' Clb: 1(A) 13(K) 12(Q) 11(J) 10 ...  2  '
    ' Dia: 14   26(K) 25(Q) 24(J) 23 ... 15  '
    ' Hrt: 27   39(K) 38(Q) 37(J) 36 ... 28  '
    ' SPd: 40   52(K) 51(Q) 50(J) 49 ... 41  '
    ' Since the loop to put the card faces in'
    ' the card controls goes from 1 to 52,   '
    ' the card controls are not done from    '
    ' left to right but in Index order.      '
     
    PlayScr.b_CancelHands.Visible = True
    For i = 1 To 52
        j = DoEvents() ' Cards wouldn't show otherwise, and it has to be before DrawCard. '
        DrawCard PlayScr.Card(i).hwnd, i, 0, 0
    Next
    
    PlayScr.L_EnterHandsHelp.FontSize = 12
    PlayScr.L_EnterHandsHelp = vbCrLf & _
                               Space$(7) & "Click on N, S, E, or W, then" & Chr$(13) & _
                               Space$(7) & "click on cards for that hand."
End Sub


Public Sub FillSuits()
    ' This routine fills the int_Suit(4, 13) '
    '    array from the int_Cards(52) array. '
    ' The int_Cards() array contains all the '
    '    cards shuffled for the current deal.'
    ' The int_Suit array contains all the    '
    '    cards for each player arranged by   '
    '    suit and card.                      '
    
    If bln_SameDeal Then Exit Sub
    
    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim c As Integer
    Dim cv As Integer
    Dim s As Integer
    Dim hStart As Integer
    
    Erase int_SuitQty()
    Erase int_Suit()
  
    ' In the card display DLL, Ace=1 and '
    ' K=13. So in the int_Suit() array,  '
    ' Ace is element 1, 2=2, ..., K=13.  '
    
    For h = 1 To 4
        hStart = (h - 1) * 13
        For j = 1 To 13
            c = int_Cards(hStart + j)
            s = SuitOf(c)
            cv = c Mod 13
            If cv = 0 Then cv = 13
      
            If cv = Ace Then
                int_Suit(h, s, Ace) = c
            Else
                int_Suit(h, s, cv) = c
            End If
            
            int_SuitQty(h, s) = int_SuitQty(h, s) + 1
        Next
    Next
    
    ' Set up arrays in "xxx-xxx-xxx-xxxx" format: '
    For h = 1 To 4
        str_Hands(h) = ""
        For s = Spd To Clb Step -1
            If int_Suit(h, s, Ace) > 0 Then
                str_Hands(h) = _
                    str_Hands(h) & _
                    "A"
            End If
      
            For c = King To 2 Step -1
                If int_Suit(h, s, c) > 0 Then
                    str_Hands(h) = _
                        str_Hands(h) & _
                        Mid$("A23456789TJQK", c, 1)
                End If
            Next
            
            If s > 1 Then
                str_Hands(h) = str_Hands(h) + "-"
            End If
        Next
    Next

End Sub


Sub GetDealNum()
    Dim ans As String
    Dim dln As Long
    
    'Get/set a deal number:
    Do
        PlayScr.L_InboxLbl = "Enter a deal number or 'user' to enter hands:"
        PlayScr.b_DoneDealNum.Default = True
        PlayScr.t_Inbox = "" & lng_DealNum
        PlayScr.f_Inbox.Visible = True
        PlayScr.t_Inbox.SetFocus
        Do While PlayScr.f_Inbox.Visible
            dln = DoEvents()
        Loop
        
        ans = "" & Str$(lng_DealNum) ' Had +1. Don't know why. '
        
        If LCase$(ans) = "user" Then
            ' User enters hand(s). '
            PlayScr.mEnterHands_Click
            bln_Abort = True
            Exit Sub
   
        ElseIf lng_DealNum = 0 Then
            bln_Abort = True
            Exit Sub
   
        ElseIf lng_DealNum > 9999999 Then
            MsgBox "Deal number cannot be greater than 9,999,999."
            PlayScr.f_Inbox.Visible = True
        End If
    Loop While lng_DealNum > 9999999
  
    'lng_DealNum = lng_DealNum + 1  ' The desired number was entered; why increment it? '

End Sub


Public Sub ShuffleCards()
    Dim dln As Long
    Dim ans As String
    Dim z As Long
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim tempi As Integer
        
    z = Rnd(-lng_DealNum)
    ' z is not used anywhere, but the call '
    ' above causes the same deal to occur  '
    ' every time for the specified deal #. '
    ' See "Rnd" in the VB6 help file.      '
  
    If Val(PlayScr.L_Deal.Caption) = lng_DealNum Then
        bln_SameDeal = True
        Exit Sub
    End If
  
    PlayScr.L_Deal.Caption = "#" & lng_DealNum
    PlayScr.L_Deal.Refresh

    ' Create an array of card numbers: '
    For i = 1 To 52
        int_Cards(i) = i
    Next i
    
    ' Scramble ("shuffle") the array:  '
    For j = 1 To 10 + Int(5 * Rnd) ' # of times to shuffle '
        For i = 1 To 52
            ' Get a random array # '
            k = Int(52 * Rnd + 1)
            ' Swap with current array #: '
            tempi = int_Cards(i)
            int_Cards(i) = int_Cards(k)
            int_Cards(k) = tempi
        Next
    Next
  
End Sub


Sub SortHands()
    '*****************************************'
    ' The Shuffle routine created an array in '
    ' int_Cards of the numbers 1-52 in random '
    ' order, such as 4-33-17-6-12, etc.       '
    '                                         '
    ' Now the 52 "cards" must be split into 4 '
    ' sets of 13; e.g.: 1-13, 14-26, 27-39,   '
    ' and 40-52, and then the cards within    '
    ' those sets must be sorted by number.    '
    '*****************************************'
    
    Dim h As Integer
    Dim ply As Integer
    Dim cStart As Integer
    Dim cStop As Integer
    Dim Bottom(10) As Integer
    Dim topp(10) As Integer
    Dim pt As Integer
    Dim tempi As Integer
    Dim i As Integer
    Dim j As Integer
    Dim hStart As Integer
    Dim hStop As Integer
    
    If bln_SameDeal Then Exit Sub
  
    For h = 1 To 4
      Do While h = bln_PickedHand(h)
        If h = 4 Then Exit Do
        h = h + 1
      Loop
        
      If bln_PickedHand(h) Or int_NumPickedHands = 0 Then
        ply = 1
        cStart = (h - 1) * 13 + 1
        cStop = cStart + 12
        Bottom(1) = cStart
        topp(1) = cStop
    
        While ply > 0
            If Bottom(ply) >= topp(ply) Then
                ply = ply - 1
            Else
                i = Bottom(ply) - 1
                j = topp(ply)
                pt = int_Cards(j)
        
                While i < j
                    i = i + 1
                    j = j - 1
                    
                    While int_Cards(i) < pt
                        i = i + 1
                    Wend
          
                    While int_Cards(j) > pt _
                          And j > i
                        j = j - 1
                    Wend
          
                    If i < j Then
                        ' Swap: '
                        tempi = int_Cards(i)
                        int_Cards(i) = int_Cards(j)
                        int_Cards(j) = tempi
                    End If
                Wend
        
                j = topp(ply)
                
                If i <> j And _
                    int_Cards(i) > int_Cards(j) _
                Then
                    ' Swap: '
                    tempi = int_Cards(i)
                    int_Cards(i) = int_Cards(j)
                    int_Cards(j) = tempi
                End If
        
                If i - Bottom(ply) < topp(ply) - i Then
                    Bottom(ply + 1) = Bottom(ply)
                    topp(ply + 1) = i - 1
                    Bottom(ply) = i + 1
                Else
                    topp(ply + 1) = topp(ply)
                    Bottom(ply + 1) = i + 1
                    topp(ply) = i - 1
                End If
                ply = ply + 1
            End If
        Wend
      End If
    Next

    ' Now the cards are sorted by hand  '
    ' within the array, but from low to '
    ' high for each hand. We want them  '
    ' sorted from high to low for each  '
    ' hand:                             '
    For h = 1 To 4
      If bln_PickedHand(h) Or int_NumPickedHands = 0 Then
        hStart = (h - 1) * 13 + 1
        hStop = hStart + 12
    
        For i = 0 To 5
            tempi = int_Cards(hStart + i)
            int_Cards(hStart + i) = int_Cards(hStop - i)
            int_Cards(hStop - i) = tempi
        Next
      End If
    Next

End Sub


Function SuitOf(ByVal n)
    ' Calculate the suit number from '
    ' the card array pointer number: '
    Dim s As Integer
    
    s = Fix(n / 13) + 1
    If n / 13 = Fix(n / 13) Then
        s = s - 1
    End If
    SuitOf = s
End Function




Sub ScrambleHands()
    ' Scramble/redeal the cards to user- '
    ' specified hands:  bln_PickedHand(  '
    ' The array int_Cards(1-52) contains '
    ' the last shuffled set of cards.    '
    ' S's cards are 1-13, W's are 14-26, '
    ' N's are 27-39, and E's are 40-52.  '
    ' So if use wants to scramble E-W,   '
    ' int_Cards(14-26) and (40-52) are   '
    ' intermixed while (1-13) and (27-39)'
    ' are left-as is.                    '
    
    Dim h As Integer
    Dim nrh As String
    Dim m As String
    Dim n As Integer
    Dim nc As Integer
    Dim s As Integer
    Dim c As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim cards2(52) As String
    Dim temp As Integer
    Dim numCards As Integer
    
    ' Get the cards to scramble: '
    n = 0
    For h = 1 To 4
        If bln_PickedHand(h) Or int_NumPickedHands = 0 Then
            i = (h - 1) * 13
            For c = i + 1 To i + 13
                n = n + 1
                cards2(n) = int_Cards(c)
            Next
        End If
    Next
    
    ' "Shuffle" the cards: '
    For j = 1 To 10 + Int(5 * Rnd) ' # of times to shuffle the cards
        For i = 1 To n
            ' Random card to swap with current card:
            k = Int(n * Rnd + 1)
            ' Swap them: '
            temp = cards2(i)
            cards2(i) = cards2(k)
            cards2(k) = temp
        Next
    Next
    numCards = n
    
    ' Put scrambled cards back into the array:
    n = 0
    For h = 1 To 4
        If bln_PickedHand(h) _
        Or int_NumPickedHands = 0 _
        Then
            i = (h - 1) * 13
            For c = i + 1 To i + 13
                n = n + 1
                int_Cards(c) = cards2(n)
            Next
        End If
    Next
    
End Sub


