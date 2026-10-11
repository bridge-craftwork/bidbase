Attribute VB_Name = "dds"
Option Explicit

Private Declare Sub CalcDDtablePBNDDS Lib "dds.dll" Alias "CalcDDtablePBNVB" (ByRef ddaDeal As ddTableDealPBN, ByRef res As ddTableResults)
               
' See DLL-Description.html in the dds290complete.zip file in Bridge/Notes. '

Public Type futureTricks
    nodes As Long
    cards As Long
    suit(12) As Long
    rank(12) As Long
    equals(12) As Long
    score(12) As Long
End Type

Public Type ddaDeal
    trump As Long
    first As Long
    currentTrickSuit(2) As Long
    currentTrickRank(2) As Long
    remainCards(3, 3) As Long
End Type

Public Type dealPBN  ' This is used to calculate the play(?). '
    trump As Long
    first As Long
    currentTrickSuit(2) As Long
    currentTrickRank(2) As Long
    remainCards(79) As Byte ' The DLL description says this is string, not byte. '
End Type

Public Type boards
    noOfBoards As Long
    deals(199) As ddaDeal
    target(199) As Long
    solutions(199) As Long
    mode(199) As Long
End Type

Public Type boardsPBN
    noOfBoards As Long
    deals(199) As dealPBN
    target(199) As Long
    solutions(199) As Long
    mode(199) As Long
End Type

Public Type solvedBoards
    noOfBoards As Long
    solvedBoard(199) As futureTricks
End Type

Public Type ddTableDeal
    cards(3, 3) As Long
End Type

Public Type ddTableDeals
    noOfTables As Long
    deals(999) As ddTableDeal
End Type

Public Type ddTableDealPBN ''''''''''''''''''''''''''''''''
    cards(79) As Byte
End Type

Public Type ddTableResults ''''''''''''''''''''''''''''''''
    resTable(4, 3) As Long
End Type

Public Type ddTableRes
    noOfBoards As Long
    results(199) As ddTableResults
End Type

Public Type parResults
    parScore(1, 15) As Byte
    parContractsString(1, 127) As Byte
End Type

Public Type allParResults
    presults(39) As parResults
End Type

Public Type DDSInfo
    major As Long
    minor As Long
    patch As Long
    versionString(9) As Byte
    system As Long
    numBits As Long
    compiler As Long
    constructor As Long
    numCores As Long
    threading As Long
    threadsizes(127) As Byte
    systemString(1023) As Byte
End Type


Public Function IntToSuit(iSuit As Integer) As String
    Select Case iSuit
        Case 0
            IntToSuit = "Spades"
        Case 1
            IntToSuit = "Hearts"
        Case 2
            IntToSuit = "Diamonds"
        Case 3
            IntToSuit = "Clubs"
        Case 4
            IntToSuit = "Notrump"
        Case Else
            IntToSuit = "Error - suit unknown"
    End Select
End Function


Public Function IntToHand(iHand As Integer) As String
    Select Case iHand
        Case 0
            IntToHand = "North"
        Case 1
            IntToHand = "East"
        Case 2
            IntToHand = "South"
        Case 3
            IntToHand = "West"
        Case Else
            IntToHand = "Error - hand unknown"
    End Select
End Function


Public Function IntToVuln(iVuln As Integer) As String
    Select Case iVuln
        Case 0
            IntToVuln = "None"
        Case 1
            IntToVuln = "Both"
        Case 2
            IntToVuln = "NS only"
        Case 3
            IntToVuln = "EW only"
        Case Else
            IntToVuln = "Error - vulnerability unknown"
    End Select
End Function


Public Function IntToSide(iSide As Integer) As String
    Select Case iSide
        Case 0
            IntToSide = "N-S"
        Case 1
            IntToSide = "E-W"
        Case Else
            IntToSide = "Error - side unknown"
    End Select
End Function


Public Function CalcDDTablePBN(pbn As ddTableDealPBN) As ddTableResults
    Dim res As ddTableResults
    Dim temp(19) As Integer
 
    ChDir App.Path
    CalcDDtablePBNDDS pbn, res
 
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
 
    k = 0
    ' Flatten the array: '
    For i = 0 To 3
        For j = 0 To 4
            temp(k) = res.resTable(j, i)
            k = k + 1
        Next
    Next
 
    k = 0
    ' Reconstruct the array inversely: '
    For i = 0 To 4
        For j = 0 To 3
            res.resTable(i, j) = temp(k)
            k = k + 1
        Next
    Next
 
    CalcDDTablePBN = res

End Function


