-- ----------------------------------------------------------
-- MDB Tools - A library for reading MS Access database files
-- Copyright (C) 2000-2011 Brian Bruns and others.
-- Files in libmdb are licensed under LGPL and the utilities under
-- the GPL, see COPYING.LIB and COPYING files respectively.
-- Check out http://mdbtools.sourceforge.net
-- ----------------------------------------------------------

-- That file uses encoding UTF-8

CREATE TABLE [BidNames]
 (
	[IDnum]			Long Integer, 
	[Selected]			Text (1), 
	[SubCat_Selected]			Text (1), 
	[Bid_Name]			Text (50) NOT NULL, 
	[Bid_Description]			Text (100)
);

CREATE TABLE [Bids]
 (
	[Pct_Used]			Text (3) NOT NULL, 
	[RoundOfBidding]			Text (1) NOT NULL, 
	[Player]			Text (1) NOT NULL, 
	[PB_Sort_Order]			Text (50), 
	[Prior_Bids]			Text (50) NOT NULL, 
	[Bid_Order_Number]			Text (6) NOT NULL, 
	[Bid]			Text (5) NOT NULL, 
	[ConvID]			Long Integer, 
	[Dist_Pts]			Text (6) NOT NULL, 
	[Cust_Pts]			Text (5) NOT NULL, 
	[A_Cust]			Text (3) NOT NULL, 
	[K_Cust]			Text (3) NOT NULL, 
	[Q_Cust]			Text (3) NOT NULL, 
	[J_Cust]			Text (3) NOT NULL, 
	[T_Cust]			Text (3) NOT NULL, 
	[Winrs]			Text (5) NOT NULL, 
	[Losers]			Text (5) NOT NULL, 
	[QTs]			Text (5) NOT NULL, 
	[Stopped_Suits]			Text (1) NOT NULL, 
	[Shape]			Text (1) NOT NULL, 
	[Intermed]			Text (1) NOT NULL, 
	[xNum]			Text (10) NOT NULL, 
	[xHCP]			Text (10) NOT NULL, 
	[sNum]			Text (10) NOT NULL, 
	[sHCP]			Text (10) NOT NULL, 
	[hNum]			Text (10) NOT NULL, 
	[hHCP]			Text (10) NOT NULL, 
	[dNum]			Text (10) NOT NULL, 
	[dHCP]			Text (10) NOT NULL, 
	[cNum]			Text (10) NOT NULL, 
	[cHCP]			Text (10) NOT NULL, 
	[cmNum]			Text (9), 
	[cmHCP]			Text (25), 
	[Vul]			Text (1) NOT NULL, 
	[Pos]			Text (1) NOT NULL, 
	[Artificial]			Text (1) NOT NULL, 
	[Alert]			Text (1) NOT NULL, 
	[Explanation]			Text (50) NOT NULL, 
	[Test_Hand]			Text (16) NOT NULL, 
	[Asks]			Text (5) NOT NULL, 
	[DateOfEntry]			DateTime, 
	[Best_Contract]			Text (5) NOT NULL, 
	[Num_Used]			Long Integer NOT NULL, 
	[Source]			Text (25) NOT NULL, 
	[ID]			Long Integer, 
	[HCP]			Text (12) NOT NULL
);

CREATE TABLE [Combos]
 (
	[ID]			Long Integer, 
	[Missing]			Text (4), 
	[NumCards]			Integer, 
	[Each]			Integer, 
	[Need]			Integer, 
	[PctMake]			Long Integer, 
	[North]			Text (12), 
	[South]			Text (12), 
	[Play1]			Text (4), 
	[Play2]			Text (4), 
	[Play3]			Text (4)
);

CREATE TABLE [DupePBs]
 (
	[Description]			Text (30), 
	[pb1]			Text (50) NOT NULL, 
	[pb2]			Text (50), 
	[pb3]			Text (50), 
	[pb4]			Text (50), 
	[pb5]			Text (50), 
	[pb6]			Text (50), 
	[pb7]			Text (50), 
	[pb8]			Text (50), 
	[pb9]			Text (50)
);

CREATE TABLE [Garbage]
 (
	[Pct_Used]			Text (3) NOT NULL, 
	[RoundOfBidding]			Text (1) NOT NULL, 
	[Player]			Text (1) NOT NULL, 
	[PB_Sort_Order]			Text (50), 
	[Prior_Bids]			Text (50) NOT NULL, 
	[Bid_Order_Number]			Text (6) NOT NULL, 
	[Bid]			Text (5) NOT NULL, 
	[ConvID]			Long Integer, 
	[HCP]			Text (6) NOT NULL, 
	[Dist_Pts]			Text (6) NOT NULL, 
	[Cust_Pts]			Text (5) NOT NULL, 
	[A_Cust]			Text (3) NOT NULL, 
	[K_Cust]			Text (3) NOT NULL, 
	[Q_Cust]			Text (3) NOT NULL, 
	[J_Cust]			Text (3) NOT NULL, 
	[T_Cust]			Text (3) NOT NULL, 
	[Winrs]			Text (5) NOT NULL, 
	[Losers]			Text (5) NOT NULL, 
	[QTs]			Text (5) NOT NULL, 
	[Stopped_Suits]			Text (1) NOT NULL, 
	[Shape]			Text (1) NOT NULL, 
	[Intermed]			Text (1) NOT NULL, 
	[xNum]			Text (10) NOT NULL, 
	[xHCP]			Text (10) NOT NULL, 
	[sNum]			Text (10) NOT NULL, 
	[sHCP]			Text (10) NOT NULL, 
	[hNum]			Text (10) NOT NULL, 
	[hHCP]			Text (10) NOT NULL, 
	[dNum]			Text (10) NOT NULL, 
	[dHCP]			Text (10) NOT NULL, 
	[cNum]			Text (10) NOT NULL, 
	[cHCP]			Text (10) NOT NULL, 
	[cmNum]			Text (9), 
	[cmHCP]			Text (25), 
	[Vul]			Text (1) NOT NULL, 
	[Pos]			Text (1) NOT NULL, 
	[Artificial]			Text (1) NOT NULL, 
	[Alert]			Text (1) NOT NULL, 
	[Explanation]			Text (50) NOT NULL, 
	[Test_Hand]			Text (16) NOT NULL, 
	[Asks]			Text (5) NOT NULL, 
	[DateOfEntry]			DateTime, 
	[Best_Contract]			Text (5) NOT NULL, 
	[Num_Used]			Long Integer NOT NULL, 
	[Source]			Text (25) NOT NULL, 
	[LHO]			Long Integer, 
	[Pard]			Long Integer, 
	[RHO]			Long Integer, 
	[ID]			Long Integer, 
	[BidsID]			Long Integer NOT NULL
);

CREATE TABLE [HandData]
 (
	[Specs]			Text (92) NOT NULL, 
	[ID]			Long Integer
);


