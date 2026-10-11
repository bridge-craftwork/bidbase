-- BidNames Query Backup
SELECT BidNames.Selected,BidNames.SubCat_Selected,BidNames.Bid_Name,Bids.* FROM [BidNames],[Bids] 

-- GridQuery
SELECT BidNames.Selected,BidNames.SubCat_Selected,Bids.Pct_Used,Bids.RoundOfBidding,Bids.Player,Bids.PB_Sort_Order,Bids.Prior_Bids,Bids.Bid_Order_Number,Bids.Bid,BidNames.Bid_Name,Bids.HCP,Bids.Dist_Pts,Bids.Cust_Pts,Bids.A_Cust,Bids.K_Cust,Bids.Q_Cust,Bids.J_Cust,Bids.T_Cust,Bids.Winrs,Bids.Losers,Bids.QTs,Bids.Stopped_Suits,Bids.Shape,Bids.Intermed,Bids.xNum,Bids.xHCP,Bids.sNum,Bids.sHCP,Bids.hNum,Bids.hHCP,Bids.dNum,Bids.dHCP,Bids.cNum,Bids.cHCP,Bids.cmNum,Bids.cmHCP,Bids.Vul,Bids.Pos,Bids.Artificial,Bids.Alert,Bids.Explanation,Bids.Test_Hand,Bids.Asks,Bids.DateOfEntry,Bids.Best_Contract,Bids.Num_Used,Bids.Source,Bids.LHO,Bids.Pard,Bids.RHO,Bids.ID FROM [BidNames],[Bids] 

-- BidNamesQuery
SELECT BidNames.Selected,BidNames.SubCat_Selected,BidNames.Bid_Name,Bids.ConvID FROM [BidNames],[Bids] 

