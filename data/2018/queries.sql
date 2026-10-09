-- BidNamesQuery
SELECT BidNames.Selected,BidNames.SubCat_Selected,BidNames.Bid_Name,BidNames_1.Bid_Name,Bids.* FROM [BidNames],[BidNames],[Bids] 

-- BidNames Query Backup
SELECT BidNames.Selected,BidNames.SubCat_Selected,BidNames.Bid_Name,BidNames_1.Selected,BidNames_1.SubCat_Selected,BidNames_1.Bid_Name,Bids.* FROM [BidNames],[BidNames],[Bids] 

-- GridQuery
SELECT BidNames.Selected,BidNames.SubCat_Selected,Bids.Pct_Used,Bids.RoundOfBidding,Bids.Player,Bids.PB_Sort_Order,Bids.Prior_Bids,Bids.Bid_Order_Number,Bids.Bid,BidNames.Bid_Name,UseAgainst.Bid_Name,Bids.HCP,Bids.Dist_Pts,Bids.ValSystem,Bids.SuitVal,Bids.NT_Val,Bids.Cust_Pts,Bids.A_Cust,Bids.K_Cust,Bids.Q_Cust,Bids.J_Cust,Bids.T_Cust,Bids.Winrs,Bids.Losers,Bids.QTs,Bids.Stopped_Suits,Bids.Shape,Bids.Intermed,Bids.xNum,Bids.xHCP,Bids.sNum,Bids.sHCP,Bids.hNum,Bids.hHCP,Bids.dNum,Bids.dHCP,Bids.cNum,Bids.cHCP,Bids.cmNum,Bids.cmHCP,Bids.Vul,Bids.Pos,Bids.Artificial,Bids.Alert,Bids.Explanation,Bids.Test_Hand,Bids.Strength_Discl,Bids.xLen,Bids.SLen,Bids.HLen,Bids.DLen,Bids.CLen,Bids.Asks,Bids.DateOfEntry,Bids.Best_Contract,Bids.Num_Used,Bids.Source,Bids.ID FROM [BidNames],[BidNames],[Bids] 

