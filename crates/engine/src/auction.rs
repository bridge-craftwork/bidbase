//! The auction so far, as the VB engine sees it (`Calculate_BidNum` and
//! `Calculate_Position`).
//!
//! Calls are BidBase tokens joined by dashes: "P-1C-1N-2C-D-R-2S", with P
//! for pass, D for double, R for redouble and N for notrump. Opening
//! passes count for the seat but are dropped from the prior-bids key, so
//! "P-P-1H-P" looks up rules filed under "1H-P".

/// Facts about the auction for the player now to call.
#[derive(Debug, Clone)]
pub struct AuctionFacts {
    /// The calls as given (upper case).
    pub prior_calls: String,
    /// The rules' key: the calls after the opening passes.
    pub prior_bids: String,
    /// Every call, 1-based (index 0 unused), opening passes included.
    pub calls: Vec<String>,
    /// Calls from the first non-pass on, 1-based.
    pub bids: Vec<String>,
    /// Number of the call now being made, counting from the dealer.
    pub call_num: i32,
    /// Number of the call now being made, counting from the first bid.
    pub bid_num: i32,
    /// Bidding round, counting from the first bid.
    pub round: i32,
    /// Seat counted from the opener: 1 opener, 2 opener's LHO, 3 opener's
    /// partner, 4 opener's RHO.
    pub pos_num: i32,
    /// Seat counted from the dealer, 1-4.
    pub seat: i32,
    /// Codes the rule's `Pos` field is matched against: Y/N (bidder has /
    /// hasn't passed), 1-4 (seat, first round only), F/T (seat 1-2 / 3-4),
    /// P/U (partner passed / didn't on the last round).
    pub position: String,
}

impl AuctionFacts {
    pub fn new(prior_calls: &str) -> Self {
        // Upper case, but "Mw" (Minorwood, in 2026 ace-asking codes) keeps
        // its lower-case w, as the VB `PriorCalls` restores it.
        let prior_calls = prior_calls.trim().to_ascii_uppercase().replace("MW", "Mw");
        let tokens: Vec<String> = if prior_calls.is_empty() {
            Vec::new()
        } else {
            prior_calls
                .split('-')
                .map(|t| t.trim().to_string())
                .collect()
        };
        // At most three opening passes come off; a fourth would end it.
        let passes = tokens.iter().take(3).take_while(|t| *t == "P").count();
        let rest = &tokens[passes..];

        let mut calls = vec![String::new()];
        calls.extend(tokens.iter().cloned());
        let mut bids = vec![String::new()];
        bids.extend(rest.iter().cloned());

        let call_num = tokens.len() as i32 + 1;
        let bid_num = rest.len() as i32 + 1;
        let round = (bid_num + 3) / 4;
        let pos_num = match bid_num % 4 {
            0 => 4,
            n => n,
        };
        let seat = (call_num - 1) % 4 + 1;

        let mut facts = AuctionFacts {
            prior_bids: rest.join("-"),
            prior_calls,
            calls,
            bids,
            call_num,
            bid_num,
            round,
            pos_num,
            seat,
            position: String::new(),
        };
        facts.position = facts.position_codes();
        facts
    }

    /// The call with 1-based number `n`, or "" before the auction.
    pub fn call(&self, n: i32) -> &str {
        if n < 1 {
            return "";
        }
        self.calls.get(n as usize).map_or("", String::as_str)
    }

    /// The bid with 1-based number `n` (counted from the first bid).
    pub fn bid(&self, n: i32) -> &str {
        if n < 1 {
            return "";
        }
        self.bids.get(n as usize).map_or("", String::as_str)
    }

    /// VB `Calculate_Position`, reading seats from the dealer.
    ///
    /// The 2018 code took the bidder's seat from `int_PosNum`, which counts
    /// from the opener, so an opener was always "F" (a 3rd-seat "Rule of
    /// 20" opening passed the `F` test) and an overcaller in 3rd seat was
    /// never "T". The rules' own notes ("1x in 1st seat, 2x in 2nd, 3x in
    /// 3rd", "pard passed" on `T` overcalls) read the seat from the dealer.
    fn position_codes(&self) -> String {
        let mut p = String::new();
        // Passed hand: the bidder's first call (one round ago or more).
        if self.call_num < 5 || self.call(self.seat) != "P" {
            p.push('N');
        } else {
            p.push('Y');
        }
        if self.call_num < 5 {
            p.push_str(&self.call_num.to_string());
        }
        p.push(if self.seat < 3 { 'F' } else { 'T' });
        // Partner's last call, two calls back.
        let mut pard = self.call_num - 2;
        if pard < 1 {
            pard += 4;
        }
        if self.call_num < 3 || self.call(pard) != "P" {
            p.push('U');
        } else {
            p.push('P');
        }
        p
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn opening_passes_are_dropped_from_the_key() {
        let a = AuctionFacts::new("P-P-1H-P");
        assert_eq!(a.prior_bids, "1H-P");
        assert_eq!(a.call_num, 5);
        assert_eq!(a.bid_num, 3);
        assert_eq!(a.pos_num, 3);
        assert_eq!(a.seat, 1);
        assert_eq!(a.position, "YFU");
    }

    #[test]
    fn the_first_call() {
        let a = AuctionFacts::new("");
        assert_eq!(a.prior_bids, "");
        assert_eq!((a.call_num, a.bid_num, a.round, a.pos_num), (1, 1, 1, 1));
        assert_eq!(a.position, "N1FU");
    }

    #[test]
    fn third_seat_opener() {
        let a = AuctionFacts::new("p-p");
        assert_eq!(a.prior_bids, "");
        assert_eq!(a.pos_num, 1);
        assert_eq!(a.position, "N3TP");
    }

    #[test]
    fn counts_rounds_from_the_first_bid() {
        let a = AuctionFacts::new("1C-P-1H-P-1S-P");
        assert_eq!(a.bid_num, 7);
        assert_eq!(a.round, 2);
        assert_eq!(a.pos_num, 3);
        assert_eq!(a.bid(5), "1S");
    }
}
