//! The per-rule tests (VB `TotalHCPs_Ok` ... `Eval_Pts_Ok`). Each returns
//! `Err(reason)` when the hand fails the rule, which stands in for the
//! VB analysis form's "NO MATCH" lines.
//!
//! Where the 2018 code did something its own comments and data say it did
//! not mean, the port does what was meant; docs/PORTING.md lists each case.

use crate::auction::AuctionFacts;
use crate::hand::{HandFacts, ACE, KING, QUEEN};
use crate::lookup::Request;
use crate::rules::{Dialect, NtRanges, Rule};
use crate::vb::{ival, split_alt, suit_index, suit_letter, val};

pub type Check = Result<(), String>;

pub(crate) type Test<'a> = fn(&mut Ctx<'a>, &Rule) -> Check;

/// What the checks read, plus the state one rule's tests share.
pub struct Ctx<'a> {
    pub hand: &'a HandFacts,
    pub auction: &'a AuctionFacts,
    pub req: &'a Request,
    pub nt: NtRanges,
    /// The hand's intermediates, credited with any HCP or total points
    /// beyond the rule's minimum (one per extra point).
    pub intermediates: i32,
    /// The suit (1-4) that matched the rule's "x" specs, 0 if none.
    pub x_suit: usize,
    pub dialect: Dialect,
    /// The trump suit (1-4, 0 unknown) for ace-asking replies: from the
    /// request, else (2026) the "@" code of the prior bids. A 2026 rule's
    /// "Trmp" can move it; `trump0` keeps the one the counts were made for.
    pub trump: usize,
    pub trump0: usize,
}

impl<'a> Ctx<'a> {
    pub fn new(
        hand: &'a HandFacts,
        auction: &'a AuctionFacts,
        req: &'a Request,
        nt: NtRanges,
        dialect: Dialect,
    ) -> Self {
        let trump = match (req.trump, dialect) {
            (0, Dialect::V2026) => crate::checks2026::trump_from_prior(&auction.prior_bids),
            (t, _) => t,
        };
        Ctx {
            hand,
            auction,
            req,
            nt,
            intermediates: hand.intermediates[0],
            x_suit: 0,
            dialect,
            trump,
            trump0: trump,
        }
    }

    /// Every test after `Pct_Used`, in the VB order. The first failure is
    /// returned with the name of the test.
    pub fn run(&mut self, r: &Rule) -> Result<(), (&'static str, String)> {
        // State is per rule: the VB kept the intermediates credit in a
        // module variable, so it grew with every rule tested.
        self.intermediates = self.hand.intermediates[0];
        self.x_suit = 0;
        self.trump = self.trump0;
        if self.dialect == Dialect::V2026 {
            let tests: [(&'static str, Test<'a>); 12] = [
                ("HCP", Self::total_hcp_2026),
                ("Cust_Pts", Self::custom_points_2026),
                ("Dist_Pts", Self::distribution_points_2026),
                ("Winrs", Self::winners),
                ("Losers", Self::losers_2026),
                ("QTs", Self::quick_tricks),
                ("Stopped_Suits", Self::stoppers),
                ("Shape", Self::shape),
                ("Intermed", Self::intermed_2026),
                ("Vul", Self::vul_2026),
                ("Pos", Self::position_2026),
                ("Suits", Self::suits_2026),
            ];
            for (name, test) in tests {
                test(self, r).map_err(|why| (name, why))?;
            }
            return Ok(());
        }
        let tests: [(&'static str, Test<'a>); 13] = [
            ("HCP", Self::total_hcp),
            ("Cust_Pts", Self::custom_points),
            ("Dist_Pts", Self::distribution_points),
            ("Winrs", Self::winners),
            ("Losers", Self::losers),
            ("QTs", Self::quick_tricks),
            ("Stopped_Suits", Self::stoppers),
            ("Shape", Self::shape),
            ("Intermed", Self::intermed),
            ("Vul", Self::vul),
            ("Pos", Self::position),
            ("Suits", Self::suits),
            ("ValSystem", Self::valuation),
        ];
        for (name, test) in tests {
            test(self, r).map_err(|why| (name, why))?;
        }
        Ok(())
    }

    fn trump(&self, r: &Rule) -> usize {
        match r.source_trump() {
            0 => self.req.trump,
            t => t,
        }
    }

    // ----- HCP -----------------------------------------------------------

    /// VB `TotalHCPs_Ok`.
    fn total_hcp(&mut self, r: &Rule) -> Check {
        let spec = r.hcp.trim().to_ascii_lowercase();
        if spec.is_empty() {
            return Ok(());
        }
        let have = self.hand.hcp[0];
        if let Some(result) = self.count_spec(&spec, r) {
            return result;
        }
        let (lo, hi) = hcp_bounds(&spec, &self.auction.prior_bids, &r.bid, self.nt);
        if have < lo || have > hi {
            return Err(format!("{} HCP, needs {lo}-{hi}", have));
        }
        if ival(&r.intermed) > 0 && have > lo {
            self.intermediates += have - lo;
        }
        Ok(())
    }

    /// Ace, king and key-card replies written in the HCP field: "A2",
    /// "K1", "KC2+Q" (two key cards with the trump queen). The 2018 code
    /// read these as numbers (all 0), so they only matched 0-HCP hands.
    fn count_spec(&self, spec: &str, r: &Rule) -> Option<Check> {
        let (kind, rest) = [("kc", "kc"), ("a", "a"), ("k", "k")]
            .into_iter()
            .find_map(|(kind, prefix)| Some((kind, spec.strip_prefix(prefix)?)))?;
        let digits: String = rest.chars().take_while(char::is_ascii_digit).collect();
        if digits.is_empty() {
            return None;
        }
        let want: i32 = digits.parse().ok()?;
        let queen = &rest[digits.len()..];
        Some(match kind {
            "a" => eq_or(self.hand.aces(), want, "aces"),
            "k" => eq_or(self.hand.kings(), want, "kings"),
            _ => {
                let trump = self.trump(r);
                if trump == 0 {
                    return Some(Err("key cards need a trump suit".into()));
                }
                let kc = self.hand.aces() + i32::from(self.hand.held[trump][KING]);
                let tq = self.hand.held[trump][QUEEN];
                eq_or(kc, want, "key cards").and_then(|()| match queen {
                    "+q" if !tq => Err("no trump queen".into()),
                    "-q" if tq => Err("has the trump queen".into()),
                    _ => Ok(()),
                })
            }
        })
    }

    // ----- Custom points ---------------------------------------------------

    /// VB `CustPoints_Ok`: points per honour from `A_Cust`..`T_Cust`; with
    /// `Cust_Pts` a minimum ("<n" a maximum) total, else counts of honours.
    fn custom_points(&mut self, r: &Rule) -> Check {
        let cp = [
            0.0,
            val(&r.a_cust),
            val(&r.k_cust),
            val(&r.q_cust),
            val(&r.j_cust),
            val(&r.t_cust),
        ];
        let specified = cp[1..].iter().any(|&v| v > 0.0);
        let rp = r.cust_pts.trim();
        let mut reqd = val(rp);
        let mut less = false;
        if reqd == 0.0 && rp.len() > 1 {
            reqd = val(&rp[1..]);
            less = rp.starts_with('<');
        }
        if reqd == 0.0 && !specified {
            return Ok(());
        }
        let mut total = 0.0;
        let mut count = [0.0; 6];
        for s in 1..=4 {
            for card in 10..=14 {
                if self.hand.held[s][card] {
                    let c = 15 - card; // ace 1 ... ten 5
                    total += cp[c];
                    count[c] += 1.0;
                }
            }
        }
        if reqd > 0.0 {
            let ok = if less { total < reqd } else { total >= reqd };
            return if ok {
                Ok(())
            } else {
                Err(format!(
                    "{total} custom points, needs {}{reqd}",
                    if less { "<" } else { "" }
                ))
            };
        }
        if (1..=5).all(|c| count[c] >= cp[c]) {
            Ok(())
        } else {
            Err("doesn't have the specified honours".into())
        }
    }

    // ----- Distribution points ---------------------------------------------

    /// VB `DistribPts_Ok`: "2+", "<3", "3-4"; with a leading T the spec is
    /// for distribution points plus HCP ("T13+").
    fn distribution_points(&mut self, r: &Rule) -> Check {
        let spec = r.dist_pts.trim();
        if spec.is_empty() {
            return Ok(());
        }
        let (total, body) = match spec.strip_prefix('T') {
            Some(rest) => (true, rest),
            None => (false, spec),
        };
        let have = self.hand.distrib[0] + if total { self.hand.hcp[0] } else { 0 };
        let (lo, hi) = int_range(body, 40);
        if have < lo || have > hi {
            let what = if total {
                "total points"
            } else {
                "distribution points"
            };
            return Err(format!("{have} {what}, needs {lo}-{hi}"));
        }
        if total && ival(&r.intermed) > 0 && have > lo {
            self.intermediates += have - lo;
        }
        Ok(())
    }

    // ----- Winners, losers, quick tricks, stoppers ---------------------------

    /// VB `Winners_Ok`: a bare number is a minimum ("7.5" = 7.5+).
    fn winners(&mut self, r: &Rule) -> Check {
        min_spec(r.winners.trim(), self.hand.winners[0], "winners")
    }

    /// VB `Losers_Ok`. A bare number is a minimum here too, as written.
    fn losers(&mut self, r: &Rule) -> Check {
        min_spec(r.losers.trim(), self.hand.losers[0] as f64, "losers")
    }

    /// VB `QuickTricks_Ok`.
    fn quick_tricks(&mut self, r: &Rule) -> Check {
        min_spec(
            r.quick_tricks.trim(),
            self.hand.quick_tricks[0],
            "quick tricks",
        )
    }

    /// VB `Stoppers_Ok`: at least this many suits stopped.
    fn stoppers(&mut self, r: &Rule) -> Check {
        min_spec(r.stopped_suits.trim(), self.hand.stops[0], "suits stopped")
    }

    // ----- Shape, intermediates, vulnerability, position --------------------

    /// VB `Shape_Ok`: 1/2/3-suiter, b(alanced), u(nbalanced), s(ingleton),
    /// v(oid), e(ither).
    fn shape(&mut self, r: &Rule) -> Check {
        let want = r.shape.trim().to_ascii_lowercase();
        let (shape, short) = (self.hand.shape, self.hand.shortness);
        let ok = match want.as_str() {
            "" => true,
            "1" | "2" | "3" | "b" => shape == want,
            "u" => shape != "b",
            "e" => !short.is_empty(),
            "s" | "v" => short == want,
            _ => false,
        };
        if ok {
            Ok(())
        } else {
            Err(format!("shape is {shape:?}{short}, needs {want}"))
        }
    }

    /// VB `Intermed_Ok`: intermediates, credited with extra points above.
    fn intermed(&mut self, r: &Rule) -> Check {
        if r.intermed.is_empty() {
            return Ok(());
        }
        let want = ival(&r.intermed);
        if self.intermediates >= want {
            Ok(())
        } else {
            Err(format!(
                "{} intermediates, needs {want}",
                self.intermediates
            ))
        }
    }

    /// VB `Vul_Ok`. A rule's `Vul` (Y, N, Fav, Unfav, Eq, Eq+) must be one
    /// of the tokens the vulnerability carries (`Vul::tokens`).
    fn vul(&mut self, r: &Rule) -> Check {
        let want = r.vul.trim();
        let have = self.req.vul.tokens();
        let ok = if want.is_empty() {
            // Testing an entry at "N": blank-vul rules stand aside.
            !(have == ["N"] && self.req.test_id != 0)
        } else {
            have.contains(&want) || (have.is_empty() && want == "N")
        };
        if ok {
            Ok(())
        } else {
            Err(format!("vulnerability {have:?}, needs {want}"))
        }
    }

    /// VB `Position_Ok`. Ignored for the entry under test.
    fn position(&mut self, r: &Rule) -> Check {
        let want = r.pos.trim().to_ascii_uppercase();
        if want.is_empty() || (self.req.test_id > 0 && self.req.test_id == r.id) {
            return Ok(());
        }
        if self.auction.position.contains(&want) {
            Ok(())
        } else {
            Err(format!("position {}, needs {want}", self.auction.position))
        }
    }

    // ----- Suits -----------------------------------------------------------

    /// VB `Suits_Ok`: length and strength specs per suit, the "x" suit
    /// (the suit a "2x" bid names), and the combined-suit specs.
    fn suits(&mut self, r: &Rule) -> Check {
        let hand = self.hand;
        let db_num: Vec<String> = (0..=4).map(|s| r.suit_num(s).replace(' ', "")).collect();
        let db_hcp: Vec<String> = (0..=4).map(|s| r.suit_hcp(s).replace(' ', "")).collect();
        let mut x_num = r.x_num.trim().to_string();
        // "U" after xNum: look for the x suit from clubs up, not spades down.
        let order: [usize; 4] = if x_num.ends_with('U') {
            x_num.pop();
            [1, 2, 3, 4]
        } else {
            [4, 3, 2, 1]
        };
        // "+" after xNum: a longer suit later in the order takes over.
        let x_plus = x_num.ends_with('+');
        let mut longest_x = 0usize;
        let mut compare: Vec<(usize, String, String)> = Vec::new();
        let mut failed: Option<String> = None;
        self.x_suit = 0;

        for s in order {
            let (mut num1, mut num2) = split_alt(&db_num[s]);
            let (mut pts1, mut pts2) = split_alt(&db_hcp[s]);
            if num1.is_empty() && pts1.is_empty() {
                continue;
            }
            let longest_len = if longest_x == 0 {
                0
            } else {
                hand.len[longest_x]
            };
            let mut x_field = false;
            if self.x_suit == 0 || (hand.len[s] > longest_len && x_plus) {
                if num1 == "x" {
                    num1 = x_num.clone();
                    x_field = true;
                }
                if pts1 == "x" {
                    pts1 = r.x_hcp.trim().to_string();
                    x_field = true;
                }
            } else if num1 == "x" || pts1 == "x" {
                // The x suit is found: this suit takes its alternative.
                num1 = std::mem::take(&mut num2);
                pts1 = std::mem::take(&mut pts2);
            }
            loop {
                if self.suit_ok(s, &num1, &pts1) {
                    if x_field {
                        self.x_suit = s;
                        if longest_x == 0 || hand.len[s] > hand.len[longest_x] {
                            longest_x = s;
                        }
                    } else if num1.ends_with('x') || pts1.ends_with('x') {
                        compare.push((s, num1.clone(), pts1.clone()));
                    }
                    break;
                }
                x_field = false;
                if !num2.is_empty() {
                    (num1, num2) = split_alt(&num2);
                    if !pts2.is_empty() {
                        (pts1, pts2) = split_alt(&pts2);
                    } else if db_hcp[s].starts_with('x') {
                        pts1.clear();
                    }
                } else if !pts2.is_empty() {
                    (pts1, pts2) = split_alt(&pts2);
                } else {
                    // An "x" suit that doesn't qualify is no failure: the
                    // x can be another suit.
                    if !(db_num[s] == "x" && (db_hcp[s].is_empty() || db_hcp[s] == "x")) {
                        failed.get_or_insert(format!(
                            "{}: \"{}\" fails {}/{}",
                            suit_letter(s),
                            hand.suit[s],
                            db_num[s],
                            db_hcp[s]
                        ));
                    }
                    break;
                }
            }
        }

        if self.x_suit == 0 && (!r.x_num.trim().is_empty() || !r.x_hcp.trim().is_empty()) {
            return Err(format!("no suit matches x specs {}/{}", r.x_num, r.x_hcp));
        }
        if let Some(why) = failed {
            return Err(why);
        }
        // A "2x" bid needs a suit for x. The VB returned the bare level
        // ("5") when nothing named one.
        if r.bid.contains('x') && self.x_suit == 0 {
            return Err(format!("{} names no suit for x", r.bid));
        }
        // "<x", ">=x" ...: compare with the x suit's length or points.
        if self.x_suit > 0 {
            let x = self.x_suit;
            for (s, num, pts) in &compare {
                for (spec, mine, theirs, what) in [
                    (num, hand.len[*s], hand.len[x], "length"),
                    (pts, hand.hcp[*s], hand.hcp[x], "HCP"),
                ] {
                    if spec.is_empty() {
                        continue;
                    }
                    let op = spec.trim_end_matches('x');
                    if !compare_op(op, mine, theirs) {
                        return Err(format!(
                            "{} {what} {mine} not {op} x suit's {theirs}",
                            suit_letter(*s)
                        ));
                    }
                }
            }
        }
        self.combo(r)
    }

    /// VB `Suit_Ok`: a suit's length then strength.
    fn suit_ok(&self, s: usize, num: &str, hcps: &str) -> bool {
        if !num.is_empty() && !num.ends_with('x') && !self.suit_qty_ok(s, num) {
            return false;
        }
        if !hcps.is_empty() && !hcps.ends_with('x') && !self.suit_hcp_ok(s, hcps) {
            return false;
        }
        true
    }

    /// VB `SuitQty_Ok`: "4", "4+", "4-" (fewer than 4), "2-4", "2-S" (2 up
    /// to the spade length), "<4", "<H", "<=4", "<=H", ">..", "=4", "=H".
    fn suit_qty_ok(&self, s: usize, num: &str) -> bool {
        let len = self.hand.len[s];
        let other = |t: &str| -> i32 {
            match t.chars().next().map(suit_index) {
                Some(o) if o > 0 && t.len() == 1 => self.hand.len[o],
                _ => ival(t),
            }
        };
        let n = ival(num);
        if num == "0" {
            len == 0
        } else if num.ends_with('+') {
            len >= n
        } else if num.ends_with('-') {
            len < n
        } else if let Some(t) = num.strip_prefix("<=") {
            len <= other(t)
        } else if let Some(t) = num.strip_prefix(">=") {
            len >= other(t)
        } else if let Some(t) = num.strip_prefix('<') {
            len < other(t)
        } else if let Some(t) = num.strip_prefix('>') {
            len > other(t)
        } else if let Some(t) = num.strip_prefix('=') {
            len == other(t)
        } else if let Some((_, hi)) = num.split_once('-') {
            len >= n && len <= other(hi)
        } else {
            n > 0 && len == n
        }
    }

    /// VB `SuitHCP_Ok`: points ("3", "3+", "<4", "2-4"), honours ("A",
    /// "K", "Q", codes a/b/c/d and pairs like "ab"), stoppers ("1stop",
    /// "2stop", "hstop", "0stop"), controls ("FRC", "SRC"), protected
    /// honours ("P"), and "i2" for intermediates, alone or after a spec.
    fn suit_hcp_ok(&self, s: usize, spec: &str) -> bool {
        let hand = self.hand;
        let original = spec.replace(' ', "");
        if original.is_empty() {
            return true;
        }
        let mut hcps = original.to_ascii_lowercase();
        let pts = hand.hcp[s];
        let mut intermed_reqd = 0;
        if let Some(i) = hcps.find('i') {
            intermed_reqd = ival(&hcps[i + 1..]);
            if i == 0 {
                // "i2": two intermediates, or one with an honour; any suit
                // with 5+ HCP passes.
                if pts > 4 {
                    return true;
                }
                if pts > 0 {
                    intermed_reqd -= 1;
                }
                return hand.intermediates[s] >= intermed_reqd;
            }
            hcps.truncate(i);
        }
        let mut ok = false;
        let mut extra = false;
        if !hcps.is_empty() {
            let suit = &hand.suit[s];
            if hcps == "p" {
                ok = self.protected(s);
            } else if hcps == "1stop" {
                ok = hand.stops[s] >= 1.0;
                extra = hand.stops[s] > 1.0;
            } else if hcps == "2stop" {
                ok = hand.stops[s] > 1.0;
                extra = hand.stops[s] > 2.0;
            } else if hcps == "hstop" {
                ok = hand.stops[s] > 0.0;
                extra = hand.stops[s] > 0.5;
            } else if hcps == "0stop" {
                ok = hand.stops[s] < 1.0;
            } else if original.contains('A') {
                ok = hand.held[s][ACE];
            } else if hcps.contains('k') {
                ok = hand.held[s][KING];
            } else if hcps.contains('q') {
                ok = hand.held[s][QUEEN];
            } else if hcps == "frc" {
                ok = hand.len[s] == 0 || hand.held[s][ACE];
            } else if hcps == "src" {
                ok = hand.len[s] == 1 || hand.held[s][KING];
            } else if let Some(t) = hcps.strip_prefix('<') {
                ok = pts < self.other_hcp(t);
            } else if let Some(t) = hcps.strip_prefix('>') {
                // Not in the 2018 code ("<"-only); the data has ">H".
                ok = pts > self.other_hcp(t);
            } else if val(&hcps) > 0.0 || hcps == "0" {
                let need = ival(&hcps);
                if hcps.ends_with('+') {
                    ok = pts >= need;
                    extra = pts > need;
                    // Two extra HCP stand in for an intermediate.
                    intermed_reqd = (intermed_reqd - (pts - need) / 2).max(0);
                } else if hcps.ends_with('-') {
                    ok = pts <= need;
                } else if let Some((_, hi)) = hcps.split_once('-') {
                    ok = pts >= need && pts <= ival(hi);
                } else {
                    ok = pts == need;
                }
            } else {
                let list: &[&str] = if hcps.starts_with("ab") {
                    &["AK", "AQ", "KQ", "AJT", "KJT", "QJT", "AKQ"]
                } else if hcps.starts_with("bc") {
                    &["AJT", "KJT", "QJT", "KQJ", "AQJ", "AKQ", "AKJ", "AKQJ"]
                } else if hcps.starts_with('a') {
                    &["AK", "AQ", "KQ"]
                } else if hcps.starts_with('b') {
                    &["AJT", "KJT", "QJT"]
                } else if hcps.starts_with('c') {
                    &["AKQ", "AKJ", "AQJ", "KQJ"]
                } else if hcps.starts_with('d') {
                    &["AKQJ", "KQJT", "AKQT", "AKQJT"]
                } else {
                    &[]
                };
                for (k, combo) in list.iter().enumerate() {
                    // In the VB, any combination before the last that starts
                    // with the ace marks the holding as better than needed.
                    if k + 1 < list.len() && combo.starts_with('A') {
                        extra = true;
                    }
                    if suit.starts_with(combo) {
                        ok = true;
                        break;
                    }
                }
            }
        }
        if (ok || hcps.is_empty()) && intermed_reqd > 0 {
            if extra {
                intermed_reqd -= 1;
            }
            ok = hand.intermediates[s] >= intermed_reqd;
        }
        ok
    }

    fn other_hcp(&self, t: &str) -> i32 {
        match t.chars().next().map(|c| suit_index(c.to_ascii_uppercase())) {
            Some(o) if o > 0 && t.len() == 1 => self.hand.hcp[o],
            _ => ival(t),
        }
    }

    /// VB "P" spec: honours protected by the bidding. Kx, Qx or AQ is not
    /// protected when LHO bid the suit, or (4th seat) when RHO didn't.
    fn protected(&self, s: usize) -> bool {
        let suit = &self.hand.suit[s];
        let l1 = suit.get(..1).unwrap_or(suit);
        let l2 = suit.get(..2).unwrap_or(suit);
        let l3 = suit.get(..3).unwrap_or(suit);
        if !(l2 == "AQ" || l1 == "K" || l1 == "Q") || l3 == "KQJ" || l3 == "QJT" || l3 == "AQJ" {
            return true;
        }
        let a = self.auction;
        if a.bid_num == 4 {
            let last = |b: &str| b.chars().last().map_or(0, suit_index);
            if last(a.bid(1)) == s || last(a.bid(3)) != s {
                return false;
            }
        }
        true
    }

    /// VB `Combo_Ok`: `cmNum` (combined length, "SH>8", "CDH=10-12") and
    /// `cmHCP` (combined points, "HDC>11", or key-card specs for slam
    /// replies).
    fn combo(&self, r: &Rule) -> Check {
        let num = r.cm_num.trim();
        let pts = r.cm_hcp.trim();
        if !num.is_empty() {
            if let Some((suits, op, lo, hi)) = self.parse_combo(num) {
                let have: i32 = suits.iter().map(|&s| self.hand.len[s]).sum();
                if !combo_holds(have, op, lo, hi) {
                    return Err(format!("combined length {have} fails {num}"));
                }
            }
        }
        if pts.is_empty() {
            return Ok(());
        }
        if let Some((suits, op, lo, hi)) = self.parse_combo(pts) {
            let have: i32 = suits.iter().map(|&s| self.hand.hcp[s]).sum();
            return if combo_holds(have, op, lo, hi) {
                Ok(())
            } else {
                Err(format!("combined HCP {have} fails {pts}"))
            };
        }
        self.slam_spec(pts, r)
    }

    /// "SH>8" -> (suits, '>', 8, 8); "CDH=10-12" -> (suits, '=', 10, 12).
    /// An "x" stands for the x suit; with no x suit the spec is skipped.
    fn parse_combo(&self, spec: &str) -> Option<(Vec<usize>, char, i32, i32)> {
        let letters: String = spec.chars().take_while(|c| "CDHSx".contains(*c)).collect();
        if !(2..=3).contains(&letters.len()) {
            return None;
        }
        let rest = &spec[letters.len()..];
        let op = rest.chars().next().filter(|c| "<>=".contains(*c))?;
        let body = &rest[1..];
        let lo = ival(body);
        let hi = body.split_once('-').map_or(lo, |(_, h)| ival(h));
        let mut suits = Vec::new();
        for c in letters.chars() {
            let s = if c == 'x' { self.x_suit } else { suit_index(c) };
            if s == 0 {
                return None;
            }
            suits.push(s);
        }
        Some((suits, op, lo, hi))
    }

    /// Ace and key-card replies in `cmHCP`, alternatives split by commas:
    /// "2A+TQ,1A+TK+TQ,4A+TK+TQ" is two or five key cards with the trump
    /// queen. "nA" is the number of aces (0 when left out), "TK" and "TQ"
    /// the trump king and queen, "+" required and "-" refused. A trump
    /// king not mentioned must be absent (it would be another key card);
    /// a queen not mentioned may be either. "kc1,3" is a key-card count.
    ///
    /// The 2018 engine read these only behind an "SB:" prefix, which no
    /// row has, so it passed every hand.
    fn slam_spec(&self, spec: &str, r: &Rule) -> Check {
        let spec = spec.strip_prefix("SB:").unwrap_or(spec);
        let trump = self.trump(r);
        let aces = self.hand.aces();
        let (tk, tq) = if trump > 0 {
            (self.hand.held[trump][KING], self.hand.held[trump][QUEEN])
        } else {
            (false, false)
        };
        if let Some(list) = spec.strip_prefix("kc") {
            if trump == 0 {
                return Err("key cards need a trump suit".into());
            }
            let kc = aces + i32::from(tk);
            return if list
                .split(',')
                .any(|n| ival(n) == kc && !n.trim().is_empty())
            {
                Ok(())
            } else {
                Err(format!("{kc} key cards, needs {list}"))
            };
        }
        let needs_trump = spec.contains("TK") || spec.contains("TQ");
        if needs_trump && trump == 0 {
            return Err("ace-asking reply needs a trump suit".into());
        }
        for alt in spec.split(',').map(str::trim).filter(|a| !a.is_empty()) {
            let want_aces = alt
                .find('A')
                .map(|i| ival(alt[..i].trim_start_matches(['+', '-'])))
                .unwrap_or(0);
            if aces != want_aces {
                continue;
            }
            let want_k = alt.contains("TK") && !alt.contains("-TK");
            if tk != want_k {
                continue;
            }
            if alt.contains("-TQ") && tq {
                continue;
            }
            if (alt.contains("+TQ") || alt.starts_with("TQ")) && !tq {
                continue;
            }
            return Ok(());
        }
        Err(format!("{aces} aces, trump K {tk}, Q {tq}: none of {spec}"))
    }

    /// VB `Eval_Pts_Ok`: points under a named valuation system from the
    /// hand-evaluation database (BB-HEP). Not ported; only three rules
    /// name a system together with the points it must reach.
    fn valuation(&mut self, r: &Rule) -> Check {
        if !r.val_system.trim().is_empty()
            && (!r.suit_val.trim().is_empty() || !r.nt_val.trim().is_empty())
        {
            return Err(format!("valuation system {:?} not ported", r.val_system));
        }
        Ok(())
    }
}

fn eq_or(have: i32, want: i32, what: &str) -> Check {
    if have == want {
        Ok(())
    } else {
        Err(format!("{have} {what}, needs {want}"))
    }
}

/// A count spec read as the VB reads Winrs, Losers, QTs and Stopped_Suits:
/// "Any value not preceded by '<' is considered a minimum, not a specific
/// value" (VB `CustPoints_Ok`), so "8", "8+" and "8-9" all mean 8 or more.
/// "<n" is fewer than n (the VB read it so only for Losers).
fn min_spec(spec: &str, have: f64, what: &str) -> Check {
    if spec.is_empty() {
        return Ok(());
    }
    let ok = match spec.strip_prefix('<') {
        Some(t) => have < val(t),
        None => have >= val(spec),
    };
    if ok {
        Ok(())
    } else {
        Err(format!("{have} {what}, needs {spec}"))
    }
}

/// The HCP range a rule's `HCP` spec (lower-cased, not an ace or key-card
/// count) stands for after `prior` bids. "Min", "HCP1N" and "P24-25"
/// (partnership total) read from the notrump ranges.
pub fn hcp_bounds(spec: &str, prior: &str, bid: &str, nt: NtRanges) -> (i32, i32) {
    let (min_nt, max_nt) = if prior.starts_with("1N")
        || (spec.starts_with("min") && bid.starts_with("1N"))
        || spec.starts_with("hcp1n")
    {
        (nt.one_min, nt.one_max)
    } else if prior.starts_with("2N") || spec.starts_with("min") || spec.starts_with("hcp2n") {
        (nt.two_min, nt.two_max)
    } else {
        (0, 0)
    };
    if spec == "min" {
        (min_nt, min_nt)
    } else if let Some(rest) = spec.strip_prefix("min") {
        match ival(rest) {
            0 => (min_nt + 1, max_nt),
            n => (min_nt + n, min_nt + n),
        }
    } else if spec == "hcp1n" || spec == "hcp2n" {
        (min_nt, max_nt)
    } else if let Some(rest) = spec.strip_prefix('p') {
        // Partnership total: "P24-25" opposite 15-17 is 9-10. The VB
        // tested for "P" after lower-casing the spec, so it never fired
        // and "P24-25" read as 0-25.
        let (lo, hi) = int_range(rest, 40);
        let lo = if lo > 0 { lo - min_nt } else { 0 };
        let hi = if hi < 40 { hi - min_nt } else { 40 };
        (lo.max(0), hi)
    } else {
        int_range(spec, 40)
    }
}

/// A points range as the HCP and Dist_Pts fields write it: "12" (exactly),
/// "12+", "12-14", "<12", ">12", "12-" (up to 12). `top` caps open ranges.
pub fn int_range(spec: &str, top: i32) -> (i32, i32) {
    let spec = spec.trim();
    let n = ival(spec);
    if let Some(t) = spec.strip_prefix('<') {
        (0, ival(t) - 1)
    } else if let Some(t) = spec.strip_prefix('>') {
        // The VB read ">12" as 0-40 (`Val(">12")` is 0).
        (ival(t) + 1, top)
    } else if spec.ends_with('+') {
        (n, top)
    } else if let Some((_, hi)) = spec.split_once('-') {
        if hi.is_empty() {
            (0, n)
        } else {
            (n, ival(hi))
        }
    } else {
        (n, n)
    }
}

fn compare_op(op: &str, a: i32, b: i32) -> bool {
    match op {
        "<" => a < b,
        ">" => a > b,
        "<=" | "=<" => a <= b,
        ">=" | "=>" => a >= b,
        "=" => a == b,
        _ => true,
    }
}

fn combo_holds(have: i32, op: char, lo: i32, hi: i32) -> bool {
    match op {
        '=' => have >= lo && have <= hi,
        '<' => have < lo,
        '>' => have > lo,
        _ => true,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ranges() {
        assert_eq!(int_range("12", 40), (12, 12));
        assert_eq!(int_range("12+", 40), (12, 40));
        assert_eq!(int_range("12-14", 40), (12, 14));
        assert_eq!(int_range("<12", 40), (0, 11));
        assert_eq!(int_range(">12", 40), (13, 40));
        assert_eq!(int_range("6-", 40), (0, 6));
    }

    #[test]
    fn minimum_specs() {
        assert!(min_spec("7.5", 8.0, "w").is_ok());
        assert!(min_spec("7.5", 7.0, "w").is_err());
        assert!(min_spec("<6", 5.0, "l").is_ok());
        assert!(min_spec("2-3", 4.0, "l").is_ok());
        assert!(min_spec("2-3", 1.0, "l").is_err());
    }
}
