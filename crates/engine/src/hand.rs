//! Hand analysis: everything about the hand that does not depend on the
//! rule being tested, computed once per lookup (VB `AnalyzeHand`).
//!
//! Arrays are indexed as in the VB source: 1 = clubs ... 4 = spades, and
//! index 0 holds the whole-hand total. Card ranks are 2..=14 (ace = 14).

// The branches follow the VB routines case by case, including cases that
// give the same answer, so each can be checked against the original.
#![allow(clippy::if_same_then_else)]

use std::fmt;

pub const ACE: usize = 14;
pub const KING: usize = 13;
pub const QUEEN: usize = 12;
pub const JACK: usize = 11;
pub const TEN: usize = 10;

const SUIT_NAMES: [&str; 5] = ["", "Clubs", "Diamonds", "Hearts", "Spades"];

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum HandError {
    /// Error 3: not 13 cards in "S-H-D-C" form.
    NotThirteen(String),
    /// Error 4: a character that is not a card.
    BadCard(String),
    /// Error 5: cards not written high to low.
    OutOfOrder(String),
}

impl HandError {
    /// The VB `ErrCode` for this error.
    pub fn code(&self) -> i32 {
        match self {
            HandError::NotThirteen(_) => 3,
            HandError::BadCard(_) => 4,
            HandError::OutOfOrder(_) => 5,
        }
    }
}

impl fmt::Display for HandError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            HandError::NotThirteen(m) | HandError::BadCard(m) | HandError::OutOfOrder(m) => {
                f.write_str(m)
            }
        }
    }
}

impl std::error::Error for HandError {}

/// The analysed hand.
#[derive(Debug, Clone)]
pub struct HandFacts {
    /// The hand as given, "AK432-64-QJT3-T2" (spades first).
    pub text: String,
    /// Cards of each suit, high to low, e.g. "AJ732".
    pub suit: [String; 5],
    pub held: [[bool; 15]; 5],
    pub len: [i32; 5],
    pub hcp: [i32; 5],
    /// 9s and 10s, and 8s that come with a 9 or 10.
    pub intermediates: [i32; 5],
    pub winners: [f64; 5],
    pub runners: [f64; 5],
    /// Losing-trick count.
    pub losers: [i32; 5],
    pub quick_tricks: [f64; 5],
    /// Stoppers per suit; index 0 is the number of suits stopped
    /// (half-stopped suits count a half).
    pub stops: [f64; 5],
    /// 1 doubleton, 2 singleton, 3 void; index 0 is the total.
    pub distrib: [i32; 5],
    /// "1", "2", "3" (suiters), "b" (balanced) or "" (other).
    pub shape: &'static str,
    /// "v" void, "s" singleton (and no void), or "".
    pub shortness: &'static str,
}

impl HandFacts {
    /// Analyse a hand. `lho_suit` (1-4, 0 = none) is LHO's suit, which
    /// costs some holdings a stopper; the 2018 engine never filled it in
    /// on the bidding path, so callers pass 0 for now.
    pub fn analyze(hand: &str, lho_suit: usize) -> Result<Self, HandError> {
        let mut h = HandFacts {
            text: hand.to_string(),
            suit: Default::default(),
            held: [[false; 15]; 5],
            len: [0; 5],
            hcp: [0; 5],
            intermediates: [0; 5],
            winners: [0.0; 5],
            runners: [0.0; 5],
            losers: [0; 5],
            quick_tricks: [0.0; 5],
            stops: [0.0; 5],
            distrib: [0; 5],
            shape: "",
            shortness: "",
        };
        h.read_cards()?;
        h.count_winners();
        h.count_losers();
        h.count_quick_tricks();
        h.count_stoppers(lho_suit);
        h.classify_shape();
        Ok(h)
    }

    pub fn aces(&self) -> i32 {
        (1..=4).filter(|&s| self.held[s][ACE]).count() as i32
    }

    pub fn kings(&self) -> i32 {
        (1..=4).filter(|&s| self.held[s][KING]).count() as i32
    }

    fn starts(&self, s: usize, prefix: &str) -> bool {
        self.suit[s].starts_with(prefix)
    }

    /// VB `ConvertHandStringToArrays`.
    fn read_cards(&mut self) -> Result<(), HandError> {
        let hand = self.text.clone();
        if hand.chars().count() != 16 || hand.split('-').count() != 4 {
            return Err(HandError::NotThirteen(format!(
                "Hand doesn't have 13 cards: {hand}"
            )));
        }
        for (k, part) in hand.split('-').enumerate() {
            let s = 4 - k; // spades first
            let mut last = 99;
            for c in part.chars() {
                let n = " 23456789TJQKA".find(c).unwrap_or(0);
                if n < 1 || c == ' ' {
                    return Err(HandError::BadCard(format!(
                        "{c} in {} in {hand} is not a valid card.",
                        SUIT_NAMES[s]
                    )));
                }
                let rank = n + 1;
                if rank >= last {
                    return Err(HandError::OutOfOrder(format!(
                        "The {c} of {} is not in sequence in \"{}{c}\"",
                        SUIT_NAMES[s], self.suit[s]
                    )));
                }
                self.hcp[s] += "JQKA".find(c).map_or(0, |i| i as i32 + 1);
                if c == '9' || c == 'T' {
                    self.intermediates[s] += 1;
                } else if c == '8' && (self.held[s][9] || self.held[s][10]) {
                    self.intermediates[s] += 1;
                }
                self.held[s][rank] = true;
                self.suit[s].push(c);
                last = rank;
            }
        }
        for s in 1..=4 {
            self.len[s] = self.suit[s].len() as i32;
            self.hcp[0] += self.hcp[s];
            self.intermediates[0] += self.intermediates[s];
        }
        self.len[0] = 13;
        Ok(())
    }

    /// VB `Calculate_Winners`: AKQJ=4, AKQ=3, AKJT=3.5, AKJ=2.5, AK=2,
    /// AQJT=3.5, AQJ=2.5, AQ=1.5, AJT=2, A=1, KQJT=3, KQJ=2, KQ=1, Kx=.5,
    /// QJT=1 (2 with length), plus one per card past four.
    fn count_winners(&mut self) {
        for s in 1..=4 {
            let len = self.len[s];
            let (mut w, mut r) = (0.0, 0.0);
            if self.starts(s, "AKQJ") {
                w = 4.0;
                r = len as f64;
            } else if self.starts(s, "AKQ") {
                w = 3.0;
                r = if len > 6 { len as f64 } else { 3.0 };
            } else if self.starts(s, "AK") {
                r = if len > 7 { len as f64 } else { 2.0 };
                w = if self.starts(s, "AKJT") || self.starts(s, "AKQT") {
                    3.5
                } else if self.starts(s, "AKJ") {
                    2.5
                } else {
                    2.0
                };
            } else if self.starts(s, "A") {
                r = if len > 10 { len as f64 } else { 1.0 };
                w = if self.starts(s, "AQJT") {
                    3.5
                } else if self.starts(s, "AQJ") {
                    2.5
                } else if self.starts(s, "AQ") {
                    1.5
                } else if self.starts(s, "AJT") {
                    2.0
                } else {
                    1.0
                };
            } else if self.starts(s, "KQJT") {
                w = 3.0;
            } else if self.starts(s, "KQJ") {
                w = 2.0;
            } else if self.starts(s, "KQ") {
                w = 1.0;
            } else if self.starts(s, "K") && len > 1 {
                w = 0.5;
            } else if self.starts(s, "QJT") {
                w = if len > 3 { 2.0 } else { 1.0 };
            }
            if len > 4 {
                w += (len - 4) as f64;
            }
            self.winners[s] = w;
            self.runners[s] = r;
            self.winners[0] += w;
            self.runners[0] += r;
        }
    }

    /// VB `Calculate_LosingTrickCount` (standard LTC).
    fn count_losers(&mut self) {
        for s in 1..=4 {
            let len = self.len[s];
            let l = if self.starts(s, "AKQ") {
                0
            } else if self.starts(s, "AK") {
                i32::from(len > 2)
            } else if self.starts(s, "AQ") {
                1
            } else if self.starts(s, "A") {
                if len > 2 {
                    2
                } else if len == 2 {
                    1
                } else {
                    0
                }
            } else if self.starts(s, "KQJ") {
                1
            } else if self.starts(s, "KQ") || self.starts(s, "K") {
                if len > 2 {
                    2
                } else if len == 2 {
                    1
                } else if self.starts(s, "KQ") {
                    0
                } else {
                    // A singleton K: the VB adds nothing here, so K alone
                    // counts no loser. Kept as written.
                    0
                }
            } else if self.starts(s, "Q") {
                if len > 2 {
                    2
                } else {
                    len
                }
            } else if len > 3 {
                3
            } else {
                len
            };
            self.losers[s] = l;
            self.losers[0] += l;
        }
    }

    /// VB `Calculate_QuickTricks`: AK 2, AQ 1.5, A 1, KQ 1, Kx .5; AK, AQ
    /// and KQ lose a half in a suit of six or more.
    fn count_quick_tricks(&mut self) {
        for s in 1..=4 {
            let adj = if self.len[s] > 5 { 0.5 } else { 0.0 };
            let q = if self.starts(s, "AK") {
                2.0 - adj
            } else if self.starts(s, "AQ") {
                1.5 - adj
            } else if self.starts(s, "A") {
                1.0
            } else if self.starts(s, "KQ") {
                1.0 - adj
            } else if self.starts(s, "K") && self.len[s] > 1 {
                0.5
            } else {
                0.0
            };
            self.quick_tricks[s] = q;
            self.quick_tricks[0] += q;
        }
    }

    /// VB `Calculate_Stoppers`: notrump stoppers per suit (0, .5, 1 or 2)
    /// and the number of suits stopped.
    fn count_stoppers(&mut self, lho_suit: usize) {
        for s in 1..=4 {
            let has = |r: usize| self.held[s][r];
            let len = self.len[s];
            let pad = format!("{}    ", self.suit[s]);
            let (c2, c3, c4) = (&pad[..2], &pad[..3], &pad[..4]);
            let lho = lho_suit == s;
            let mut n = 0.0;
            if has(ACE) {
                n = if c2 == "AK" || c2 == "AQ" {
                    2.0
                } else if (c3 == "AJT" || c3 == "AJ9") && len > 3 {
                    2.0
                } else if c4 == "AT98" && len > 4 {
                    2.0
                } else {
                    1.0
                };
            } else if has(KING) && len > 1 {
                n = if c2 == "KQ" {
                    if c3 == "KQJ" || c3 == "KQT" || (c3 == "KQ9" && len > 3) {
                        2.0
                    } else if lho {
                        1.0
                    } else {
                        2.0
                    }
                } else if has(JACK) {
                    if len > 3 {
                        if has(TEN) || has(9) || has(8) {
                            2.0
                        } else {
                            1.0
                        }
                    } else if lho || len < 3 {
                        1.0
                    } else {
                        2.0
                    }
                } else if has(TEN) && len > 3 {
                    if has(9) || has(8) {
                        2.0
                    } else {
                        1.0
                    }
                } else {
                    1.0
                };
            } else if has(QUEEN) && len > 2 {
                if has(JACK) {
                    n = 1.0;
                } else if has(TEN) && len > 3 {
                    if has(9) {
                        n = if lho { 1.0 } else { 2.0 };
                    }
                    if has(8) && has(7) {
                        n = if lho { 1.0 } else { 2.0 };
                    }
                    if has(7) {
                        n = 1.0;
                    }
                } else if has(9) && len > 3 {
                    if has(TEN) || has(8) {
                        n = 1.0;
                    }
                } else {
                    n = if lho { 0.0 } else { 1.0 };
                }
            } else if has(JACK) && len > 3 {
                n = if has(TEN) && (has(9) || has(8)) {
                    1.0
                } else if c3 == "J98" && len > 4 {
                    1.0
                } else {
                    0.5
                };
            } else if (c3 == "T98" || c3 == "T97") && len > 4 {
                n = 1.0;
            }
            self.stops[s] = n;
            if n == 0.5 {
                self.stops[0] += 0.5;
            } else if n > 0.0 {
                self.stops[0] += 1.0;
            }
        }
    }

    /// VB `Calculate_Shape`.
    fn classify_shape(&mut self) {
        let (mut n4, mut n5, mut n6, mut singles, mut voids) = (0, 0, 0, 0, 0);
        for s in 1..=4 {
            match self.len[s] {
                l if l > 5 => n6 += 1,
                5 => n5 += 1,
                4 => n4 += 1,
                2 => self.distrib[s] += 1,
                1 => {
                    singles += 1;
                    self.distrib[s] += 2
                }
                0 => {
                    voids += 1;
                    self.distrib[s] += 3
                }
                _ => {}
            }
            self.distrib[0] += self.distrib[s];
        }
        self.shape = if n6 == 1 {
            if n5 == 1 || n4 == 1 {
                "2"
            } else {
                "1"
            }
        } else if n5 == 2 || n6 == 2 || (n5 == 1 && n6 == 1) {
            "2"
        } else if n5 + n4 == 3 {
            "3"
        } else if n6 == 0 && self.distrib[0] < 2 {
            "b"
        } else {
            ""
        };
        self.shortness = if voids > 0 {
            "v"
        } else if singles > 0 {
            "s"
        } else {
            ""
        };
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn h(s: &str) -> HandFacts {
        HandFacts::analyze(s, 0).unwrap()
    }

    #[test]
    fn counts_points_and_lengths() {
        let f = h("AK432-64-QJT3-T2");
        assert_eq!(f.hcp, [10, 0, 3, 0, 7]);
        assert_eq!(f.len, [13, 2, 4, 2, 5]);
        assert_eq!(f.suit[4], "AK432");
        assert_eq!(f.intermediates[0], 2);
    }

    #[test]
    fn reads_voids() {
        let f = h("K--AKQJ632-AKT52");
        assert_eq!(f.len[3], 0);
        assert_eq!(f.shortness, "v");
        assert_eq!(f.shape, "2");
    }

    #[test]
    fn rejects_bad_hands() {
        assert_eq!(HandFacts::analyze("AK2-AK2-AK2", 0).unwrap_err().code(), 3);
        assert_eq!(
            HandFacts::analyze("AK2-AK2-AK2-K43Z", 0)
                .unwrap_err()
                .code(),
            4
        );
        assert_eq!(
            HandFacts::analyze("AK2-AK2-AK2-4K32", 0)
                .unwrap_err()
                .code(),
            5
        );
    }

    #[test]
    fn winners_losers_and_quick_tricks_follow_the_vb_tables() {
        let f = h("AKQJ-AQ2-K32-432");
        assert_eq!(f.winners[4], 4.0);
        assert_eq!(f.winners[3], 1.5);
        assert_eq!(f.winners[2], 0.5);
        assert_eq!(f.losers, [6, 3, 2, 1, 0]);
        assert_eq!(f.quick_tricks[0], 2.0 + 1.5 + 0.5);
    }

    #[test]
    fn shapes() {
        assert_eq!(h("AK2-AK2-AK2-K432").shape, "b");
        assert_eq!(h("AK32-AK2-AK2-K32").shape, "b");
        assert_eq!(h("AK432-AK2-AK2-K2").shape, "b");
        assert_eq!(h("AK432-AK32-AK-K2").shape, "");
        assert_eq!(h("AK5432-AK2-AK2-K").shape, "1");
        assert_eq!(h("AK54-AK32-AK32-K").shape, "3");
        assert_eq!(h("AK54-AK32-AK2-K2").shape, "b");
    }

    #[test]
    fn stoppers() {
        let f = h("AQ2-K2-QJ2-65432");
        assert_eq!(f.stops[4], 2.0);
        assert_eq!(f.stops[3], 1.0);
        assert_eq!(f.stops[2], 1.0);
        assert_eq!(f.stops[1], 0.0);
        assert_eq!(f.stops[0], 3.0);
    }
}
