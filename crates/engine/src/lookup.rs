//! Finding a call (VB `LookUpBid`): take the rules filed under the
//! auction, in `Bid_Order_Number` order, and return the first one the
//! hand passes.

use std::fmt;

use crate::auction::AuctionFacts;
use crate::checks::Ctx;
use crate::hand::{HandError, HandFacts};
use crate::rules::{pct_number, Dialect, Rule, RuleBase};
use crate::vb::{suit_index, suit_letter};

/// Vulnerability as the client gives it. A rule's `Vul` field passes
/// when it names one of `tokens()`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum Vul {
    /// Not given: only rules with no `Vul` or `Vul = N` pass.
    #[default]
    Unknown,
    /// We are vulnerable, they are not.
    We,
    /// They are vulnerable, we are not.
    They,
    Neither,
    Both,
    /// One editor code (Y, N, Fav, Unfav, Eq, Eq+), as the BidBase Editor
    /// passes it when testing an entry.
    Spec(String),
}

impl Vul {
    /// The `Vul` codes this vulnerability satisfies. "Eq+" ("equal or
    /// favourable") is added for `They`, which the 2018 client mapping
    /// left out.
    pub fn tokens(&self) -> Vec<&str> {
        match self {
            Vul::Unknown => vec![],
            Vul::We => vec!["Y", "Unfav"],
            Vul::They => vec!["N", "Fav", "Eq+"],
            Vul::Neither => vec!["Eq", "N", "Eq+"],
            Vul::Both => vec!["Eq", "Y", "Eq+"],
            Vul::Spec(s) => vec![s.as_str()],
        }
    }

    /// The one-letter code the 2026 engine compares with a rule's `Vul`:
    /// U (we alone vulnerable), F (they alone), B (both), N (neither), ""
    /// (not given). The practice program passed "" for neither, so the
    /// rules Nelson marked N never matched; his abbreviations page
    /// defines N as neither side vulnerable (docs/ENGINE-2026.md).
    pub fn code_2026(&self) -> &str {
        match self {
            Vul::We => "U",
            Vul::They => "F",
            Vul::Both => "B",
            Vul::Neither => "N",
            Vul::Unknown => "",
            Vul::Spec(s) => s.as_str(),
        }
    }

    /// Parse "we", "they", "none"/"neither", "both", or an editor code.
    pub fn parse(s: &str) -> Option<Vul> {
        Some(match s.trim().to_ascii_uppercase().as_str() {
            "" => Vul::Unknown,
            "WE" | "US" => Vul::We,
            "THEY" | "THEM" => Vul::They,
            "NEITHER" | "NONE" => Vul::Neither,
            "BOTH" | "ALL" => Vul::Both,
            "Y" => Vul::Spec("Y".into()),
            "N" => Vul::Spec("N".into()),
            "FAV" => Vul::Spec("Fav".into()),
            "UNFAV" => Vul::Spec("Unfav".into()),
            "EQ" => Vul::Spec("Eq".into()),
            "EQ+" => Vul::Spec("Eq+".into()),
            // 2026 codes.
            c @ ("U" | "F" | "B" | "V" | "X") => Vul::Spec(c.into()),
            _ => return None,
        })
    }
}

/// Matchpoints or IMPs: rules with `Pct_Used` "M" or "I" apply only then.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Scoring {
    #[default]
    Unspecified,
    Matchpoints,
    Imps,
}

impl Scoring {
    fn code(self) -> &'static str {
        match self {
            Scoring::Unspecified => "",
            Scoring::Matchpoints => "M",
            Scoring::Imps => "I",
        }
    }
}

/// A question for the engine (the VB client properties).
#[derive(Debug, Clone, Default)]
pub struct Request {
    /// "AK432-64-QJT3-T2": spades, hearts, diamonds, clubs.
    pub hand: String,
    /// "P-1C-1N-2C-D-R-2S", opening passes included.
    pub prior_calls: String,
    pub vul: Vul,
    pub scoring: Scoring,
    /// Trump suit 1-4 (C..S) for ace-asking replies, 0 if unknown. A
    /// rule's `Source` ("<H...") overrides it.
    pub trump: usize,
    /// The BidBase Editor's entry check: the rule whose test hand this
    /// is. Its position is not checked and its percentage always applies.
    pub test_id: i64,
    /// Test only rule `test_id`, not every rule of the auction.
    pub test_entry: bool,
    /// Treat every rule as enabled whatever its `Pct_Used`.
    pub check_all_pcts: bool,
    /// Build the disclosure text for the call found.
    pub disclose: bool,
    /// Record why each rule was passed over.
    pub trace: bool,
}

/// Why a lookup found no call. Codes are the VB `ErrCode`s.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LookupError {
    /// 1: no rules for this auction (or no rule `test_id`).
    NoRules(String),
    /// 2: rules exist but the hand passes none of them.
    NoMatch,
    /// 3-5: the hand could not be read.
    Hand(HandError),
    /// 97: the rule under test does not accept its own test hand.
    TestHandFails,
}

impl LookupError {
    pub fn code(&self) -> i32 {
        match self {
            LookupError::NoRules(_) => 1,
            LookupError::NoMatch => 2,
            LookupError::Hand(e) => e.code(),
            LookupError::TestHandFails => 97,
        }
    }
}

impl fmt::Display for LookupError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            LookupError::NoRules(m) => f.write_str(m),
            LookupError::NoMatch => f.write_str("No (additional) entry found for this hand."),
            LookupError::Hand(e) => write!(f, "{e}"),
            LookupError::TestHandFails => f.write_str("Test hand didn't match entry's specs."),
        }
    }
}

impl std::error::Error for LookupError {}

/// One rule tried, and why it failed (`None` for the rule that matched).
#[derive(Debug, Clone)]
pub struct Tried {
    pub rule_id: i64,
    pub bid: String,
    pub bid_order: String,
    pub failed: Option<(&'static str, String)>,
}

/// The answer.
#[derive(Debug, Clone)]
pub struct Found {
    /// The call, with "x" replaced by the suit ("2x" -> "2H").
    pub call: String,
    /// The rule's bid as the auction carries it: the call plus a 2026
    /// ace-asking code ("4N@30" asks with hearts as trumps). The calls
    /// after it are looked up from this token on ("4N@30-P").
    pub token: String,
    pub rule_id: i64,
    /// What the call shows by the rule that made it (`meaning::describe`).
    pub meaning: String,
    /// With `disclose`: every enabled rule for this auction that makes the
    /// call, with its meaning. That is what partner and the opponents may
    /// know: which of them matched is the bidder's secret (VB `GetDiscl`).
    pub meanings: Vec<(i64, String)>,
    /// With `disclose`: the 2018 engine's disclosure text.
    pub disclosure: String,
}

#[derive(Debug, Clone)]
pub struct Outcome {
    pub result: Result<Found, LookupError>,
    pub tried: Vec<Tried>,
}

/// The engine: a rule base plus the random state `Pct_Used` draws from.
pub struct Bidder {
    pub base: RuleBase,
    rng: u64,
}

impl Bidder {
    pub fn new(base: RuleBase) -> Self {
        Bidder {
            base,
            rng: 0x9E37_79B9_7F4A_7C15,
        }
    }

    /// Seed the generator behind `Pct_Used` percentages.
    pub fn seed(&mut self, seed: u64) {
        self.rng = seed ^ 0x9E37_79B9_7F4A_7C15;
        if self.rng == 0 {
            self.rng = 1;
        }
    }

    pub fn lookup(&mut self, req: &Request) -> Outcome {
        let mut tried = Vec::new();
        let result = self.lookup_inner(req, &mut tried);
        Outcome { result, tried }
    }

    fn lookup_inner(
        &mut self,
        req: &Request,
        tried: &mut Vec<Tried>,
    ) -> Result<Found, LookupError> {
        let auction = AuctionFacts::new(&req.prior_calls);
        // LHO's suit stays unknown, as in the 2018 engine (PORTING.md).
        let hand = HandFacts::analyze_as(req.hand.trim(), 0, self.base.dialect)
            .map_err(LookupError::Hand)?;

        let list: Vec<usize> = if req.test_entry && req.test_id > 0 {
            match self.base.index_of(req.test_id) {
                Some(i) => vec![i],
                None => {
                    return Err(LookupError::NoRules(format!(
                        "Entry #{} not found.",
                        req.test_id
                    )))
                }
            }
        } else {
            self.base.for_auction(&auction.prior_bids).to_vec()
        };
        if list.is_empty() {
            return Err(LookupError::NoRules(format!(
                "No entries found matching the prior bids: {} (opening passes are ignored)",
                auction.prior_calls
            )));
        }

        let mut ctx = Ctx::new(&hand, &auction, req, self.base.nt, self.base.dialect);
        let mut k = 0;
        while k < list.len() {
            let rule = &self.base.rules[list[k]];
            if rule.is_section_end() {
                break;
            }
            let (next, verdict) = match self.pct_used(req, &auction, &list, k) {
                Err(why) => (k, Err(("Pct_Used", why))),
                Ok(at) => (at, ctx.run(&self.base.rules[list[at]])),
            };
            k = next;
            let rule = &self.base.rules[list[k]];
            if req.trace {
                tried.push(Tried {
                    rule_id: rule.id,
                    bid: rule.bid.clone(),
                    bid_order: rule.bid_order.clone(),
                    failed: verdict.clone().err(),
                });
            }
            match verdict {
                Ok(()) => {
                    let mut x_suit = ctx.x_suit;
                    if self.base.dialect == Dialect::V2026
                        && rule.bid.to_ascii_lowercase().contains('x')
                        && x_suit == 0
                    {
                        // 2026: an "x" bid with no x suit bids the longest
                        // suit (the lowest of equal ones).
                        let mut best = 0;
                        for s in 1..=4 {
                            if hand.len[s] > best {
                                best = hand.len[s];
                                x_suit = s;
                            }
                        }
                    }
                    let token = if rule.bid.to_ascii_lowercase().contains('x') {
                        format!("{}{}", &rule.bid[..1], suit_letter(x_suit))
                    } else {
                        rule.bid.trim().to_string()
                    };
                    let call = call_of_token(&token);
                    let x_suit = call.chars().nth(1).map_or(0, suit_index);
                    let meaning =
                        crate::meaning::describe(&self.base, rule, &auction.prior_bids, x_suit);
                    let (meanings, disclosure) = if req.disclose {
                        (
                            self.meanings(&auction, &call),
                            self.disclosure(&auction, &call),
                        )
                    } else {
                        (Vec::new(), String::new())
                    };
                    return Ok(Found {
                        call,
                        token,
                        rule_id: rule.id,
                        meaning,
                        meanings,
                        disclosure,
                    });
                }
                Err(_) if req.test_id > 0 && rule.id == req.test_id => {
                    return Err(LookupError::TestHandFails);
                }
                Err(_) => k += 1,
            }
        }
        Err(LookupError::NoMatch)
    }

    /// Whether the rule can be chosen at all in play: its `Pct_Used` is
    /// blank, a percentage, M or I, and its convention is on for a bidder
    /// at `pos_num`.
    pub fn is_enabled(&self, rule: &Rule, pos_num: i32) -> bool {
        let pct = rule.pct_used.trim();
        (pct.is_empty() || pct_number(pct) > 0.0 || pct == "M" || pct == "I")
            && self.base.convention_allows(rule, pos_num).is_ok()
    }

    /// VB `PctUsed_Ok`. Blank: always. "0", "x", "D": never (unless
    /// `check_all_pcts`). "M"/"I": only at that scoring. A number: that
    /// percent of the time, drawn over the run of numbered rules sharing a
    /// `Bid_Order_Number`; the last of the run takes what is left. Returns
    /// the index of the rule to test, which moves forward when the draw
    /// picks a later alternative.
    fn pct_used(
        &mut self,
        req: &Request,
        auction: &AuctionFacts,
        list: &[usize],
        k: usize,
    ) -> Result<usize, String> {
        let rules = &self.base.rules;
        let rule = &rules[list[k]];
        if req.test_entry {
            return Ok(k);
        }
        if self.base.dialect == Dialect::V2026 && req.test_id != 0 && rule.id == req.test_id {
            // 2026 `Activation_Ok`: the entry under test always passes.
            return Ok(k);
        }
        if rule.id != req.test_id {
            self.base.convention_allows(rule, auction.pos_num)?;
        }
        let pct = rule.pct_used.trim();
        let n = pct_number(pct);
        if req.test_id > 0 && n > 0.0 {
            return if rule.id == req.test_id {
                Ok(k)
            } else {
                Err(format!("Pct_Used {pct} (only the entry under test)"))
            };
        }
        if pct.is_empty() || req.check_all_pcts || pct == req.scoring.code() {
            return Ok(k);
        }
        if n > 0.0 {
            let draw = percent(&mut self.rng);
            let order = &rule.bid_order;
            let mut total = 0.0;
            let mut at = k;
            loop {
                let r = &rules[list[at]];
                if draw - total <= pct_number(&r.pct_used) {
                    return Ok(at);
                }
                total += pct_number(&r.pct_used);
                match list.get(at + 1).map(|&i| &rules[i]) {
                    Some(next) if &next.bid_order == order && pct_number(&next.pct_used) > 0.0 => {
                        at += 1
                    }
                    _ => return Ok(at),
                }
            }
        }
        let first = pct.get(..1).unwrap_or("");
        if (first == "M" || first == "I") && (first == req.scoring.code() || req.test_id > 0) {
            return Ok(k);
        }
        Err(format!("Pct_Used {pct:?}"))
    }

    /// Every enabled rule for the auction that makes `call`, with what it
    /// shows.
    pub fn meanings(&self, auction: &AuctionFacts, call: &str) -> Vec<(i64, String)> {
        let x_suit = call.chars().nth(1).map_or(0, suit_index);
        self.base
            .for_auction(&auction.prior_bids)
            .iter()
            .map(|&i| &self.base.rules[i])
            .take_while(|r| !r.is_section_end())
            .filter(|r| bid_makes(&r.bid, call) && self.is_enabled(r, auction.pos_num))
            .map(|r| {
                (
                    r.id,
                    crate::meaning::describe(&self.base, r, &auction.prior_bids, x_suit),
                )
            })
            .collect()
    }

    /// VB `GetDiscl`: what partner (and the opponents) may know about the
    /// call: every enabled rule for this auction that makes it, not just
    /// the one that matched.
    fn disclosure(&self, auction: &AuctionFacts, call: &str) -> String {
        if call == "P" {
            return String::new();
        }
        let mut out = String::new();
        for &i in self.base.for_auction(&auction.prior_bids) {
            let r = &self.base.rules[i];
            if r.is_section_end() {
                break;
            }
            let pct = r.pct_used.trim();
            let enabled = pct.is_empty() || pct_number(pct) != 0.0;
            if enabled && bid_makes(&r.bid, call) {
                self.add_disclosure(&mut out, r);
            }
        }
        out
    }

    /// VB `AddToDisclosure`.
    fn add_disclosure(&self, out: &mut String, r: &Rule) {
        let conv = self.base.conventions.get(&r.convention());
        let (sel, sel2) = conv.map_or(("", ""), |c| (c.selected.as_str(), c.sub_selected.as_str()));
        if sel == "0" || sel2 == "0" {
            return;
        }
        let name = conv.map_or("", |c| c.name.trim());
        out.push_str(&format!("{}:  BON:{}. ", r.bid, r.bid_order));
        if !name.is_empty() {
            out.push_str(&format!("\"{name}\". "));
        }
        if !r.pos.is_empty() {
            out.push_str(&format!("Position: {}. ", r.pos));
        }
        let strength = match r.strength_discl.as_str() {
            "V" => "Very Weak",
            "W" => "Weak",
            "I" => "Intermediate",
            "O" => "Opening",
            "S" => "Strong",
            "WS" => "Weak-Strong",
            "WI" => "Weak-Intermediate",
            "IS" | "I+" => "Intermediate-Strong",
            "W+" => "Weak-up",
            "S+" => "Strong+",
            "VS" => "Very Strong",
            _ => "",
        };
        out.push_str(&format!("\nStrength={strength}.  "));
        out.push_str(&format!(
            ".  x:{} S:{} H:{} D:{} C:{}\n",
            r.x_hcp, r.s_hcp, r.h_hcp, r.d_hcp, r.c_hcp
        ));
        out.push_str(&format!("Shape: {}.  ", r.shape));
        let lens = format!(
            "x:{} S:{} H:{} D:{} C:{}",
            r.x_len, r.s_len, r.h_len, r.d_len, r.c_len
        );
        if lens.len() == 14 {
            out.push_str(&format!(
                "x:{} S:{} H:{} D:{} C:{}\n",
                r.x_num, r.s_num, r.h_num, r.d_num, r.c_num
            ));
        } else {
            out.push_str(&lens);
            out.push('\n');
        }
        if let Some(asks) = asks_text(&r.asks) {
            out.push_str(&format!("Asks partner {asks}.\n"));
        }
        if pct_number(&r.pct_used) > 0.0 {
            out.push_str(&format!("   Pct Used: {}%\n", r.pct_used));
        }
        if r.alert != "D" && !r.explanation.is_empty() {
            let text = public_explanation(&r.explanation);
            if !text.is_empty() {
                let label = if r.alert == "A" { "ALERT: " } else { "NOTE:  " };
                out.push_str(&format!("{label}{text}\n"));
            }
        }
        if !r.source.trim().is_empty() {
            out.push_str(&format!("Source: '{}'", r.source.trim()));
        }
        out.push_str("\n\n");
    }
}

/// The call a rule's bid makes: "4N@30" is 4N, "2Hj" (a stray suffix)
/// 2H, "p" P.
pub fn call_of_token(token: &str) -> String {
    let t = token.split('@').next().unwrap_or("").trim();
    match t.chars().next() {
        Some(c) if c.is_ascii_digit() => t.chars().take(2).collect::<String>().to_ascii_uppercase(),
        _ => t.to_ascii_uppercase(),
    }
}

/// 1-100, like the VB `Fix(Rnd * 100 + 1)` (xorshift64*).
fn percent(state: &mut u64) -> f64 {
    let mut x = *state;
    x ^= x >> 12;
    x ^= x << 25;
    x ^= x >> 27;
    *state = x;
    let r = x.wrapping_mul(0x2545_F491_4F6C_DD1D) >> 11;
    ((r as f64 / (1u64 << 53) as f64) * 100.0).floor() + 1.0
}

/// Whether a rule's bid ("2H", or "2x" for any suit) makes `call`.
fn bid_makes(bid: &str, call: &str) -> bool {
    bid == call
        || (bid.ends_with('x')
            && bid.len() == 2
            && call.len() == 2
            && bid[..1] == call[..1]
            && call[1..].chars().next().is_some_and(|c| suit_index(c) > 0))
}

/// The explanation with "[private notes]" removed, as the VB showed it.
pub(crate) fn public_explanation(text: &str) -> String {
    let a = text.trim();
    match (a.find('['), a.find(']')) {
        (Some(open), Some(close)) if close > open => {
            if open > 0 {
                format!("{}{}", &a[..open], &a[close + 1..])
                    .trim()
                    .to_string()
            } else {
                a[close + 1..].trim().to_string()
            }
        }
        (Some(_), _) => String::new(),
        _ => a.to_string(),
    }
}

/// VB `AddToDisclosure`'s reading of the `Asks` field.
fn asks_text(asks: &str) -> Option<String> {
    let a = asks.to_ascii_lowercase();
    let suit_name = |c: Option<char>| match c {
        Some('c') => "Clubs",
        Some('d') => "Diamonds",
        Some('h') => "Hearts",
        Some('s') => "Spades",
        _ => "?",
    };
    let second = a.chars().nth(1);
    Some(match a.as_str() {
        "" => return None,
        "g" => "game force".into(),
        "1" => "bid again".into(),
        "poc" => "pass or correct".into(),
        "p" => "pick a suit/slam".into(),
        "a" => "show Aces".into(),
        "k" => "show Kings".into(),
        "kc" => "show # of Key-Cards".into(),
        _ if a.starts_with('h') => format!("for help in {}", suit_name(second)),
        _ if a.starts_with('c') => format!("for controls in {}", suit_name(second)),
        _ if a.starts_with('b') => match asks.trim().get(1..).unwrap_or("") {
            "m" => "to bid a minor".into(),
            "M" => "to bid a major".into(),
            "" => "to bid".into(),
            rest if rest.len() == 1 => format!("to bid {}", suit_name(second)),
            // "B2C", "B4S": a contract.
            rest => format!("to bid {rest}"),
        },
        _ => format!("to {}", asks.trim()),
    })
}
