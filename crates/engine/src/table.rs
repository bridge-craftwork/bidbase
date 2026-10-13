//! Bidding at a table: the client side of the engine, as Nelson's
//! practice program drives it. When no rule fits, the bidder passes.
//!
//! Two additions to the plain lookup:
//! - 2026 ace asking: a rule's bid can carry a code ("4N@30", 4NT with
//!   hearts agreed). The token stays in the auction, and the calls after
//!   it are looked up from the ask on ("4N@30-P") with the trump suit from
//!   the code, as the 2026 practice program passes them.
//! - 2018 slam keys: the 2018 rules file replies to 4NT under "RKCB: 4N-P"
//!   and "Blackwood: 4N-P", which nothing in that release reaches; after
//!   partner's 4NT with no rule for the real auction, those keys are tried.

use bridge_types::{Auction, Call, Direction, Vulnerability};

use crate::convert::{call_from_bidbase, call_to_bidbase, vul_for};
use crate::lookup::{Bidder, Found, LookupError, Request, Scoring};
use crate::vb::suit_index;

/// Why a seat passed for want of a rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Missing {
    /// BidBase has no rules for the auction so far.
    UnknownAuction,
    /// It has rules, but the hand fits none of them.
    NoFit,
    /// The rule's call was not legal at this point.
    Illegal,
}

/// One call at the table.
#[derive(Debug, Clone)]
pub struct TableCall {
    pub seat: Direction,
    pub call: Call,
    /// The token the auction carries ("1N", or "4N@30").
    pub token: String,
    /// The rule that made the call; `None` when the seat passed for want
    /// of one (`missing` says why).
    pub found: Option<Found>,
    pub missing: Option<Missing>,
    /// Found under a 2018 "RKCB:" / "Blackwood:" key.
    pub slam_key: bool,
}

/// The prior calls to look up and the trump suit: after an ace-asking
/// token ("4N@30") only the calls from the first such token on, with the
/// suit from its code. The first, as the practice program cuts it
/// (frm:6101): the 2026 rules key later asks from the first one
/// ("3S@CH-P-4N@30-P-4H").
pub fn key_for(tokens: &[String]) -> (String, usize) {
    match tokens.iter().position(|t| t.contains('@')) {
        Some(i) => {
            let code = tokens[i].split('@').nth(1).unwrap_or("");
            let trump = code
                .chars()
                .next()
                .and_then(|c| c.to_digit(10))
                .unwrap_or(0) as usize;
            (tokens[i..].join("-"), trump.min(4))
        }
        None => (tokens.join("-"), 0),
    }
}

/// After partner's 4NT and RHO's call, the reply under the 2018 ace-asking
/// keys, with our side's last suit before 4NT as trumps.
pub fn slam_reply(bidder: &mut Bidder, req: &Request, tokens: &[String]) -> Option<Found> {
    let n = tokens.len();
    if n < 2 || tokens[n - 2] != "4N" {
        return None;
    }
    // Our side's calls: partner at n-2, n-6, ...; ours at n-4, n-8, ...
    let trump = (2..=n)
        .step_by(2)
        .skip(1)
        .map(|back| tokens[n - back].as_str())
        .find_map(|t| match t.as_bytes() {
            [b'1'..=b'7', s] => Some(suit_index(*s as char)).filter(|&x| x > 0),
            _ => None,
        })
        .unwrap_or(0);
    for key in ["RKCB: ", "Blackwood: "] {
        let r = Request {
            prior_calls: format!("{key}4N-{}", tokens[n - 1]),
            trump,
            ..req.clone()
        };
        if let Ok(found) = bidder.lookup(&r).result {
            return Some(found);
        }
    }
    None
}

/// The next call: `hand` (BidBase form) for the seat now to call, after
/// `tokens` (the auction so far as BidBase tokens).
pub fn next_call(
    bidder: &mut Bidder,
    auction: &Auction,
    tokens: &[String],
    hand: &str,
    vulnerability: Vulnerability,
    scoring: Scoring,
    trace: bool,
) -> (TableCall, Vec<crate::lookup::Tried>) {
    let seat = auction.next_caller();
    let template = Request {
        hand: hand.to_string(),
        vul: vul_for(vulnerability, seat),
        scoring,
        trace,
        ..Default::default()
    };
    next_call_with(bidder, auction, tokens, template)
}

/// [`next_call`] from a request template: its hand, vulnerability,
/// scoring and flags (`disclose`, `trace`) are kept; the prior calls and
/// trump suit come from `tokens`.
pub fn next_call_with(
    bidder: &mut Bidder,
    auction: &Auction,
    tokens: &[String],
    template: Request,
) -> (TableCall, Vec<crate::lookup::Tried>) {
    let seat = auction.next_caller();
    let (prior_calls, trump) = key_for(tokens);
    let req = Request {
        prior_calls,
        trump,
        ..template
    };
    let out = bidder.lookup(&req);
    let tried = out.tried;
    let (found, mut missing, slam_key) = match out.result {
        Ok(f) => (Some(f), None, false),
        Err(LookupError::NoRules(_)) => match slam_reply(bidder, &req, tokens) {
            Some(f) => (Some(f), None, true),
            None => (None, Some(Missing::UnknownAuction), false),
        },
        Err(_) => (None, Some(Missing::NoFit), false),
    };
    let mut call = Call::Pass;
    let mut token = "P".to_string();
    let mut found = found;
    if let Some(f) = &found {
        match call_from_bidbase(&f.call) {
            Some(c) if auction.is_legal(&c) => {
                call = c;
                token = if f.token.contains('@') {
                    f.token.clone()
                } else {
                    call_to_bidbase(&call)
                };
            }
            _ => {
                missing = Some(Missing::Illegal);
                found = None;
            }
        }
    }
    (
        TableCall {
            seat,
            call,
            token,
            found,
            missing,
            slam_key,
        },
        tried,
    )
}

/// Bid a whole deal at all four seats. `hands` are BidBase strings in the
/// order N, E, S, W. Returns the auction and each call; an auction still
/// open after `max_calls` is finished with passes (`runaway`).
pub fn bid_deal(
    bidder: &mut Bidder,
    hands: &[String; 4],
    dealer: Direction,
    vulnerability: Vulnerability,
    scoring: Scoring,
    max_calls: usize,
) -> (Auction, Vec<TableCall>, bool) {
    let mut auction = Auction::new(dealer);
    let mut tokens: Vec<String> = Vec::new();
    let mut calls = Vec::new();
    while !auction.is_complete() {
        if tokens.len() >= max_calls {
            while !auction.is_complete() {
                auction.add_call(Call::Pass);
            }
            return (auction, calls, true);
        }
        let seat = auction.next_caller();
        let (tc, _) = next_call(
            bidder,
            &auction,
            &tokens,
            &hands[seat.to_index()],
            vulnerability,
            scoring,
            false,
        );
        tokens.push(tc.token.clone());
        auction.add_call(tc.call.clone());
        calls.push(tc);
    }
    (auction, calls, false)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn toks(s: &str) -> Vec<String> {
        s.split('-').map(String::from).collect()
    }

    #[test]
    fn key_starts_at_first_ask() {
        assert_eq!(key_for(&toks("1H-P-1S")), ("1H-P-1S".into(), 0));
        assert_eq!(
            key_for(&toks("1H-P-1S-2D-2S-P-4N@30-P")),
            ("4N@30-P".into(), 3)
        );
        assert_eq!(
            key_for(&toks("1S-P-3S@CH-P-4N@30-P-4H")),
            ("3S@CH-P-4N@30-P-4H".into(), 0)
        );
    }
}
