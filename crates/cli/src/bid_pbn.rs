//! `bidbase bid-pbn`: bid every deal of a PBN file with the BidBase port
//! at all four seats and write it back in the layout bba-cli and
//! `rbb bid-pbn` write (`[Auction]` rows, `[Declarer]`, `[Contract]`, a
//! `% CC1` / `% CC2` header), so the same tools read all three.
//!
//! The client side follows the 2018 practice program: when no rule fits,
//! the bidder passes. One addition: BidBase files its replies to 4NT
//! under "RKCB: 4N-P" and "Blackwood: 4N-P" (the prior bids of the ask
//! alone), which nothing in the 2018 release reaches; after partner's 4NT
//! with no rule for the real auction, those keys are tried, with our
//! side's last suit as trumps.

use bidbase_engine::convert::{call_from_bidbase, call_to_bidbase, hand_to_bidbase, vul_for};
use bidbase_engine::{Bidder, LookupError, Request, Scoring};
use bridge_encodings::pbn::PbnDocument;
use bridge_types::{Auction, Call, Direction};

#[derive(Debug, Default)]
pub struct Stats {
    pub deals: usize,
    pub calls: usize,
    /// Calls where no rule fitted, so the bidder passed.
    pub no_rule: usize,
    /// Of those, calls in an auction with no rules at all (BidBase doesn't
    /// know the sequence); the rest had rules the hand didn't fit.
    pub unknown_auction: usize,
    /// Calls a rule made that were not legal at that point (made Pass).
    pub illegal: usize,
    /// Replies to 4NT found under the "RKCB:" / "Blackwood:" keys.
    pub slam_keys: usize,
    /// Auctions cut off at `MAX_CALLS`.
    pub runaway: usize,
}

const MAX_CALLS: usize = 60;
const SEATS: [Direction; 4] = [
    Direction::North,
    Direction::East,
    Direction::South,
    Direction::West,
];

pub fn bid_pbn(
    input: &str,
    bidder: &mut Bidder,
    ns_card: &str,
    ew_card: &str,
) -> Result<(String, Stats), String> {
    let mut doc = PbnDocument::parse(input).map_err(|e| e.to_string())?;
    let boards = doc.boards().to_vec();
    let mut stats = Stats {
        deals: boards.len(),
        ..Stats::default()
    };
    for (idx, board) in boards.iter().enumerate() {
        let hands: Vec<String> = SEATS
            .iter()
            .map(|&d| hand_to_bidbase(board.deal.hand(d)))
            .collect();
        let dealer = board.dealer.unwrap_or(Direction::North);
        let mut auction = Auction::new(dealer);
        let mut tokens: Vec<String> = Vec::new();
        let (mut unknown, mut unfit) = (0, 0);
        while !auction.is_complete() {
            if tokens.len() >= MAX_CALLS {
                stats.runaway += 1;
                while !auction.is_complete() {
                    auction.add_call(Call::Pass);
                }
                break;
            }
            let seat = auction.next_caller();
            let req = Request {
                hand: hands[seat.to_index()].clone(),
                prior_calls: tokens.join("-"),
                vul: vul_for(board.vulnerable, seat),
                scoring: Scoring::Matchpoints,
                ..Default::default()
            };
            let mut no_rules = false;
            let found = match bidder.lookup(&req).result {
                Ok(found) => Some(found),
                Err(LookupError::NoRules(_)) => {
                    let found = slam_reply(bidder, &req, &tokens);
                    if found.is_some() {
                        stats.slam_keys += 1;
                    }
                    no_rules = found.is_none();
                    found
                }
                Err(_) => None,
            };
            stats.calls += 1;
            let call = match found.and_then(|f| call_from_bidbase(&f.call)) {
                Some(c) if auction.is_legal(&c) => c,
                Some(_) => {
                    stats.illegal += 1;
                    Call::Pass
                }
                None => {
                    stats.no_rule += 1;
                    if no_rules {
                        stats.unknown_auction += 1;
                        unknown += 1;
                    } else {
                        unfit += 1;
                    }
                    Call::Pass
                }
            };
            tokens.push(call_to_bidbase(&call));
            auction.add_call(call);
        }

        let e = |e: bridge_encodings::ParseError| format!("board {}: {e}", idx + 1);
        for name in ["North", "East", "South", "West"] {
            doc.set_tag(idx, name, "BidBase").map_err(e)?;
        }
        let (contract, declarer) = match auction.final_contract() {
            Some(fc) => (fc.to_pbn(), fc.declarer.to_char().to_string()),
            None => ("Pass".to_string(), String::new()),
        };
        doc.set_tag(idx, "Declarer", &declarer).map_err(e)?;
        doc.set_tag(idx, "Contract", &contract).map_err(e)?;
        // Passes made for want of a rule: in a sequence BidBase has no
        // rules for, and where its rules didn't fit the hand.
        doc.set_tag(idx, "BidBaseUnknownAuction", &unknown.to_string())
            .map_err(e)?;
        doc.set_tag(idx, "BidBaseNoFit", &unfit.to_string())
            .map_err(e)?;
        // They described another contract.
        doc.remove_tag(idx, "Result").map_err(e)?;
        doc.remove_tag(idx, "Score").map_err(e)?;
        let calls: Vec<String> = auction
            .calls
            .iter()
            .map(|c| c.call.to_pbn().replace("NT", "N"))
            .collect();
        let rows: Vec<String> = calls
            .chunks(4)
            .map(|row| {
                row.iter()
                    .map(|c| format!("{c:<5}"))
                    .collect::<String>()
                    .trim_end()
                    .to_string()
            })
            .collect();
        let rows: Vec<&str> = rows.iter().map(String::as_str).collect();
        doc.set_section(idx, "Auction", &dealer.to_char().to_string(), &rows)
            .map_err(e)?;
    }
    let header = format!(
        "% PBN 2.1\n% Generated by bidbase {}\n% CC1 - {ns_card}\n% CC2 - {ew_card}\n",
        env!("CARGO_PKG_VERSION")
    );
    Ok((header + &doc.to_pbn(), stats))
}

/// After partner's 4NT and RHO's call, look the reply up under the ace-
/// asking keys, with our side's last suit before 4NT as trumps.
fn slam_reply(
    bidder: &mut Bidder,
    req: &Request,
    tokens: &[String],
) -> Option<bidbase_engine::Found> {
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
            [b'1'..=b'7', s] => Some(bidbase_engine::vb::suit_index(*s as char)).filter(|&x| x > 0),
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

/// One call of a replayed auction: what the file's engine called and what
/// BidBase would call in the same position.
pub struct Replayed {
    pub board: usize,
    /// The calls before this one, BidBase style ("P-1C-1N").
    pub prior: String,
    pub theirs: String,
    /// BidBase's call, `None` when no rule fits (it would pass).
    pub ours: Option<String>,
    /// Whether BidBase has any rules for the sequence.
    pub known: bool,
    pub rule_id: Option<i64>,
}

/// Walk every auction in a PBN file (another engine's) and ask BidBase for
/// each call in the same position: the auction so far, the caller's hand
/// and vulnerability.
pub fn replay(input: &str, bidder: &mut Bidder) -> Result<Vec<Replayed>, String> {
    let doc = PbnDocument::parse(input).map_err(|e| e.to_string())?;
    let mut out = Vec::new();
    for (idx, board) in doc.boards().iter().enumerate() {
        let Some(auction) = &board.auction else {
            continue;
        };
        let hands: Vec<String> = SEATS
            .iter()
            .map(|&d| hand_to_bidbase(board.deal.hand(d)))
            .collect();
        let mut seat = auction.dealer;
        let mut tokens: Vec<String> = Vec::new();
        for c in &auction.calls {
            let theirs = call_to_bidbase(&c.call);
            let req = Request {
                hand: hands[seat.to_index()].clone(),
                prior_calls: tokens.join("-"),
                vul: vul_for(board.vulnerable, seat),
                scoring: Scoring::Matchpoints,
                ..Default::default()
            };
            let (ours, known, rule_id) = match bidder.lookup(&req).result {
                Ok(f) => (Some(f.call), true, Some(f.rule_id)),
                Err(LookupError::NoRules(_)) => match slam_reply(bidder, &req, &tokens) {
                    Some(f) => (Some(f.call), true, Some(f.rule_id)),
                    None => (None, false, None),
                },
                Err(_) => (None, true, None),
            };
            out.push(Replayed {
                board: idx + 1,
                prior: tokens.join("-"),
                theirs: theirs.clone(),
                ours,
                known,
                rule_id,
            });
            tokens.push(theirs);
            seat = seat.next();
        }
    }
    Ok(out)
}
