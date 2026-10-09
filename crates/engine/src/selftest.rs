//! The BidBase Editor's check of an entry: does the rule bid its own
//! test hand? (VB `LookUpBid` with `TestID` and `TestEntry` set.)

use crate::auction::AuctionFacts;
use crate::lookup::{Bidder, LookupError, Request, Vul};
use crate::rules::Rule;

/// What the Editor's check reports for one rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Verdict {
    /// The rule accepts its test hand and is the first rule to take it.
    Ok,
    /// The rule rejects its own test hand: which test, and why.
    OwnFails(String),
    /// The rule accepts it, but an earlier rule takes the hand first.
    Shadowed(i64, String),
    /// Switched off (`Pct_Used` 0/D or its convention off), so only the
    /// rule's own check ran, and it passed.
    Off,
    /// The test hand can't be read (a typo such as "K33").
    BadHand(String),
    /// Something else (no rules found).
    Error(String),
}

/// The lookup the Editor runs for a rule: its test hand, its auction, its
/// vulnerability code, and the trump suit its `Source` names.
pub fn request_for(rule: &Rule, test_entry: bool) -> Request {
    Request {
        hand: rule.test_hand.trim().to_string(),
        prior_calls: rule.prior_bids.clone(),
        vul: match rule.vul.trim() {
            "" => Vul::Unknown,
            v => Vul::Spec(v.to_string()),
        },
        trump: rule.source_trump(),
        test_id: rule.id,
        test_entry,
        trace: true,
        ..Default::default()
    }
}

/// Check a rule as the Editor does on saving it: first the rule alone
/// (`TestEntry`), then the whole auction, where the rule should be the
/// first to take the hand.
pub fn check_rule(bidder: &mut Bidder, rule: &Rule) -> Verdict {
    let own = bidder.lookup(&request_for(rule, true));
    match own.result {
        Ok(_) => {}
        Err(LookupError::Hand(e)) => return Verdict::BadHand(e.to_string()),
        Err(LookupError::TestHandFails) | Err(LookupError::NoMatch) => {
            let why = own
                .tried
                .iter()
                .find_map(|t| t.failed.as_ref())
                .map_or_else(|| "?".to_string(), |(test, why)| format!("{test}: {why}"));
            return Verdict::OwnFails(why);
        }
        Err(e) => return Verdict::Error(e.to_string()),
    }
    let pos_num = AuctionFacts::new(&rule.prior_bids).pos_num;
    if !bidder.is_enabled(rule, pos_num) {
        return Verdict::Off;
    }
    let first = bidder.lookup(&request_for(rule, false));
    match first.result {
        Ok(found) if found.rule_id == rule.id => Verdict::Ok,
        Ok(found) => Verdict::Shadowed(found.rule_id, found.call),
        Err(e) => Verdict::Error(e.to_string()),
    }
}
