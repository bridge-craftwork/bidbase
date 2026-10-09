//! The port against the September 2018 rule set in data/2018.

use std::path::Path;

use bidbase_engine::selftest::{check_rule, Verdict};
use bidbase_engine::{Bidder, Request, RuleBase, Vul};

fn base() -> RuleBase {
    RuleBase::load(&Path::new(env!("CARGO_MANIFEST_DIR")).join("../../data/2018")).unwrap()
}

fn call(bidder: &mut Bidder, hand: &str, auction: &str) -> (String, i64) {
    let out = bidder.lookup(&Request {
        hand: hand.into(),
        prior_calls: auction.into(),
        vul: Vul::Neither,
        ..Default::default()
    });
    let found = out
        .result
        .unwrap_or_else(|e| panic!("{hand} after {auction:?}: {e}"));
    (found.call, found.rule_id)
}

#[test]
fn notrump_ranges_come_from_the_opening_rules() {
    let nt = base().nt;
    assert_eq!(
        (nt.one_min, nt.one_max, nt.two_min, nt.two_max),
        (15, 17, 20, 21)
    );
}

#[test]
fn familiar_auctions() {
    let mut b = Bidder::new(base());
    assert_eq!(call(&mut b, "AK2-KQ2-A432-J32", ""), ("1N".into(), 13));
    assert_eq!(call(&mut b, "AK2-KQ2-A432-J32", "P-P"), ("1N".into(), 13));
    assert_eq!(call(&mut b, "KQ32-A32-K432-32", "1N-P").0, "2C");
    assert_eq!(call(&mut b, "Q8432-32-K32-J32", "1N-P").0, "2H");
    assert_eq!(call(&mut b, "32-2-KQJ9842-432", "P-P").0, "3D");
}

#[test]
fn convention_switches_are_honoured() {
    // In the 2018 copy every four-card 1H response to 1C is switched off.
    let hand = "K32-KJ32-A32-Q32";
    let mut b = Bidder::new(base());
    let out = b.lookup(&Request {
        hand: hand.into(),
        prior_calls: "1C-P".into(),
        vul: Vul::Neither,
        ..Default::default()
    });
    assert!(out.result.is_err());
    let mut all = base();
    all.switch_all_on();
    let mut b = Bidder::new(all);
    assert_eq!(call(&mut b, hand, "1C-P").0, "1H");
}

/// Regression guard: how many rules bid their own test hands. Raise the
/// floors (or lower the ceilings) when a change improves them.
#[test]
fn rules_bid_their_test_hands() {
    let base = base();
    let rules: Vec<_> = base
        .rules
        .iter()
        .filter(|r| !r.is_section_end() && r.has_test_hand())
        .cloned()
        .collect();
    let mut b = Bidder::new(base);
    let (mut ok, mut off, mut shadowed, mut own, mut other) = (0, 0, 0, 0, 0);
    for r in &rules {
        match check_rule(&mut b, r) {
            Verdict::Ok => ok += 1,
            Verdict::Off => off += 1,
            Verdict::Shadowed(..) => shadowed += 1,
            Verdict::OwnFails(_) => own += 1,
            Verdict::BadHand(_) | Verdict::Error(_) => other += 1,
        }
    }
    assert_eq!(rules.len(), 5258);
    assert!(ok >= 3440, "first match is the rule itself: {ok}");
    assert!(off + ok >= 4894, "accepted: {}", off + ok);
    assert!(shadowed <= 289, "taken by an earlier rule: {shadowed}");
    assert!(own <= 66, "rejects its own test hand: {own}");
    assert!(other <= 9, "unreadable or errors: {other}");
}

#[test]
fn a_two_over_one_card_sets_the_switches() {
    use bidbase_engine::card::CardMap;
    let (map, vocab) = CardMap::standard().unwrap();
    let text = std::fs::read_to_string(
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../card/examples/21GF-DEFAULT.card.json"),
    )
    .unwrap();
    let (card, _) = bridge_card::Card::from_json(&vocab, &text).unwrap();
    let mut base = base();
    let applied = map.apply(&card, &mut base);
    // 2/1 game force on, its non-forcing alternative and Precision off.
    assert!(applied.on.contains(&5));
    assert!(applied.off.contains(&179));
    assert!(applied.off.contains(&47));
    let mut b = Bidder::new(base);
    assert_eq!(call(&mut b, "K32-KJ32-A32-Q32", "1C-P").0, "1H");
    assert_eq!(call(&mut b, "KQ32-A32-K432-32", "1N-P").0, "2C");
}
