//! Nelson's 2026 rules use vulnerability and position codes as his
//! abbreviations page defines them (aeyec.com/bidbase/notes/
//! cardShark_BidBase_abbreviations.htm); his VB matched some of them
//! differently, so those rules never fired (docs/ENGINE-2026.md).

use std::path::Path;

use bidbase_engine::{Bidder, Request, RuleBase, Vul};

fn bidder() -> Bidder {
    Bidder::new(
        RuleBase::load(&Path::new(env!("CARGO_MANIFEST_DIR")).join("../../data/2026")).unwrap(),
    )
}

/// The rule chosen, if any.
fn rule_for(b: &mut Bidder, hand: &str, auction: &str, vul: Vul) -> Option<i64> {
    let out = b.lookup(&Request {
        hand: hand.into(),
        prior_calls: auction.into(),
        vul,
        ..Default::default()
    });
    out.result.ok().map(|f| f.rule_id)
}

#[test]
fn vul_v_means_both_vulnerable() {
    // Rule 269, a weak jump shift marked V: the client sends B for both.
    let mut b = bidder();
    assert_eq!(
        rule_for(&mut b, "Q65432-32-K2-432", "1C-1D", Vul::Both),
        Some(269)
    );
    assert_ne!(
        rule_for(&mut b, "Q65432-32-K2-432", "1C-1D", Vul::We),
        Some(269)
    );
}

#[test]
fn vul_n_means_neither_vulnerable() {
    // Rule 22490, marked N: "neither" used to be sent as "".
    let mut b = bidder();
    assert_eq!(
        rule_for(&mut b, "2-2-KQT8632-J863", "", Vul::Neither),
        Some(22490)
    );
    assert_ne!(
        rule_for(&mut b, "2-2-KQT8632-J863", "", Vul::Both),
        Some(22490)
    );
}

#[test]
fn position_f_is_first_or_second_seat() {
    // Rule 14728, marked F, in second seat (VB: first seat only).
    let mut b = bidder();
    assert_eq!(
        rule_for(&mut b, "2-T6-KQ864-AJT32", "P", Vul::Neither),
        Some(14728)
    );
    assert_ne!(
        rule_for(&mut b, "2-T6-KQ864-AJT32", "P-P", Vul::Neither),
        Some(14728)
    );
}
