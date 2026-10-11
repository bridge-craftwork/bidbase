//! Between bridge-types and BidBase's own notation.

use bridge_types::{Call, Direction, Hand, Strain, Suit, Vulnerability};

use crate::lookup::Vul;

/// A hand as BidBase writes it: "AK432-64-QJT3-T2" (spades first).
pub fn hand_to_bidbase(hand: &Hand) -> String {
    hand.to_pbn().replace('.', "-")
}

/// A call as a BidBase token: P, D, R, 1C ... 7N.
pub fn call_to_bidbase(call: &Call) -> String {
    match call {
        Call::Pass => "P".into(),
        Call::Double => "D".into(),
        Call::Redouble => "R".into(),
        Call::Bid { level, strain } => format!("{level}{}", strain_char(*strain)),
        Call::Continue | Call::Blank => "?".into(),
    }
}

/// A BidBase token back to a call.
pub fn call_from_bidbase(token: &str) -> Option<Call> {
    match token {
        "P" => Some(Call::Pass),
        "D" => Some(Call::Double),
        "R" => Some(Call::Redouble),
        _ => Call::from_pbn(token),
    }
}

/// A list of calls as BidBase prior calls: "P-1C-1N".
pub fn calls_to_bidbase(calls: &[Call]) -> String {
    calls
        .iter()
        .map(call_to_bidbase)
        .collect::<Vec<_>>()
        .join("-")
}

/// A hand typed by a person: BidBase "AK432-64-QJT3-T2" or PBN
/// "AK432.64.QJT3.T2" (spades first; "10" for T allowed).
pub fn normalize_hand(hand: &str) -> Option<String> {
    let hand = hand.trim().to_ascii_uppercase().replace("10", "T");
    if hand.contains('.') {
        let parsed = Hand::from_pbn(&hand)?;
        return Some(hand_to_bidbase(&parsed));
    }
    Some(hand)
}

/// An auction typed by a person: "P-1C-1N", "Pass 1C 1NT", "p 1c x" as
/// BidBase tokens joined by dashes.
pub fn normalize_auction(auction: &str) -> Option<String> {
    let mut out = Vec::new();
    for t in auction
        .split(|c: char| c == '-' || c.is_whitespace() || c == ',')
        .filter(|t| !t.is_empty())
    {
        let up = t.to_ascii_uppercase();
        out.push(match up.as_str() {
            "P" | "PASS" => "P".to_string(),
            "D" | "X" | "DBL" => "D".to_string(),
            "R" | "XX" | "RDBL" => "R".to_string(),
            _ => {
                let call =
                    Call::from_pbn(&up).or_else(|| Call::from_pbn(&up.replace('N', "NT")))?;
                call_to_bidbase(&call)
            }
        });
    }
    Some(out.join("-"))
}

/// The board's vulnerability from the bidder's side.
pub fn vul_for(vul: Vulnerability, bidder: Direction) -> Vul {
    let we = vul.is_vulnerable(bidder);
    let they = vul.is_vulnerable(bidder.next());
    match (we, they) {
        (true, true) => Vul::Both,
        (true, false) => Vul::We,
        (false, true) => Vul::They,
        (false, false) => Vul::Neither,
    }
}

/// Suit 1-4 (C..S) as BidBase numbers them.
pub fn suit_number(suit: Suit) -> usize {
    suit as usize + 1
}

fn strain_char(strain: Strain) -> char {
    match strain {
        Strain::Clubs => 'C',
        Strain::Diamonds => 'D',
        Strain::Hearts => 'H',
        Strain::Spades => 'S',
        Strain::NoTrump => 'N',
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trips() {
        let hand = Hand::from_pbn("AK432.64.QJT3.T2").unwrap();
        assert_eq!(hand_to_bidbase(&hand), "AK432-64-QJT3-T2");
        let calls: Vec<Call> = ["Pass", "1NT", "X", "XX", "2S"]
            .iter()
            .map(|c| Call::from_pbn(c).unwrap())
            .collect();
        assert_eq!(calls_to_bidbase(&calls), "P-1N-D-R-2S");
        assert_eq!(call_from_bidbase("1N"), Call::from_pbn("1NT"));
        assert_eq!(call_from_bidbase("D"), Some(Call::Double));
        assert_eq!(
            normalize_auction("Pass 1c 1NT x").as_deref(),
            Some("P-1C-1N-D")
        );
        assert_eq!(
            normalize_hand("AK432.64.QJT3.T2").as_deref(),
            Some("AK432-64-QJT3-T2")
        );
    }

    #[test]
    fn vulnerability_from_the_bidders_side() {
        assert_eq!(
            vul_for(Vulnerability::NorthSouth, Direction::North),
            Vul::We
        );
        assert_eq!(
            vul_for(Vulnerability::NorthSouth, Direction::East),
            Vul::They
        );
        assert_eq!(vul_for(Vulnerability::Both, Direction::East), Vul::Both);
    }
}
