//! The DDA table under the hands (`ShowDDA`, frm:7161): double-dummy
//! tricks per declarer and strain, with each hand's HCP, the pairs' Law of
//! Total Tricks fit and the losing-trick count. Five lines, as Nelson's
//! program writes them into `cmb_DDA`.
//!
//! One departure: his East-West "self-sufficient suit" test adds the
//! suit's HCP where North-South's adds its honours (frm:7353, a slip);
//! both use honours here.

use bridge_types::{Card, DdTable, Deal, Direction, Hand, Rank, Strain, Suit};

/// The text line above the list.
pub const CAPTION: &str = "DDA = Maximum tricks that can be taken by each player in each suit/NT.";

const STRAINS: [Strain; 5] = [
    Strain::Clubs,
    Strain::Diamonds,
    Strain::Hearts,
    Strain::Spades,
    Strain::NoTrump,
];
const SUITS: [Suit; 4] = [Suit::Clubs, Suit::Diamonds, Suit::Hearts, Suit::Spades];

/// A declarer's five trick counts, C D H S N: ` 7  8 10  9  6`.
pub fn tricks_line(t: &DdTable, d: Direction) -> String {
    let v: Vec<u8> = STRAINS.iter().map(|s| t.tricks(d, *s)).collect();
    format!("{:>2}{:>3}{:>3}{:>3}{:>3}", v[0], v[1], v[2], v[3], v[4])
}

fn has(h: &Hand, s: Suit, r: Rank) -> bool {
    h.cards().contains(&Card { suit: s, rank: r })
}

/// Losing-trick count, with Nelson's half losers (frm:7361-7441).
pub fn ltc(h: &Hand) -> f32 {
    let mut n = 0.0;
    for s in SUITS {
        let len = h.suit_length(s);
        let (a, k, q, j, t) = (
            has(h, s, Rank::Ace),
            has(h, s, Rank::King),
            has(h, s, Rank::Queen),
            has(h, s, Rank::Jack),
            has(h, s, Rank::Ten),
        );
        n += match len {
            0 => 0.0,
            1 => {
                if a {
                    0.0
                } else {
                    1.0
                }
            }
            2 => match (a, k, q) {
                (true, true, _) => 0.0,
                (true, false, true) => 0.5,
                (true, false, false) | (false, true, _) => 1.0,
                _ => 2.0,
            },
            _ => {
                if a && k {
                    if q {
                        0.0
                    } else if j {
                        0.5
                    } else {
                        1.0
                    }
                } else if a && q {
                    if j {
                        1.0
                    } else {
                        1.5
                    }
                } else if a {
                    2.0
                } else if k {
                    if q {
                        1.0
                    } else {
                        2.0
                    }
                } else if q {
                    if j || t {
                        2.0
                    } else {
                        2.5
                    }
                } else {
                    3.0
                }
            }
        };
    }
    n
}

/// A number as VB prints a Single: "7", "7.5", "-0.5".
pub fn vb_single(x: f32) -> String {
    if x.fract() == 0.0 {
        format!("{}", x as i32)
    } else {
        format!("{x}")
    }
}

fn honours(h: &Hand, s: Suit) -> usize {
    h.cards_in_suit(s)
        .iter()
        .filter(|c| c.rank >= Rank::Ten)
        .count()
}

fn hcp(h: &Hand) -> u8 {
    h.hcp()
}

/// The pair's longest combined suit, and whether it has a fit: eight cards
/// together, or one hand's suit with length + honours over 9.
fn pair(a: &Hand, b: &Hand) -> (usize, bool) {
    let mut lott = 0;
    let mut fit = false;
    for s in SUITS {
        let n = a.suit_length(s) + b.suit_length(s);
        lott = lott.max(n);
        if n > 7 || a.suit_length(s) + honours(a, s) > 9 || b.suit_length(s) + honours(b, s) > 9 {
            fit = true;
        }
    }
    (lott, fit)
}

fn ltc_cell(x: f32) -> String {
    let s = vb_single(x);
    match s.len() {
        1 => format!(" {s} "),
        2 => format!("{s} "),
        _ => s,
    }
}

/// The five lines of the table.
pub fn lines(deal: &Deal, t: &DdTable) -> Vec<String> {
    let h = |d| deal.hand(d);
    let (n, s, e, w) = (
        h(Direction::North),
        h(Direction::South),
        h(Direction::East),
        h(Direction::West),
    );
    let (lott_ns, fit_ns) = pair(s, n);
    let (lott_ew, fit_ew) = pair(w, e);
    let caption = |a: f32, b: f32, fit: bool| {
        if fit {
            format!(
                "-> 24 - {} - {} = {}",
                vb_single(a),
                vb_single(b),
                vb_single(24.0 - a - b)
            )
        } else {
            "-> No 8-card fit.".to_string()
        }
    };
    let (ln, ls, le, lw) = (ltc(n), ltc(s), ltc(e), ltc(w));
    vec![
        "DDA:C  D  H  S  N | HCP | LoTT | LTC: Calculate Tricks".to_string(),
        format!(
            "N: {} | {:>3} |      | {}   Pairs Can Take:",
            tricks_line(t, Direction::North),
            hcp(n),
            ltc_cell(ln)
        ),
        format!(
            "S: {} |  {:>2} | {:>3}  | {}{}",
            tricks_line(t, Direction::South),
            hcp(s),
            lott_ns,
            ltc_cell(ls),
            caption(ln, ls, fit_ns)
        ),
        format!(
            "E: {} |  {:>2} |      | {}",
            tricks_line(t, Direction::East),
            hcp(e),
            ltc_cell(le)
        ),
        format!(
            "W: {} |  {:>2} | {:>3}  | {}{}",
            tricks_line(t, Direction::West),
            hcp(w),
            lott_ew,
            ltc_cell(lw),
            caption(le, lw, fit_ew)
        ),
    ]
}

#[cfg(test)]
mod tests {
    use super::*;

    fn hand(s: &str) -> Hand {
        // "AKQ-JT9-876-5432": spades-hearts-diamonds-clubs.
        let mut h = Hand::new();
        for (part, suit) in
            s.split('-')
                .zip([Suit::Spades, Suit::Hearts, Suit::Diamonds, Suit::Clubs])
        {
            for c in part.chars() {
                h.add_card(Card {
                    suit,
                    rank: Rank::from_char(c).unwrap(),
                });
            }
        }
        h
    }

    #[test]
    fn losing_tricks() {
        assert_eq!(ltc(&hand("AKQ-AKJ-AQJ-A2")), 0.0 + 0.5 + 1.0 + 1.0);
        assert_eq!(ltc(&hand("AQ-K2-Q32-J98765")), 0.5 + 1.0 + 2.5 + 3.0);
        assert_eq!(ltc(&hand("-A-QT2-J98765432")), 0.0 + 2.0 + 3.0);
    }

    #[test]
    fn line_layout() {
        let t = DdTable::from_fn(|d, s| match (d, s) {
            (Direction::North, Strain::Clubs) => 10,
            (_, Strain::NoTrump) => 6,
            _ => 7,
        });
        assert_eq!(tricks_line(&t, Direction::North), "10  7  7  7  6");
        assert_eq!(tricks_line(&t, Direction::South), " 7  7  7  7  6");
        let deal = crate::dealer::deal(31118);
        let l = lines(&deal, &t);
        assert_eq!(l.len(), 5);
        // North J965-J4-JT73-J87: 4 HCP, 3 + 2 + 3 + 3 = 11 losers.
        assert_eq!(
            l[1],
            "N: 10  7  7  7  6 |   4 |      | 11    Pairs Can Take:"
        );
        assert_eq!(
            l[2],
            "S:  7  7  7  7  6 |   7 |   9  |  9 -> 24 - 11 - 9 = 4"
        );
        assert_eq!(vb_single(7.5), "7.5");
        assert_eq!(ltc_cell(7.0), " 7 ");
    }
}
