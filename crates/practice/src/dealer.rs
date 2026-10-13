//! Nelson's dealer: deal n of 1..9,999,999 is always the same four hands.
//! VB6's `Rnd` reseeded with `Rnd(-n)` drives his "old routine" shuffle
//! (`BidBase Bidding Practice Dealer.bas`:479-546), so deals cited in
//! the rules ("Source: Deal #24-N") come out as he saw them.
//! docs/PRACTICE-PROGRAM.md, section 4.

use bridge_types::{Card, Deal, Direction, Rank, Suit};

/// The highest deal number the program accepts.
pub const MAX_DEAL: u32 = 9_999_999;

/// VB6's `Rnd`: a 24-bit linear congruential generator.
#[derive(Debug, Clone)]
pub struct VbRnd {
    seed: u32,
}

impl Default for VbRnd {
    fn default() -> Self {
        VbRnd { seed: 0x50000 }
    }
}

impl VbRnd {
    /// `Rnd(x)` for x < 0: reseed from the bits of x as a Single, then step.
    pub fn reseed(&mut self, x: f32) -> f32 {
        let b = x.to_bits();
        self.seed = b.wrapping_add(b >> 24) & 0xFF_FFFF;
        self.next()
    }

    /// `Rnd()`.
    #[allow(clippy::should_implement_trait)]
    pub fn next(&mut self) -> f32 {
        self.seed = self.seed.wrapping_mul(0x43FD_43FD).wrapping_add(0xC3_9EC3) & 0xFF_FFFF;
        self.seed as f32 / 16_777_216.0
    }
}

/// VB card number 1..52 (1-13 clubs, 14-26 diamonds, 27-39 hearts,
/// 40-52 spades; in a suit 1 = ace, 2..13 = two..king) as a card.
pub fn card_of(n: u8) -> Card {
    let suit = match (n - 1) / 13 {
        0 => Suit::Clubs,
        1 => Suit::Diamonds,
        2 => Suit::Hearts,
        _ => Suit::Spades,
    };
    let r = (n - 1) % 13 + 1;
    let rank = if r == 1 {
        Rank::Ace
    } else {
        Rank::from_value(r).unwrap()
    };
    Card { suit, rank }
}

/// The inverse of [`card_of`].
pub fn number_of(c: Card) -> u8 {
    let s = match c.suit {
        Suit::Clubs => 0,
        Suit::Diamonds => 1,
        Suit::Hearts => 2,
        Suit::Spades => 3,
    };
    let r = if c.rank == Rank::Ace { 1 } else { c.rank as u8 };
    s * 13 + r
}

/// The program's seats in its own order: 1 = South, 2 = West, 3 = North,
/// 4 = East (`p_Hand(1..4)`).
pub const SEATS: [Direction; 4] = [
    Direction::South,
    Direction::West,
    Direction::North,
    Direction::East,
];

/// Deal number `n`.
pub fn deal(n: u32) -> Deal {
    let mut r = VbRnd::default();
    r.reseed(-(n as f32));
    let mut c: Vec<u8> = (0..=52).collect();
    let passes = 10 + (5.0 * r.next() as f64).floor() as usize;
    for _ in 0..passes {
        for i in 1..=52 {
            // Double precision, as VB evaluates 52 * Rnd + 1.
            let k = (52.0 * r.next() as f64 + 1.0).floor() as usize;
            c.swap(i, k);
        }
    }
    let mut d = Deal::new();
    for (s, seat) in SEATS.iter().enumerate() {
        for &n in &c[1 + 13 * s..=13 * (s + 1)] {
            d.hand_mut(*seat).add_card(card_of(n));
        }
    }
    d
}

/// Shuffle `cards` the way the program's Mix Up Cards does
/// (`ScrambleHands`): `10 + Int(5 * Rnd)` passes of n swaps, without
/// reseeding.
pub fn scramble(r: &mut VbRnd, cards: &mut [Card]) {
    let n = cards.len();
    if n < 2 {
        return;
    }
    let passes = 10 + (5.0 * r.next() as f64).floor() as usize;
    for _ in 0..passes {
        for i in 0..n {
            let k = (n as f64 * r.next() as f64).floor() as usize;
            cards.swap(i, k.min(n - 1));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use bidbase_engine::convert::hand_to_bidbase;

    #[test]
    fn rnd_matches_vb() {
        let mut r = VbRnd::default();
        let v = r.reseed(-1.0);
        assert!((v - 0.224_007_6).abs() < 1e-6, "{v}");
        assert_eq!(r.seed, 3_758_214);
    }

    #[test]
    fn card_numbers_round_trip() {
        for n in 1..=52 {
            assert_eq!(number_of(card_of(n)), n);
        }
        assert_eq!(
            card_of(1),
            Card {
                suit: Suit::Clubs,
                rank: Rank::Ace
            }
        );
        assert_eq!(
            card_of(52),
            Card {
                suit: Suit::Spades,
                rank: Rank::King
            }
        );
    }

    /// docs/PRACTICE-PROGRAM.md, section 4 (S / W / N / E).
    #[test]
    fn deals_as_nelson_dealt_them() {
        let cases: [(u32, [&str; 4]); 5] = [
            (
                1,
                [
                    "J2-T2-AQT9-AQT52",
                    "AQT8-AJ5-J72-876",
                    "9765-KQ874-5-943",
                    "K43-963-K8643-KJ",
                ],
            ),
            (
                24,
                [
                    "AQ9-KT87-T8542-5",
                    "J8742-Q-K96-KJT2",
                    "T53-AJ52-A-AQ643",
                    "K6-9643-QJ73-987",
                ],
            ),
            (
                20307,
                [
                    "QT9642-T9-A9-KJ5",
                    "J8-J86-QJT754-T9",
                    "K5-AK2-K8-AQ7432",
                    "A73-Q7543-632-86",
                ],
            ),
            (
                27214,
                [
                    "KT8-AQJ6-AK932-2",
                    "J9642-5-J8764-86",
                    "Q753-8-Q5-AKJ753",
                    "A-KT97432-T-QT94",
                ],
            ),
            (
                31118,
                [
                    "T32-QT-Q9-KT9643",
                    "A874-AK3-A862-A2",
                    "J965-J4-JT73-J87",
                    "KQ-987652-K54-Q5",
                ],
            ),
        ];
        for (n, hands) in cases {
            let d = deal(n);
            for (seat, want) in SEATS.iter().zip(hands) {
                assert_eq!(hand_to_bidbase(d.hand(*seat)), want, "deal {n} {seat:?}");
            }
        }
    }
}
