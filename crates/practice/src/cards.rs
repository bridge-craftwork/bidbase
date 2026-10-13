//! Playing cards, drawn: our stand-in for Qcard32.dll's card pictures.
//! Same size as Qcard32's (71 x 96), so the hands lay out as in Nelson's
//! program: an index in the top-left corner, pips as on a real card, a
//! framed letter for the court cards.

use std::f32::consts::PI;

use bridge_types::{Card, Rank, Suit};
use egui::epaint::TextShape;
use egui::{Align2, Color32, FontId, Painter, Pos2, Rect, Stroke, StrokeKind, Vec2};

pub const WIDTH: f32 = 71.0;
pub const HEIGHT: f32 = 96.0;

const RED: Color32 = Color32::from_rgb(0xE0, 0x10, 0x10);

pub fn suit_char(s: Suit) -> char {
    match s {
        Suit::Spades => '♠',
        Suit::Hearts => '♥',
        Suit::Diamonds => '♦',
        Suit::Clubs => '♣',
    }
}

pub fn rank_text(r: Rank) -> &'static str {
    match r {
        Rank::Ace => "A",
        Rank::King => "K",
        Rank::Queen => "Q",
        Rank::Jack => "J",
        Rank::Ten => "10",
        Rank::Nine => "9",
        Rank::Eight => "8",
        Rank::Seven => "7",
        Rank::Six => "6",
        Rank::Five => "5",
        Rank::Four => "4",
        Rank::Three => "3",
        Rank::Two => "2",
    }
}

pub fn suit_color(s: Suit) -> Color32 {
    if s.is_red() {
        RED
    } else {
        Color32::BLACK
    }
}

fn text(
    p: &Painter,
    at: Pos2,
    anchor: Align2,
    s: &str,
    font: FontId,
    color: Color32,
    flipped: bool,
) {
    let g = p.layout_no_wrap(s.to_string(), font, color);
    if flipped {
        // Rotate 180 degrees about `at`, so the anchor lands where it would.
        let r = anchor.anchor_size(at, g.size());
        let corner = Pos2::new(r.right(), r.bottom());
        p.add(TextShape::new(corner, g, color).with_angle(PI));
    } else {
        let r = anchor.anchor_size(at, g.size());
        p.galley(r.min, g, color);
    }
}

/// Where a card's pips go, as fractions of the pip area (x, y, upside down).
fn pips(n: u8) -> Vec<(f32, f32, bool)> {
    const L: f32 = 0.0;
    const M: f32 = 0.5;
    const R: f32 = 1.0;
    let col = |x: f32, ys: &[f32]| ys.iter().map(move |&y| (x, y, y > 0.5)).collect::<Vec<_>>();
    let mut v = Vec::new();
    match n {
        2 => v.extend(col(M, &[0.0, 1.0])),
        3 => v.extend(col(M, &[0.0, 0.5, 1.0])),
        4 => {
            v.extend(col(L, &[0.0, 1.0]));
            v.extend(col(R, &[0.0, 1.0]));
        }
        5 => {
            v.extend(col(L, &[0.0, 1.0]));
            v.extend(col(R, &[0.0, 1.0]));
            v.extend(col(M, &[0.5]));
        }
        6 => {
            v.extend(col(L, &[0.0, 0.5, 1.0]));
            v.extend(col(R, &[0.0, 0.5, 1.0]));
        }
        7 => {
            v.extend(col(L, &[0.0, 0.5, 1.0]));
            v.extend(col(R, &[0.0, 0.5, 1.0]));
            v.extend(col(M, &[0.25]));
        }
        8 => {
            v.extend(col(L, &[0.0, 0.5, 1.0]));
            v.extend(col(R, &[0.0, 0.5, 1.0]));
            v.extend(col(M, &[0.25, 0.75]));
        }
        9 => {
            v.extend(col(L, &[0.0, 1.0 / 3.0, 2.0 / 3.0, 1.0]));
            v.extend(col(R, &[0.0, 1.0 / 3.0, 2.0 / 3.0, 1.0]));
            v.extend(col(M, &[0.5]));
        }
        10 => {
            v.extend(col(L, &[0.0, 1.0 / 3.0, 2.0 / 3.0, 1.0]));
            v.extend(col(R, &[0.0, 1.0 / 3.0, 2.0 / 3.0, 1.0]));
            v.extend(col(M, &[1.0 / 6.0, 5.0 / 6.0]));
        }
        _ => {}
    }
    v
}

/// Draw a card face with its top-left corner at `at`.
pub fn face(p: &Painter, at: Pos2, card: Card) {
    let r = Rect::from_min_size(at, Vec2::new(WIDTH, HEIGHT));
    p.rect_filled(r, 4.0, Color32::WHITE);
    p.rect_stroke(r, 4.0, Stroke::new(1.0, Color32::BLACK), StrokeKind::Inside);
    let color = suit_color(card.suit);
    let suit = suit_char(card.suit).to_string();
    let rank = rank_text(card.rank);
    let index_font = FontId::new(12.0, vbui_bold());
    let small = FontId::new(10.0, egui::FontFamily::Proportional);
    // Corner indexes: top-left, and bottom-right upside down.
    text(
        p,
        at + Vec2::new(8.0, 2.0),
        Align2::CENTER_TOP,
        rank,
        index_font.clone(),
        color,
        false,
    );
    text(
        p,
        at + Vec2::new(8.0, 15.0),
        Align2::CENTER_TOP,
        &suit,
        small.clone(),
        color,
        false,
    );
    let br = r.max;
    text(
        p,
        br - Vec2::new(8.0, 2.0),
        Align2::CENTER_BOTTOM,
        rank,
        index_font,
        color,
        true,
    );
    text(
        p,
        br - Vec2::new(8.0, 15.0),
        Align2::CENTER_BOTTOM,
        &suit,
        small,
        color,
        true,
    );

    let v = card.rank as u8;
    if v <= 10 {
        // Pip area inside the indexes.
        let area = Rect::from_min_max(at + Vec2::new(22.0, 14.0), br - Vec2::new(22.0, 14.0));
        let pip = FontId::new(17.0, egui::FontFamily::Proportional);
        for (fx, fy, flip) in pips(v) {
            let c = Pos2::new(
                area.left() + fx * area.width(),
                area.top() + fy * area.height(),
            );
            text(p, c, Align2::CENTER_CENTER, &suit, pip.clone(), color, flip);
        }
    } else if card.rank == Rank::Ace {
        let size = if card.suit == Suit::Spades {
            44.0
        } else {
            36.0
        };
        text(
            p,
            r.center(),
            Align2::CENTER_CENTER,
            &suit,
            FontId::new(size, egui::FontFamily::Proportional),
            color,
            false,
        );
    } else {
        // Court card: a framed panel with the letter and suit.
        let panel = Rect::from_min_max(at + Vec2::new(15.0, 12.0), br - Vec2::new(15.0, 12.0));
        p.rect_filled(panel, 2.0, Color32::from_rgb(0xFF, 0xF4, 0xC8));
        p.rect_stroke(panel, 2.0, Stroke::new(1.5, color), StrokeKind::Inside);
        let letter = FontId::new(26.0, vbui_bold());
        text(
            p,
            panel.center() - Vec2::new(0.0, 9.0),
            Align2::CENTER_CENTER,
            rank,
            letter,
            color,
            false,
        );
        let pip = FontId::new(18.0, egui::FontFamily::Proportional);
        text(
            p,
            panel.center() + Vec2::new(0.0, 15.0),
            Align2::CENTER_CENTER,
            &suit,
            pip,
            color,
            false,
        );
    }
}

/// Draw a card's back (a hidden hand): red, as Qcard32's back design 3.
pub fn back(p: &Painter, at: Pos2) {
    let r = Rect::from_min_size(at, Vec2::new(WIDTH, HEIGHT));
    p.rect_filled(r, 4.0, Color32::WHITE);
    p.rect_stroke(r, 4.0, Stroke::new(1.0, Color32::BLACK), StrokeKind::Inside);
    let inner = r.shrink(4.0);
    p.rect_filled(inner, 2.0, Color32::from_rgb(0xB0, 0x18, 0x18));
    let lattice = Stroke::new(1.0, Color32::from_rgb(0xE8, 0x7A, 0x6A));
    let p2 = p.with_clip_rect(inner);
    let mut x = inner.left() - inner.height();
    while x < inner.right() {
        p2.line_segment(
            [
                Pos2::new(x, inner.top()),
                Pos2::new(x + inner.height(), inner.bottom()),
            ],
            lattice,
        );
        p2.line_segment(
            [
                Pos2::new(x + inner.height(), inner.top()),
                Pos2::new(x, inner.bottom()),
            ],
            lattice,
        );
        x += 8.0;
    }
}

fn vbui_bold() -> egui::FontFamily {
    bidbase_vbui::theme::family("", true)
}
