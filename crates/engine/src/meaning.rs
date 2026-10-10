//! What a rule's call shows, in a line: the rule's conditions (what every
//! hand that makes the call through this rule holds) and what BidBase
//! discloses about it (strength, alert, what it asks partner, the public
//! part of the explanation).
//!
//! "Jacoby Transfer: 5+♠; 0+ HCP (weak or better); artificial, alert;
//! asks 2S; Transfer"

use crate::checks::{hcp_bounds, int_range};
use crate::rules::{NtRanges, Rule, RuleBase};
use crate::vb::{split_alt, suit_index};

const SYMBOL: [&str; 5] = ["", "♣", "♦", "♥", "♠"];

/// The meaning of `rule`'s call after `prior` bids (the rules' key, no
/// opening passes). `x_suit` (1-4) is the suit a "2x" rule bid, when
/// known; 0 lists the x specs as such.
pub fn describe(base: &RuleBase, rule: &Rule, prior: &str, x_suit: usize) -> String {
    let mut parts: Vec<String> = Vec::new();

    // Suits, spades first: lengths with their strength. In a "2x" bid the
    // x suit is the bid's; otherwise ("2C" Stayman, xNum 4 with x in the
    // majors) it is one of the suits whose spec allows x.
    let x_suit = if rule.bid.contains('x') { x_suit } else { 0 };
    let x_num = rule.x_num.trim().trim_end_matches('U').to_string();
    let x_hcp = rule.x_hcp.trim().to_string();
    let mut suits: Vec<String> = Vec::new();
    if x_suit == 0 && (!x_num.is_empty() || !x_hcp.is_empty()) {
        let candidates: Vec<&str> = (1..=4)
            .rev()
            .filter(|&s| {
                rule.suit_num(s)
                    .replace(' ', "")
                    .split(',')
                    .any(|a| a == "x")
                    || rule
                        .suit_hcp(s)
                        .replace(' ', "")
                        .split(',')
                        .any(|a| a == "x")
            })
            .map(|s| SYMBOL[s])
            .collect();
        let which = match candidates.len() {
            0 | 4 => "a suit".to_string(),
            _ => candidates.join(" or "),
        };
        suits.push(suit_text(&format!(" in {which}"), &x_num, &x_hcp));
    }
    for s in (1..=4).rev() {
        let (num, hcp) = if s == x_suit {
            (x_num.clone(), x_hcp.clone())
        } else {
            (not_x(rule.suit_num(s)), not_x(rule.suit_hcp(s)))
        };
        if !num.is_empty() || !hcp.is_empty() {
            suits.push(suit_text(SYMBOL[s], &num, &hcp));
        }
    }
    if !suits.is_empty() {
        parts.push(suits.join(", "));
    }
    for (spec, unit) in [(&rule.cm_num, "cards"), (&rule.cm_hcp, "HCP")] {
        if let Some(t) = combo_text(spec.trim(), unit) {
            parts.push(t);
        }
    }

    // Strength.
    let mut strength = hcp_text(rule, prior, base.nt);
    let label = strength_label(rule.strength_discl.trim());
    if !label.is_empty() {
        strength = format!("{strength} ({label})");
    }
    parts.push(strength);

    // The other conditions.
    let shape = match rule.shape.trim().to_ascii_lowercase().as_str() {
        "b" => "balanced",
        "u" => "unbalanced",
        "1" => "one-suiter",
        "2" => "two-suiter",
        "3" => "three-suiter",
        "s" => "a singleton",
        "v" => "a void",
        "e" => "a singleton or void",
        _ => "",
    };
    if !shape.is_empty() {
        parts.push(shape.to_string());
    }
    let dist = rule.dist_pts.trim();
    if !dist.is_empty() {
        parts.push(match dist.strip_prefix('T') {
            Some(rest) => format!("{} total points", count_text(rest)),
            None => format!("{} distribution points", count_text(dist)),
        });
    }
    // Read as minimums, as the checks read them (checks::min_spec).
    let minimums = [
        (&rule.winners, "winners"),
        (&rule.losers, "losers"),
        (&rule.quick_tricks, "quick tricks"),
        (&rule.stopped_suits, "suits stopped"),
        (&rule.intermed, "intermediates"),
        (&rule.cust_pts, "custom points"),
    ];
    for (spec, what) in minimums {
        let spec = spec.trim();
        if !spec.is_empty() {
            parts.push(format!("{} {what}", min_text(spec)));
        }
    }
    if !rule.vul.trim().is_empty() {
        parts.push(format!("vul {}", rule.vul.trim()));
    }
    if !rule.pos.trim().is_empty() {
        parts.push(format!("position {}", position_text(rule.pos.trim())));
    }

    // What the table is told.
    let mut flags = Vec::new();
    if !rule.artificial.trim().is_empty() {
        flags.push("artificial");
    }
    if rule.alert.trim() == "A" {
        flags.push("alert");
    }
    if !flags.is_empty() {
        parts.push(flags.join(", "));
    }
    if let Some(asks) = asks_text(&rule.asks) {
        parts.push(asks);
    }
    if rule.alert.trim() != "D" {
        let note = crate::lookup::public_explanation(&rule.explanation);
        if !note.is_empty() && !note.starts_with("S:") {
            parts.push(note);
        }
    }

    let name = base
        .conventions
        .get(&rule.convention())
        .map(|c| c.name.trim())
        .filter(|n| !n.is_empty());
    match name {
        Some(n) => format!("{n}: {}", parts.join("; ")),
        None => parts.join("; "),
    }
}

/// A suit spec without its "x" alternative: "x,<4" is "<4" for a suit that
/// is not the x suit, and a bare "x" says nothing.
fn not_x(spec: &str) -> String {
    let spec = spec.replace(' ', "");
    match spec.as_str() {
        "x" => String::new(),
        s => match split_alt(s) {
            (first, rest) if first == "x" => rest,
            _ => s.to_string(),
        },
    }
}

/// "5+♠ with AK/AQ/KQ", "3-4♥", "♦ stopper".
fn suit_text(sym: &str, num: &str, hcp: &str) -> String {
    let len = alternatives(num, |n| length_text(n, sym));
    let strength = alternatives(hcp, |h| strength_text(h, sym));
    match (len.is_empty(), strength.is_empty()) {
        (false, false) => format!("{len} with {strength}"),
        (false, true) => len,
        (true, false) => format!("{sym}: {strength}"),
        (true, true) => String::new(),
    }
}

fn alternatives(spec: &str, f: impl Fn(&str) -> String) -> String {
    spec.split(',')
        .map(str::trim)
        .filter(|s| !s.is_empty())
        .map(f)
        .collect::<Vec<_>>()
        .join(" or ")
}

fn length_text(n: &str, sym: &str) -> String {
    let other = |t: &str| match t.chars().next().map(suit_index) {
        Some(o) if o > 0 && t.len() == 1 => SYMBOL[o].to_string(),
        _ => t.to_string(),
    };
    if n == "0" {
        format!("void in {sym}")
    } else if let Some(t) = n.strip_prefix("<=") {
        format!("{sym} no longer than {}", other(t))
    } else if let Some(t) = n.strip_prefix(">=") {
        format!("{sym} at least as long as {}", other(t))
    } else if let Some(t) = n.strip_prefix('<') {
        if t.parse::<i32>().is_ok() {
            format!("0-{}{sym}", t.parse::<i32>().unwrap() - 1)
        } else {
            format!("{sym} shorter than {}", other(t))
        }
    } else if let Some(t) = n.strip_prefix('>') {
        if t.parse::<i32>().is_ok() {
            format!("{}+{sym}", t.parse::<i32>().unwrap() + 1)
        } else {
            format!("{sym} longer than {}", other(t))
        }
    } else if let Some(t) = n.strip_prefix('=') {
        format!("{sym} as long as {}", other(t))
    } else if let Some((lo, hi)) = n
        .split_once('-')
        .filter(|(_, hi)| hi.parse::<i32>().is_err() && !hi.is_empty())
    {
        format!("{lo}{sym} up to {} length", other(hi))
    } else if let Some(t) = n.strip_suffix('-') {
        format!(
            "0-{}{sym}",
            t.parse::<i32>()
                .map_or(t.to_string(), |v| (v - 1).to_string())
        )
    } else {
        format!("{n}{sym}")
    }
}

fn strength_text(h: &str, sym: &str) -> String {
    let lower = h.to_ascii_lowercase();
    let (main, intermed) = match lower.find('i') {
        Some(i) => (&lower[..i], crate::vb::ival(&lower[i + 1..])),
        None => (lower.as_str(), 0),
    };
    let mut text = match main {
        "" => String::new(),
        "1stop" => "a stopper".into(),
        "2stop" => "a double stopper".into(),
        "hstop" => "a half stopper".into(),
        "0stop" => "no stopper".into(),
        "frc" => "first-round control".into(),
        "src" => "second-round control".into(),
        "p" => "protected honours".into(),
        "a" | "ac" => "two of the top three".into(),
        "b" => "AJT, KJT or QJT".into(),
        "c" => "three of the top four".into(),
        "d" => "four of the top five".into(),
        "ab" | "abc" => "two top honours (or AJT/KJT/QJT)".into(),
        "bc" => "AJT-type or better".into(),
        _ if h.contains('A') && main.len() <= 2 => format!("the {sym}A"),
        _ if main == "k" => format!("the {sym}K"),
        _ if main == "q" => format!("the {sym}Q"),
        _ => format!("{} HCP", count_text(main)),
    };
    if intermed > 0 {
        let i = format!(
            "{intermed}+ intermediate{}",
            if intermed > 1 { "s" } else { "" }
        );
        text = if text.is_empty() {
            i
        } else {
            format!("{text} and {i}")
        };
    }
    text
}

fn hcp_text(rule: &Rule, prior: &str, nt: NtRanges) -> String {
    let spec = rule.hcp.trim().to_ascii_lowercase();
    if spec.is_empty() {
        return "0+ HCP".into();
    }
    if let Some(n) = spec.strip_prefix("kc") {
        return format!(
            "{} key cards",
            n.replace("+q", " with the trump Q")
                .replace("-q", " without the trump Q")
        );
    }
    if let Some(n) = spec.strip_prefix('a').filter(|n| n.parse::<i32>().is_ok()) {
        return format!("{n} aces");
    }
    if let Some(n) = spec.strip_prefix('k').filter(|n| n.parse::<i32>().is_ok()) {
        return format!("{n} kings");
    }
    let (lo, hi) = hcp_bounds(&spec, prior, &rule.bid, nt);
    range_text(lo, hi, "HCP")
}

fn range_text(lo: i32, hi: i32, unit: &str) -> String {
    if hi >= 40 {
        format!("{lo}+ {unit}")
    } else if lo == hi {
        format!("{lo} {unit}")
    } else {
        format!("{lo}-{hi} {unit}")
    }
}

/// "2.5" -> "2.5+", "<3" -> "0-2", "3-4" -> "3-4", "6+" -> "6+".
fn count_text(spec: &str) -> String {
    if spec.contains('.') {
        return format!("{}+", spec.trim_end_matches('+'));
    }
    let (lo, hi) = int_range(spec, 99);
    if hi >= 99 {
        format!("{lo}+")
    } else if lo == hi {
        lo.to_string()
    } else {
        format!("{lo}-{hi}")
    }
}

/// A minimum count: "7.5" and "7-8" are 7.5+ and 7+; "<6" is 0-5.
fn min_text(spec: &str) -> String {
    match spec.strip_prefix('<') {
        Some(t) => format!("fewer than {t}"),
        None => {
            let n = spec.split(['-', '+']).next().unwrap_or(spec);
            format!("{n}+")
        }
    }
}

/// "SH>8" -> "♠+♥ 9+ cards"; key-card specs as written.
fn combo_text(spec: &str, unit: &str) -> Option<String> {
    if spec.is_empty() {
        return None;
    }
    let letters: String = spec.chars().take_while(|c| "CDHSx".contains(*c)).collect();
    let rest = &spec[letters.len()..];
    if (2..=3).contains(&letters.len()) && rest.starts_with(['<', '>', '=']) {
        let suits: Vec<&str> = letters
            .chars()
            .map(|c| if c == 'x' { "x" } else { SYMBOL[suit_index(c)] })
            .collect();
        let n = crate::vb::ival(&rest[1..]);
        let range = match &rest[..1] {
            ">" => format!("{}+", n + 1),
            "<" => format!("0-{}", n - 1),
            _ => rest[1..].to_string(),
        };
        return Some(format!("{} {range} {unit}", suits.join("+")));
    }
    Some(format!("key cards {spec}"))
}

fn strength_label(code: &str) -> &'static str {
    match code {
        "V" => "very weak",
        "W" => "weak",
        "I" => "intermediate",
        "O" => "opening strength",
        "S" => "strong",
        "WS" => "weak or strong",
        "WI" => "weak to intermediate",
        "IS" | "I+" => "intermediate or better",
        "W+" => "weak or better",
        "S+" => "strong or better",
        "VS" => "very strong",
        _ => "",
    }
}

fn position_text(pos: &str) -> String {
    match pos {
        "Y" => "passed hand".into(),
        "N" => "unpassed hand".into(),
        "F" => "1st/2nd seat".into(),
        "T" => "3rd/4th seat".into(),
        "P" => "partner passed".into(),
        "U" => "partner unpassed".into(),
        p => format!("seat {p}"),
    }
}

fn asks_text(asks: &str) -> Option<String> {
    let a = asks.trim();
    if a.is_empty() {
        return None;
    }
    Some(match a.to_ascii_lowercase().as_str() {
        "g" => "game forcing".into(),
        "1" => "forcing one round".into(),
        "poc" => "pass or correct".into(),
        "p" => "asks partner to pick".into(),
        "a" => "asks for aces".into(),
        "k" => "asks for kings".into(),
        "kc" => "asks for key cards".into(),
        "s-o" => "sign-off".into(),
        "bm" if a.ends_with('m') => "asks partner to bid a minor".into(),
        "bm" => "asks partner to bid a major".into(),
        l if l.starts_with('b') && l.len() > 1 => format!("asks partner to bid {}", &a[1..]),
        _ => format!("asks {a}"),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lengths_and_strengths() {
        assert_eq!(suit_text("♠", "5+", ""), "5+♠");
        assert_eq!(suit_text("♥", "<4", ""), "0-3♥");
        assert_eq!(suit_text("♦", "", "1stop"), "♦: a stopper");
        assert_eq!(
            suit_text("♣", "6+", "abi1"),
            "6+♣ with two top honours (or AJT/KJT/QJT) and 1+ intermediate"
        );
        assert_eq!(suit_text("♠", "<=H", ""), "♠ no longer than ♥");
        assert_eq!(not_x("x, <4"), "<4");
        assert_eq!(not_x("x"), "");
    }

    #[test]
    fn counts_and_combos() {
        assert_eq!(count_text("2.5"), "2.5+");
        assert_eq!(count_text("<3"), "0-2");
        assert_eq!(range_text(15, 17, "HCP"), "15-17 HCP");
        assert_eq!(combo_text("SH>8", "cards").unwrap(), "♠+♥ 9+ cards");
    }
}
