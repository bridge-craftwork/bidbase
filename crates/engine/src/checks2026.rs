//! The tests as Nelson's 2026 engine reads them (docs/ENGINE-2026.md).
//!
//! Ported as written, quirks included: the 2026 rules were saved against
//! this code, and their test hands are the check. Where VB evaluates a
//! number as a boolean (`ok = InStr(...)`), the port compares with 0.

// Loops index the per-suit arrays (1-4) as the VB does.
#![allow(clippy::needless_range_loop)]

use crate::checks::{Check, Ctx};
use crate::hand::{ACE, JACK, KING, QUEEN};
use crate::rules::Rule;
use crate::vb::{instr, ival, suit_index, val};

/// `Left$(s, n)` and `Mid$(s, start, len)` on ASCII text.
fn left(s: &str, n: usize) -> &str {
    &s[..n.min(s.len())]
}

fn mid(s: &str, start: usize, len: usize) -> &str {
    let a = (start - 1).min(s.len());
    &s[a..(a + len).min(s.len())]
}

fn mid_from(s: &str, start: usize) -> &str {
    &s[(start - 1).min(s.len())..]
}

fn right(s: &str, n: usize) -> &str {
    &s[s.len().saturating_sub(n)..]
}

/// `InStr("CDHS", c)` for a one-character string.
fn cdhs(c: &str) -> usize {
    if c.len() == 1 {
        suit_index(c.chars().next().unwrap())
    } else {
        0
    }
}

impl<'a> Ctx<'a> {
    // ----- HCP -----------------------------------------------------------

    /// 2026 `TotalHCPs_Ok`. "P" specs work; "1N+a-b"; no intermediates
    /// credit; ">n" reads as 0-40 (`Val(">n")` is 0).
    pub(crate) fn total_hcp_2026(&mut self, r: &Rule) -> Check {
        let raw = r.hcp.as_str();
        if raw.is_empty() {
            return Ok(());
        }
        let mut spec = raw.trim().to_ascii_lowercase();
        if let Some(rest) = spec.strip_prefix('t') {
            spec = rest.to_string();
        }
        if spec.is_empty() {
            return Ok(());
        }
        let prior = self.auction.prior_bids.as_str();
        let bid2 = left(&r.bid, 2);
        let (min_nt, max_nt) = if prior.starts_with("1N")
            || (spec.starts_with("min") && bid2 == "1N")
            || spec.starts_with("hcp1n")
        {
            (self.nt.one_min, self.nt.one_max)
        } else if prior.starts_with("2N")
            || (spec.starts_with("min") && bid2 == "2N")
            || spec.starts_with("hcp2n")
        {
            (self.nt.two_min, self.nt.two_max)
        } else {
            (0, 0)
        };
        let have = self.hand.hcp[0];
        if spec == "min" {
            spec = min_nt.to_string();
        } else if spec.starts_with("min") {
            let n = ival(right(&spec, 2));
            spec = if n == 0 {
                format!("{}-{max_nt}", min_nt + 1)
            } else {
                (min_nt + n).to_string()
            };
        } else if spec == "hcp1n" || spec == "hcp2n" {
            spec = format!("{min_nt}-{max_nt}");
        } else if spec.starts_with('p') {
            if spec.len() > 2 {
                let body = spec[1..].to_string();
                spec = if let Some(t) = body.strip_prefix('<') {
                    format!("<{}", ival(t) - min_nt)
                } else {
                    let r1 = (ival(&body) - min_nt).to_string();
                    if body.len() > 3 {
                        format!("{r1}-{}", ival(right(&body, 2)) - min_nt)
                    } else {
                        format!("{r1}{}", right(&body, 1))
                    }
                };
            }
        } else if spec.starts_with("1n+") {
            let lo = self.nt.one_min + ival(mid_from(&format!("{spec} "), 4));
            let hi = self.nt.one_min + ival(mid_from(&format!("{spec}    "), 6));
            if hi <= lo || have < lo || have > hi {
                return Err(format!("{have} HCP, needs {lo}-{hi}"));
            }
            return Ok(());
        }
        let mut lo = ival(&spec);
        let mut hi = 0;
        if let Some(t) = spec.strip_prefix('<') {
            lo = 0;
            hi = ival(t) - 1;
        } else if spec.starts_with('>') {
            hi = 40;
        }
        if let Some(dash) = spec.find('-') {
            if dash + 1 == spec.len() {
                hi = ival(&spec);
                lo = 0;
            } else {
                hi = ival(&spec[dash + 1..]);
            }
        } else if spec.ends_with('+') {
            hi = 40;
        } else if !spec.starts_with('<') && !spec.starts_with('>') {
            hi = lo;
        }
        if have >= lo && have <= hi {
            Ok(())
        } else {
            Err(format!("{have} HCP, needs {lo}-{hi}"))
        }
    }

    /// 2026 `CustPoints_Ok`: only a points total is tested (the honour-
    /// count branch cannot be reached).
    pub(crate) fn custom_points_2026(&mut self, r: &Rule) -> Check {
        let field = r.cust_pts.as_str();
        if field.is_empty() {
            return Ok(());
        }
        let cp = [
            0.0,
            val(&r.a_cust),
            val(&r.k_cust),
            val(&r.q_cust),
            val(&r.j_cust),
            val(&r.t_cust),
        ];
        let mut reqd = val(field);
        let mut less = false;
        if reqd == 0.0 && field.len() > 1 {
            reqd = val(&field[1..]);
            less = field.starts_with('<');
        }
        if reqd == 0.0 {
            return Ok(());
        }
        let mut total = 0.0;
        for s in 1..=4 {
            for card in 10..=14 {
                if self.hand.held[s][card] {
                    total += cp[15 - card];
                }
            }
        }
        let ok = if less { total < reqd } else { total >= reqd };
        if ok {
            Ok(())
        } else {
            Err(format!("{total} custom points, needs {field}"))
        }
    }

    /// 2026 `DistribPts_Ok`: as 2018 plus "Rule20", without the
    /// intermediates credit.
    pub(crate) fn distribution_points_2026(&mut self, r: &Rule) -> Check {
        let tdp = r.dist_pts.trim();
        if tdp.is_empty() {
            return Ok(());
        }
        if tdp == "Rule20" {
            let (mut ls1, mut ls2, mut hcp) = (0, 0, 0);
            for s in 1..=4 {
                hcp += self.hand.hcp[s];
                let len = self.hand.len[s];
                if len > ls1 {
                    ls1 = len;
                } else if len > ls2 {
                    ls2 = len;
                }
            }
            if hcp + ls1 + ls2 >= 20 {
                return Ok(());
            }
        }
        let mut tdpn = ival(tdp);
        let mut sign = String::new();
        if tdpn == 0 && tdp.len() > 1 {
            tdpn = ival(&tdp[1..]);
            sign = tdp[..1].to_string();
        }
        let mut total = false;
        if sign == "T" {
            total = true;
            if ival(mid(tdp, 2, 1)) == 0 {
                sign = mid(tdp, 2, 1).to_string();
                tdpn = ival(mid_from(tdp, 3));
            } else if ival(right(tdp, 2)) < 10 {
                sign = right(tdp, 1).to_string();
            }
        }
        let dash = instr(tdp, "-");
        let (lo, hi) = if sign == "<" {
            (0, tdpn - 1)
        } else if sign == ">" {
            (tdpn + 1, 40)
        } else if tdp.ends_with('+') {
            (tdpn, 40)
        } else if dash > 0 {
            if dash == tdp.len() {
                (0, tdpn)
            } else {
                (tdpn, ival(mid_from(tdp, dash + 1)))
            }
        } else if tdpn > 0 {
            (tdpn, tdpn)
        } else {
            (0, 0)
        };
        let have = self.hand.distrib[0] + if total { self.hand.hcp[0] } else { 0 };
        if have >= lo && have <= hi {
            Ok(())
        } else {
            Err(format!("{have} points, needs {lo}-{hi} ({tdp})"))
        }
    }

    /// 2026 `Losers_Ok`: a number is a maximum, "n+" a minimum; "<n"
    /// never matches.
    pub(crate) fn losers_2026(&mut self, r: &Rule) -> Check {
        let spec = r.losers.trim();
        let have = self.hand.losers[0] as f64;
        let ok = if spec.ends_with('+') {
            have >= val(spec)
        } else {
            spec.is_empty() || (have <= val(spec) && val(spec) != 0.0)
        };
        if ok {
            Ok(())
        } else {
            Err(format!("{have} losers, needs {spec} (a maximum)"))
        }
    }

    /// 2026 `Intermed_Ok`: "0" means none at all.
    pub(crate) fn intermed_2026(&mut self, r: &Rule) -> Check {
        if r.intermed.is_empty() {
            return Ok(());
        }
        let want = ival(&r.intermed);
        let have = self.hand.intermediates[0];
        let ok = if r.intermed == "0" {
            have == 0
        } else {
            have >= want
        };
        if ok {
            Ok(())
        } else {
            Err(format!("{have} intermediates, needs {}", r.intermed))
        }
    }

    /// 2026 `Vul_Ok`: the rule's one-letter code against the client's.
    pub(crate) fn vul_2026(&mut self, r: &Rule) -> Check {
        let want = r.vul.as_str();
        let have = self.req.vul.code_2026();
        let ok = want == have
            || want.is_empty()
            || (want == "X" && (instr("UV", have) == 0 || have.is_empty()));
        if ok {
            Ok(())
        } else {
            Err(format!("vulnerability {have:?}, needs {want}"))
        }
    }

    /// 2026 `Position_Ok` with `Position_calc`.
    pub(crate) fn position_2026(&mut self, r: &Rule) -> Check {
        let want = r.pos.trim().to_ascii_uppercase();
        if want.is_empty() {
            return Ok(());
        }
        let a = self.auction;
        // Position_calc: the seat from the full prior calls, and which of
        // the first four calls (each followed by a dash) were passes.
        let mut c = [""; 5];
        let mut pos_num = 0usize;
        if !a.prior_calls.is_empty() {
            let n = a.calls.len() - 1;
            for k in 1..n {
                if k < 5 && a.call(k as i32).ends_with('P') {
                    c[k] = "P";
                }
            }
            pos_num = n - 1 + 2;
            if pos_num > 4 {
                pos_num %= 4;
                if pos_num == 0 {
                    pos_num = 4;
                }
            }
        }
        if pos_num == 0 {
            pos_num = 1;
        }
        let ok = match want.as_str() {
            "1" | "F" => pos_num == 1,
            "2" => pos_num == 2,
            "4" => pos_num == 4,
            "3" | "T" => pos_num == 3,
            "Y" => c[pos_num] == "P",
            "N" => a.call_num < 5 || a.call(pos_num as i32) != "P",
            "P" => {
                if pos_num > 2 {
                    c[pos_num - 2] == "P"
                } else {
                    c[pos_num + 2] == "P"
                }
            }
            "U" => {
                let p = if pos_num < 3 { pos_num + 2 } else { pos_num };
                c[p - 2] != "P"
            }
            _ => false,
        };
        if ok {
            Ok(())
        } else {
            Err(format!("seat {pos_num}, needs position {want}"))
        }
    }

    // ----- Suits -----------------------------------------------------------

    /// 2026 `Suits_Ok`.
    pub(crate) fn suits_2026(&mut self, r: &Rule) -> Check {
        let mut num_c = r.cm_num.trim().to_string();
        let mut pts_c = r.cm_hcp.trim().to_string();
        // "M..." combined specs (the major): `m` is never set in the VB,
        // so only these rewrites of the combined fields survive.
        if num_c.starts_with('M') {
            let mut c = mid(&num_c, 2, 1).to_string();
            if c == "=" {
                c = mid(&num_c, 3, 1).to_string();
            }
            if cdhs(&c) == 0 {
                num_c.clear();
            } else if !num_c.starts_with("M=") {
                num_c = num_c[1..].to_string();
            }
        } else if pts_c.starts_with('M') {
            let c = mid(&pts_c, 2, 1);
            if instr("CDHSx", c) == 0 || c.is_empty() {
                pts_c.clear();
            } else {
                pts_c = pts_c[1..].to_string();
            }
        }

        let (s_ok, x_found, longest_x) = self.see_if_suits_ok(r)?;
        let mut s_ok = s_ok;
        for i in 1..=4 {
            let num = r.suit_num(i);
            if !s_ok[i] && num.contains('x') && !num.contains(',') && x_found {
                s_ok[i] = true;
            }
        }
        if (1..=4).all(|i| s_ok[i]) && (x_found || r.x_num.is_empty()) {
            self.combo_2026(&num_c, &pts_c)?;
            if longest_x > 0 {
                self.x_suit = longest_x;
            }
            Ok(())
        } else {
            self.x_suit = 0;
            if !x_found && !r.x_num.is_empty() && (1..=4).all(|i| s_ok[i]) {
                return Err(format!("no x suit for {}/{}", r.x_num, r.x_hcp));
            }
            // Suits after the first failure were never tested: name the
            // first one in testing order that failed.
            let order: Vec<usize> = if r.x_num.trim().ends_with('U') {
                vec![1, 2, 3, 4]
            } else {
                vec![4, 3, 2, 1]
            };
            let first = order.into_iter().find(|&i| !s_ok[i]).unwrap_or(4);
            // Mention the x specs only for a rule that has them.
            let x = if r.x_num.trim().is_empty() && r.x_hcp.trim().is_empty() {
                String::new()
            } else {
                format!(
                    " (x {}/{}{})",
                    r.x_num,
                    r.x_hcp,
                    if x_found { "" } else { ", no x suit found" }
                )
            };
            Err(format!(
                "{}: \"{}\" fails {}/{}{x}",
                crate::vb::suit_letter(first),
                self.hand.suit[first],
                r.suit_num(first),
                r.suit_hcp(first),
            ))
        }
    }

    /// The `SeeIfSuitsOk` subroutine of `Suits_Ok`: each suit's specs,
    /// with the x suit and alternatives. Returns which suits passed,
    /// whether an x suit was found, and the longest x suit.
    fn see_if_suits_ok(&mut self, r: &Rule) -> Result<([bool; 5], bool, usize), String> {
        let hand = self.hand;
        let mut db_num: Vec<String> = (0..=4).map(|s| r.suit_num(s).to_string()).collect();
        let mut db_hcp: Vec<String> = (0..=4).map(|s| r.suit_hcp(s).to_string()).collect();
        for s in 1..=4 {
            if db_hcp[s].contains("Trmp") {
                self.trump = s;
                db_hcp[s] = if db_hcp[s] == "Trmp" {
                    String::new()
                } else {
                    db_hcp[s].replace("Trmp", "").replace(',', "")
                };
            }
            db_num[s] = db_num[s].replace(' ', "");
            db_hcp[s] = db_hcp[s].replace(' ', "");
        }
        let mut x_suit = 0usize;
        let mut x_found = false;
        let mut x_num = r.x_num.trim().to_string();
        let order: Vec<usize> = if x_num.ends_with('U') {
            x_num.pop();
            vec![1, 2, 3, 4]
        } else {
            vec![4, 3, 2, 1]
        };
        let x_hcp = r.x_hcp.trim().to_string();
        let mut s_ok = [false; 5];
        let mut longest_x = 0usize;
        let mut compare: Option<(usize, String, String)> = None;

        for (k, &s) in order.iter().enumerate() {
            let mut num1 = db_num[s].clone();
            let mut num2 = String::new();
            let mut pts1 = db_hcp[s].clone();
            let mut pts2 = String::new();
            if num1.eq_ignore_ascii_case("x") || pts1.eq_ignore_ascii_case("x") {
                num1 = x_num.clone();
                pts1 = x_hcp.clone();
            }
            if db_num[s].to_ascii_lowercase().starts_with('x')
                && db_num[s].contains(',')
                && !db_hcp[s].to_ascii_lowercase().starts_with('x')
                && !db_hcp[s].is_empty()
                && !db_hcp[s].contains(',')
            {
                pts1 = format!(" ,{pts1}");
            }
            let i = instr(&num1, ",");
            let ix = instr(&x_num, ",");
            if i > 0 {
                if ix > 0 && i != ix {
                    num2 = left(&num1, ix - 1).trim().to_string();
                } else {
                    num2 = mid_from(&num1, i + 1).trim().to_string();
                }
                num1 = if num1.contains('x') {
                    if ix == 0 {
                        x_num.clone()
                    } else {
                        left(&x_num, i - 1).trim().to_string()
                    }
                } else {
                    left(&num1, i - 1).to_string()
                };
                if num1.is_empty() {
                    num1 = num2.clone();
                }
            }
            let i = instr(&pts1, ",");
            if i > 0 {
                pts2 = mid_from(&pts1, i + 1).trim().to_string();
                pts1 = left(&pts1, i - 1).trim().to_string();
            }
            if num1.is_empty() && pts1.is_empty() {
                s_ok[s] = true;
                continue;
            }
            // The VB loop; bounded here, as its alternatives shrink.
            for _ in 0..16 {
                s_ok[s] = self.suit_ok_2026(s, &num1, &pts1);
                if !s_ok[s] {
                    if !num2.is_empty() {
                        num1 = num2.clone();
                        match num1.find(',') {
                            Some(p) => {
                                num2 = num1[p + 1..].trim().to_string();
                                num1 = num1[..p].trim().to_string();
                            }
                            None => num2.clear(),
                        }
                        pts1 = pts2.clone();
                        s_ok[s] = self.suit_ok_2026(s, &num1, &pts1);
                    } else if !pts2.is_empty() {
                        pts1 = pts2.clone();
                        match pts1.find(',') {
                            Some(p) => {
                                pts2 = pts1[p + 1..].trim().to_string();
                                pts1 = pts1[..p].trim().to_string();
                            }
                            None => pts2.clear(),
                        }
                    } else {
                        break;
                    }
                } else {
                    if instr(&db_num[s], "x") == 1 && db_num[s] != format!("x,{num1}") {
                        x_suit = s;
                        if x_num.contains('<') {
                            for &later in &order[k + 1..] {
                                db_num[later] = db_num[later].replace('x', "");
                            }
                            break;
                        }
                        if longest_x == 0 || hand.len[s] > hand.len[longest_x] {
                            longest_x = s;
                        }
                    } else if num1.to_ascii_lowercase().ends_with('x')
                        || pts1.to_ascii_lowercase().ends_with('x')
                    {
                        compare = Some((s, num1.clone(), pts1.clone()));
                    }
                    break;
                }
            }
            if x_suit != 0 {
                x_found = true;
            }
            if !s_ok[s] && db_num[s] != "x" {
                break;
            }
        }
        if longest_x > 0 {
            x_suit = longest_x;
        }
        self.x_suit = x_suit;
        if x_suit == 0 && (!r.x_num.trim().is_empty() || !r.x_hcp.trim().is_empty()) {
            return Ok((s_ok, false, longest_x));
        }
        if let (Some((cs, num, pts)), true) = (compare, x_suit > 0) {
            for (spec, mine, theirs) in [
                (num, hand.len[cs], hand.len[x_suit]),
                (pts, hand.hcp[cs], hand.hcp[x_suit]),
            ] {
                if spec.is_empty() {
                    continue;
                }
                let op = if spec.len() == 3 {
                    &spec[..2]
                } else {
                    &spec[..1]
                };
                s_ok[cs] = match op {
                    "<" => mine < theirs,
                    ">" => mine > theirs,
                    "=>" | ">=" => mine >= theirs,
                    "=<" | "<=" => mine <= theirs,
                    _ => s_ok[cs],
                };
            }
        }
        Ok((s_ok, x_found, longest_x))
    }

    /// 2026 `Suit_Ok`: a length spec is required; "or..." in the strength
    /// field means length or strength.
    fn suit_ok_2026(&self, s: usize, num: &str, hcps: &str) -> bool {
        if num.is_empty() || num.to_ascii_lowercase().ends_with('x') {
            return false;
        }
        let q = self.suit_qty_2026(s, num);
        if q && hcps.is_empty() {
            return true;
        }
        let or = left(hcps, 2).eq_ignore_ascii_case("or");
        if !(q || or) {
            return false;
        }
        let hcps = hcps.replace("OR", "or");
        if hcps.is_empty() || q || (hcps.starts_with("or") && hcps.len() > 2) {
            let or_hcp = hcps.strip_prefix("or").unwrap_or(&hcps);
            let p = self.suit_hcp_2026(s, or_hcp);
            let has_or = hcps.starts_with("or");
            if (!has_or && (!q || !p)) || (has_or && !q && !p) {
                return false;
            }
        }
        true
    }

    /// 2026 `SuitQty_Ok`.
    fn suit_qty_2026(&self, s: usize, num: &str) -> bool {
        let len = self.hand.len[s];
        let mut nm = ival(num);
        let sign_r = right(num, 1);
        let sign_l = left(num, 1);
        if ival(sign_l) == 0 && !num.is_empty() {
            nm = ival(mid_from(num, 2));
        }
        let other = |t: &str| self.hand.len[cdhs(t)];
        if num == "0" {
            len == 0
        } else if sign_r == "+" {
            len >= nm
        } else if sign_r == "-" {
            len < nm
        } else if sign_l == "<" {
            if left(num, 2) == "<=" {
                match cdhs(mid(num, 3, 1)) {
                    0 => len <= ival(mid(num, 3, 1)),
                    s2 => len <= self.hand.len[s2],
                }
            } else {
                match cdhs(mid(num, 2, 1)) {
                    0 => len < ival(mid(num, 2, 1)),
                    _ => len < other(mid(num, 2, 1)),
                }
            }
        } else if sign_l == ">" {
            if left(num, 2) == ">=" {
                match cdhs(mid(num, 3, 1)) {
                    0 => len >= ival(mid(num, 2, 1)),
                    s2 => len >= self.hand.len[s2],
                }
            } else {
                match cdhs(mid(num, 2, 1)) {
                    0 => len > ival(mid(num, 2, 1)),
                    s2 => len > self.hand.len[s2],
                }
            }
        } else if sign_l == "=" {
            match cdhs(mid(num, 2, 1)) {
                0 => len == ival(mid(num, 2, 1)),
                s2 => len == self.hand.len[s2],
            }
        } else if let Some(p) = num.find('-') {
            let mut nm2 = ival(&num[p + 1..]);
            if nm2 == 0 {
                let i = cdhs(&num[p + 1..]);
                if i > 0 {
                    nm2 = self.hand.len[i];
                }
            }
            len >= nm && len <= nm2
        } else {
            nm > 0 && len == nm
        }
    }

    /// 2026 `SuitHCP_Ok`.
    fn suit_hcp_2026(&self, s: usize, spec: &str) -> bool {
        let hand = self.hand;
        let suit = hand.suit[s].as_str();
        let has = |r: usize| hand.held[s][r];
        let mut hcps = spec.to_string();
        if hcps == "SS" && hand.solid[s] {
            return true;
        }
        if hcps.contains("<A") {
            if suit.starts_with('A') {
                return false;
            }
            hcps = hcps.replace("<A", "");
        } else if hcps.contains("<K") {
            if suit.starts_with('A') || suit.starts_with('K') {
                return false;
            }
            hcps = hcps.replace("<K", "");
        }
        if left(&hcps, 2).eq_ignore_ascii_case("or") {
            if hcps.len() < 4 {
                return false;
            }
            hcps = hcps[2..].to_string();
        }
        hcps = hcps.replace(' ', "");
        if hcps.is_empty() {
            return true;
        }
        let original = hcps.clone();
        let mut hcps = hcps.to_ascii_lowercase();
        if hcps == "hh" {
            return suit.len() < 3 || suit.starts_with('A') || suit.starts_with('K');
        }
        if hcps == "ha" {
            return suit.len() > 2;
        }
        let mut intermed_reqd = 0;
        if let Some(p) = hcps.find('i') {
            intermed_reqd = ival(&hcps[p + 1..]);
            if p == 0 {
                if hand.hcp[s] > 4 {
                    return true;
                }
                if hand.hcp[s] > 0 {
                    intermed_reqd -= 1;
                }
                return hand.intermediates[s] >= intermed_reqd;
            }
            hcps.truncate(p);
        }
        let mut ok = false;
        let mut extra = false;
        if !hcps.is_empty() {
            let sign_r = right(&hcps, 1).to_string();
            let sign_l = left(&original, 1).to_string();
            let padded = format!("{suit} ");
            let left_suit = |n: usize| left(&padded, n).to_string();
            if original == "P" {
                ok = true;
                if left_suit(2) == "AQ" || left_suit(1) == "K" || left_suit(1) == "Q" {
                    let l3 = left_suit(3);
                    if l3 != "KQJ" && l3 != "QJT" && l3 != "AQJ" {
                        let a = self.auction;
                        if a.bid_num == 4 {
                            let last = |b: &str| cdhs(right(b, 1));
                            if last(a.call(a.bid_num - 3)) == s || last(a.call(a.bid_num - 1)) != s
                            {
                                ok = false;
                            }
                        }
                    }
                }
            } else if hcps == "1stop" {
                ok = hand.stops[s] >= 1.0;
                extra = hand.stops[s] > 1.0;
            } else if hcps == "2stop" {
                ok = hand.stops[s] > 1.0;
                extra = hand.stops[s] > 2.0;
            } else if hcps == "hstop" {
                ok = hand.stops[s] > 0.0;
                extra = hand.stops[s] > 0.5;
            } else if hcps == "0stop" {
                return hand.stops[s] < 1.0;
            } else if original.contains('A') {
                ok = has(ACE);
                if ok && !original.starts_with('A') {
                    let i = instr(&original, "A");
                    let before = mid(&original, i - 1, 1);
                    if !(before == " " || before == ",") {
                        ok = self.suit_hcp_2026(s, left(&original, i - 1));
                    }
                }
                if has(KING) && original.ends_with('K') {
                    ok = true;
                }
            } else if hcps.contains('k') {
                ok = has(KING);
            } else if hcps.contains('q') {
                ok = has(QUEEN);
            } else if instr(&hcps, "/") == 2 {
                // "n/m": n of the top m honours.
                let m = ival(mid(&hcps, 3, 1));
                let mut j = i32::from(has(ACE));
                let mut card = 13i32;
                while card >= 13 - m + 2 {
                    if card >= 2 && has(card as usize) {
                        j += 1;
                    }
                    card -= 1;
                }
                ok = j >= ival(left(&hcps, 1));
            } else if hcps == "frc" {
                ok = hand.len[s] == 0 || has(ACE);
            } else if hcps == "src" {
                ok = hand.len[s] == 1 || has(KING);
            }
            if ok {
                return true;
            }
            let mut ok2 = true;
            let mut i = instr(&original, "H");
            if i > 0 {
                let h = ival(mid(&format!("{original} "), i + 1, 1));
                let mut t = 0;
                for r in [ACE, KING, QUEEN, JACK] {
                    t += i32::from(has(r));
                }
                if has(10) {
                    t += 1;
                    if has(9) {
                        t += 1;
                        if has(8) {
                            t += 1;
                        }
                    }
                }
                if t < h {
                    ok2 = false;
                }
            }
            if sign_l == "<" {
                ok = hand.hcp[s] < ival(mid_from(&hcps, 2));
            } else if (ival(&hcps) > 0 && mid_from(&hcps, 2) != "stop") || hcps == "0" {
                let need = ival(&hcps);
                if sign_r == "+" {
                    ok = hand.hcp[s] >= need;
                    extra = hand.hcp[s] > need;
                    intermed_reqd = (intermed_reqd - (hand.hcp[s] - need) / 2).max(0);
                } else if sign_r == "-" {
                    ok = hand.hcp[s] <= need;
                } else if mid(&hcps, 2, 1) == "-" {
                    ok = hand.hcp[s] >= need && hand.hcp[s] <= ival(mid_from(&hcps, 3));
                } else if !hcps.contains('/') {
                    ok = hand.hcp[s] == need;
                }
            } else if !hcps.contains("stop") {
                let left1 = left(&original, 1);
                let left2 = left(&format!("{original} "), 2).to_string();
                let mut list = String::new();
                let in_list = |list: &str, top: String| instr(list, &top) > 0;
                if left2 == "ab" {
                    list = "AK,AQ,KQ".into();
                    ok = in_list(&list, left_suit(2));
                    if !ok && !left_suit(3).is_empty() {
                        list = "AJT,KJT,QJT".into();
                        ok = in_list(&list, left_suit(3));
                    }
                } else if left2 == "bc" {
                    list = "AJT,KJT,QJT,KQJ,AQJ,AKQ,AKJ,AKQJ".into();
                    ok = in_list(&list, left_suit(3));
                } else if left1 == "d" {
                    list = "AKQJ,KQJT,AKQT,AKQJT".into();
                    ok = in_list(&list, left_suit(4));
                } else if left1 == "c" {
                    list = "AKQ,AKJ,AQJ,KQJ".into();
                    ok = in_list(&list, left_suit(3));
                } else if left1 == "a" || left2 == "ac" {
                    list = "AK,AQ,KQ".into();
                    ok = in_list(&list, left_suit(2));
                } else if left1 == "b" {
                    list = "AJT,KJT,QJT".into();
                    ok = in_list(&list, left_suit(3));
                } else if left1 == "H" {
                    let h = ival(mid(&hcps, 2, 1));
                    i = ival(mid(&format!("{hcps} "), 3, 1)).max(0) as usize;
                    if i == 0 {
                        i = h.max(0) as usize;
                    }
                    match h {
                        2 => {
                            list = match i {
                                3 => "AK,AQ,KQ",
                                4 => "AK,AQ,AJ,KQ,KJ,QJ",
                                5 => "AK,AQ,AJ,AT,KQ,KJ,KT,QJ,QT",
                                _ => "",
                            }
                            .into();
                            ok = in_list(&list, left_suit(2));
                        }
                        3 => {
                            if i > 2 {
                                list = "AKQ".into();
                            }
                            if i > 3 {
                                list += ",AKJ,AQJ,KQJ";
                            }
                            if i > 4 {
                                list += ",AKT,AQT,AJT,KQT,QJT";
                            }
                            if i > 5 {
                                list += ",AK9,AQ9,AJ9,KQ9,KJ9,KT9,QJ9";
                            }
                            ok = in_list(&list, left_suit(3));
                        }
                        4 => {
                            if i > 3 {
                                list = "AKQJ".into();
                            }
                            if i > 4 {
                                list += ",AKQT,AKJT,AQJT,KQJT";
                            }
                            if i > 5 {
                                list +=
                                    ",AKQ9,AQT9,AJT9,KQT9,QJT9,AKJ9,AQJ9,AJT9,KQT9,KJT9,KT9,QJ9";
                            }
                            ok = in_list(&list, left_suit(4));
                        }
                        5 => {
                            list = "AKQJT,AKQJ9,AKQT9,AKJT9,AQJT9,KQJT9".into();
                            ok = in_list(&list, left_suit(5));
                        }
                        _ => {}
                    }
                }
                if i > 0 {
                    let n = ival(mid(&left_suit(4), i, 1));
                    if n == 9 || n == 0 {
                        ok = in_list(&list, left(&left_suit(4), i).to_string());
                    }
                }
            }
            ok = ok && ok2;
        }
        if (ok || hcps.is_empty()) && intermed_reqd > 0 {
            if extra {
                intermed_reqd -= 1;
            }
            ok = hand.intermediates[s] >= intermed_reqd;
        }
        ok
    }

    /// 2026 `Combo_Ok`: combined lengths, then ace-asking replies
    /// ("2A+TK-TQ", from `Calculate_CustPoints`) or combined points.
    fn combo_2026(&self, num1: &str, pts1: &str) -> Check {
        if num1.is_empty() && pts1.is_empty() {
            return Ok(());
        }
        let hand = self.hand;
        let fail = |why: String| Err(why);
        if num1.len() > 3 {
            let s1 = cdhs(mid(num1, 1, 1));
            if s1 > 0 {
                let s2 = cdhs(mid(num1, 2, 1));
                let s3 = cdhs(mid(num1, 3, 1));
                let (sign, total, ttl) = if s3 > 0 {
                    (
                        mid(num1, 4, 1),
                        hand.len[s1] + hand.len[s2] + hand.len[s3],
                        ival(mid_from(num1, 5)),
                    )
                } else {
                    (
                        mid(num1, 3, 1),
                        hand.len[s1] + hand.len[s2],
                        ival(mid_from(num1, 4)),
                    )
                };
                let ok = match sign {
                    "=" => match num1.find('-') {
                        Some(d) => total >= ttl && total <= ival(&num1[d + 1..]),
                        None => total == ttl,
                    },
                    "<" => total < ttl,
                    ">" => total > ttl,
                    _ => true,
                };
                if !ok {
                    return fail(format!("combined length {total} fails {num1}"));
                }
            }
        }
        if pts1.is_empty() {
            return Ok(());
        }
        if ["T", "A", "K", "Q"].iter().any(|c| pts1.contains(c)) {
            let pts = pts1.replace(' ', "");
            let (aces, tk, tq) = self.ace_counts();
            let s = format!("{aces}A{}", if tk { "+TK" } else { "-TK" });
            let ok = if pts.contains(&format!("{s}-TQ")) {
                !tq
            } else if pts.contains(&format!("{s}+TQ")) {
                tq
            } else {
                pts.contains(&s)
            };
            return if ok {
                Ok(())
            } else {
                fail(format!(
                    "{s}{}: none of {pts}",
                    if tq { "+TQ" } else { "-TQ" }
                ))
            };
        }
        let x = self.x_suit;
        let s1 = match instr("CDHSx", left(pts1, 1)) {
            5 => x,
            n => n,
        };
        if pts1.len() > 3 && s1 != 0 {
            let s2 = match instr("CDHSx", mid(&format!("{pts1}   "), 2, 1)) {
                5 if x > 0 => x,
                5 => 0,
                n => n,
            };
            let s3 = cdhs(mid(&format!("{pts1}    "), 3, 1));
            let (sign, total, ttl) = if s3 > 0 {
                (
                    mid(&format!("{pts1}     "), 4, 1).to_string(),
                    hand.hcp[s1] + hand.hcp[s2] + hand.hcp[s3],
                    ival(mid_from(pts1, 5)),
                )
            } else {
                (
                    mid(&format!("{pts1}    "), 3, 1).to_string(),
                    hand.hcp[s1] + hand.hcp[s2],
                    ival(mid_from(pts1, 4)),
                )
            };
            let ok = match sign.as_str() {
                // The VB reads the range's end from the length spec.
                "=" => match num1.find('-') {
                    Some(d) => total >= ttl && total <= ival(&num1[d + 1..]),
                    None => total == ttl,
                },
                "<" => total < ttl,
                ">" => total > ttl,
                _ => true,
            };
            return if ok {
                Ok(())
            } else {
                fail(format!("combined HCP {total} fails {pts1}"))
            };
        }
        fail(format!("combined spec {pts1:?} not understood"))
    }

    /// `Calculate_CustPoints`: aces, and the trump king and queen.
    pub(crate) fn ace_counts(&self) -> (i32, bool, bool) {
        let hand = self.hand;
        let (mut aces, mut void_ace) = (0, 0);
        for s in 1..=4 {
            if hand.suit[s].starts_with('A') {
                aces += 1;
            }
        }
        // Exclusion Blackwood ("4H@XC"): the ace of the void suit
        // doesn't count.
        let prior = self.auction.prior_bids.as_str();
        if let Some(n) = prior.find('@') {
            if mid(prior, n + 2, 1) == "X" && n >= 1 {
                let void = cdhs(mid(prior, n, 1));
                if void > 0 && hand.suit[void].starts_with('A') {
                    void_ace = 1;
                }
            }
        }
        if aces > 0 {
            aces -= void_ace;
        }
        // Counted in AnalyzeHand, before any rule's "Trmp" moves the trump.
        let t = self.trump0;
        let (tk, tq) = if t > 0 {
            (hand.suit[t].contains('K'), hand.suit[t].contains('Q'))
        } else {
            (false, false)
        };
        (aces, tk, tq)
    }
}

/// The trump suit `Calculate_CustPoints` reads from the "@" code of the
/// prior bids when the client gives none: "4N@30" is suit 3, "4H@XC"
/// (exclusion) the suit after the X, "5C@Mw" (Minorwood) the minor bid.
pub fn trump_from_prior(prior: &str) -> usize {
    let Some(n) = prior.find('@') else { return 0 };
    let n = n + 1; // VB position of '@'
    if n < 2 {
        return 0;
    }
    let before = mid(prior, n - 1, 1);
    let after = mid(prior, n + 1, 1);
    if before == "N" {
        ival(after).clamp(0, 4) as usize
    } else if after == "X" {
        cdhs(mid(prior, n + 2, 1))
    } else if mid(prior, n + 1, 2) == "Mw" {
        instr("CD", before).min(2)
    } else {
        0
    }
}
