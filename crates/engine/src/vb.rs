//! The few Visual Basic 6 string functions the rule specs lean on.
//!
//! The specs were written for VB's forgiving `Val`, so "6+" is 6, "<10" is
//! 0 and "24-25" is 24. Porting the tests literally needs the same reading.

/// VB6 `Val`: the number at the start of `s` (after blanks), or 0.
pub fn val(s: &str) -> f64 {
    let s = s.trim_start();
    let mut end = 0;
    let bytes = s.as_bytes();
    if end < bytes.len() && (bytes[end] == b'+' || bytes[end] == b'-') {
        end += 1;
    }
    let digits_start = end;
    let mut seen_dot = false;
    while end < bytes.len() {
        match bytes[end] {
            b'0'..=b'9' => end += 1,
            b'.' if !seen_dot => {
                seen_dot = true;
                end += 1
            }
            _ => break,
        }
    }
    if end == digits_start {
        return 0.0;
    }
    s[..end].parse().unwrap_or(0.0)
}

/// `Val` truncated to an integer, as the VB code stores it in an `Integer`.
pub fn ival(s: &str) -> i32 {
    val(s) as i32
}

/// VB's conversion of a number to `Integer` (`CInt`): round half to even.
pub fn cint(x: f64) -> i32 {
    let r = x.round();
    if (x - x.trunc()).abs() == 0.5 && r % 2.0 != 0.0 {
        (r - x.signum()) as i32
    } else {
        r as i32
    }
}

/// VB `InStr(haystack, needle) > 0`, where an empty needle matches.
pub fn instr(haystack: &str, needle: &str) -> usize {
    if needle.is_empty() {
        return 1;
    }
    haystack.find(needle).map_or(0, |i| i + 1)
}

/// The text before the first comma and the text after it, both trimmed.
/// The suit specs use commas for alternatives: "x,4+" is "x, else 4+".
pub fn split_alt(s: &str) -> (String, String) {
    match s.find(',') {
        Some(i) => (s[..i].trim().to_string(), s[i + 1..].trim().to_string()),
        None => (s.trim().to_string(), String::new()),
    }
}

/// Index 1-4 of a suit letter in "CDHS" (VB `InStr("CDHS", c)`), or 0.
pub fn suit_index(c: char) -> usize {
    match c {
        'C' => 1,
        'D' => 2,
        'H' => 3,
        'S' => 4,
        _ => 0,
    }
}

/// The suit letter for index 1-4.
pub fn suit_letter(s: usize) -> char {
    ['?', 'C', 'D', 'H', 'S'][s.min(4)]
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn val_reads_a_leading_number() {
        assert_eq!(val("6+"), 6.0);
        assert_eq!(val("24-25"), 24.0);
        assert_eq!(val("<10"), 0.0);
        assert_eq!(val(" 7.5"), 7.5);
        assert_eq!(val("-1"), -1.0);
        assert_eq!(val("+2"), 2.0);
        assert_eq!(val(""), 0.0);
        assert_eq!(val("x"), 0.0);
    }

    #[test]
    fn cint_rounds_half_to_even() {
        assert_eq!(
            (cint(0.5), cint(1.5), cint(2.5), cint(3.0), cint(2.4)),
            (0, 2, 2, 3, 2)
        );
    }

    #[test]
    fn split_alt_splits_on_the_first_comma() {
        assert_eq!(split_alt("x, <4"), ("x".into(), "<4".into()));
        assert_eq!(split_alt("4+"), ("4+".into(), "".into()));
        assert_eq!(split_alt("<1,a,A"), ("<1".into(), "a,A".into()));
    }
}
