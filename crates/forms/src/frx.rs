//! Reading values a `.frm` keeps in its `.frx` (binary resources).
//!
//! Layouts, as found in Nelson's files:
//! - text (`Text`, `Caption`): 1-byte length, then the bytes; a length
//!   byte of 0xFF is followed by a 2-byte length.
//! - long text (`$"x.frx":off`, e.g. an ADO `RecordSource`): 4-byte length.
//! - `List`: 2-byte count, then each item as a 2-byte length and bytes.
//!   `ItemData`: 2-byte count, then 4-byte integers.
//! - `Picture`, `Icon`: 4-byte total, `lt\0\0`, 4-byte size, then the
//!   image file (BMP, ICO, GIF or JPEG).
//!
//! All integers little-endian; text in Windows-1252.

use crate::frm::decode_cp1252;

fn u16_at(b: &[u8], at: usize) -> Option<usize> {
    Some(u16::from_le_bytes(b.get(at..at + 2)?.try_into().ok()?) as usize)
}

fn u32_at(b: &[u8], at: usize) -> Option<usize> {
    Some(u32::from_le_bytes(b.get(at..at + 4)?.try_into().ok()?) as usize)
}

/// A text value at `offset`; `long` for the `$"..."` form.
pub fn text(frx: &[u8], offset: u32, long: bool) -> Option<String> {
    let at = offset as usize;
    let (len, start) = if long {
        (u32_at(frx, at)?, at + 4)
    } else {
        match *frx.get(at)? {
            0xFF => (u16_at(frx, at + 1)?, at + 3),
            n => (n as usize, at + 1),
        }
    };
    Some(decode_cp1252(frx.get(start..start + len)?))
}

/// A `List` property: the items of a list or combo box.
pub fn list(frx: &[u8], offset: u32) -> Option<Vec<String>> {
    let mut at = offset as usize;
    let count = u16_at(frx, at)?;
    at += 2;
    let mut out = Vec::with_capacity(count);
    for _ in 0..count {
        let len = u16_at(frx, at)?;
        at += 2;
        out.push(decode_cp1252(frx.get(at..at + len)?));
        at += len;
    }
    Some(out)
}

/// A `Picture` or `Icon`: the embedded image file's bytes.
pub fn picture(frx: &[u8], offset: u32) -> Option<&[u8]> {
    let at = offset as usize;
    if frx.get(at + 4..at + 8)? != b"lt\0\0" {
        return None;
    }
    let size = u32_at(frx, at + 8)?;
    frx.get(at + 12..at + 12 + size)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn short_and_long_text() {
        let mut b = vec![3u8];
        b.extend_from_slice(b"abc");
        assert_eq!(text(&b, 0, false).as_deref(), Some("abc"));
        let mut b = 2u32.to_le_bytes().to_vec();
        b.extend_from_slice(b"hi");
        assert_eq!(text(&b, 0, true).as_deref(), Some("hi"));
    }

    #[test]
    fn list_items() {
        let b = [2, 0, 1, 0, b'a', 2, 0, b'b', b'c'];
        assert_eq!(list(&b, 0).unwrap(), vec!["a", "bc"]);
    }
}
