//! Decode a form's `.ico` (from its `.frx`) for the window icon: the first
//! image, 32- or 24-bit BMP with its AND mask.

/// (width, height, RGBA pixels top-down).
pub fn decode(ico: &[u8]) -> Option<(u32, u32, Vec<u8>)> {
    let u16_at = |at: usize| Some(u16::from_le_bytes(ico.get(at..at + 2)?.try_into().ok()?));
    let u32_at = |at: usize| Some(u32::from_le_bytes(ico.get(at..at + 4)?.try_into().ok()?));
    if u16_at(2)? != 1 || u16_at(4)? == 0 {
        return None;
    }
    let off = u32_at(6 + 12)? as usize;
    let dib = ico.get(off..)?;
    let d32 = |at: usize| Some(u32::from_le_bytes(dib.get(at..at + 4)?.try_into().ok()?));
    let header = d32(0)? as usize;
    let w = d32(4)? as usize;
    let h = d32(8)? as usize / 2; // XOR image plus AND mask
    let bpp = u16::from_le_bytes(dib.get(14..16)?.try_into().ok()?) as usize;
    if bpp != 32 && bpp != 24 {
        return None;
    }
    let row = (w * bpp / 8).div_ceil(4) * 4;
    let mask_row = w.div_ceil(32) * 4;
    let pixels = header;
    let mask = pixels + row * h;
    let mut out = vec![0u8; w * h * 4];
    for y in 0..h {
        let src = pixels + (h - 1 - y) * row;
        let msrc = mask + (h - 1 - y) * mask_row;
        for x in 0..w {
            let p = src + x * bpp / 8;
            let (b, g, r) = (*dib.get(p)?, *dib.get(p + 1)?, *dib.get(p + 2)?);
            let a = if bpp == 32 {
                *dib.get(p + 3)?
            } else {
                let m = *dib.get(msrc + x / 8)?;
                if m & (0x80 >> (x % 8)) != 0 {
                    0
                } else {
                    255
                }
            };
            let o = (y * w + x) * 4;
            out[o..o + 4].copy_from_slice(&[r, g, b, a]);
        }
    }
    Some((w as u32, h as u32, out))
}
