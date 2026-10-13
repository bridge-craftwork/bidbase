//! Reading a VB6 form file (`.frm`): the control tree at the top of the
//! file, with each control's properties as written. Values that live in the
//! form's `.frx` (long text, lists, pictures) are kept as references and
//! read on demand with [`crate::frx`].
//!
//! The layout part of a `.frm` is a nested block:
//!
//! ```text
//! Begin VB.Form PlayScr
//!    Caption         =   "CardShark BidBase Practice Dealer"
//!    BeginProperty Font
//!       Name            =   "MS Sans Serif"
//!       Size            =   8.25
//!    EndProperty
//!    Begin VB.CommandButton b_Deal
//!       Caption         =   "&Deal"
//!       Index           =   0
//!    End
//! End
//! Attribute VB_Name = "PlayScr"
//! ```
//!
//! Positions and sizes are in twips (15 to a pixel at 96 dpi), relative to
//! the containing form, frame or picture box.

use std::collections::BTreeMap;

/// Decode Windows-1252, the code page Nelson's files are written in.
pub fn decode_cp1252(bytes: &[u8]) -> String {
    const HIGH: [char; 32] = [
        '€', '\u{81}', '‚', 'ƒ', '„', '…', '†', '‡', 'ˆ', '‰', 'Š', '‹', 'Œ', '\u{8d}', 'Ž',
        '\u{8f}', '\u{90}', '‘', '’', '“', '”', '•', '–', '—', '˜', '™', 'š', '›', 'œ', '\u{9d}',
        'ž', 'Ÿ',
    ];
    bytes
        .iter()
        .map(|&b| match b {
            0x80..=0x9f => HIGH[(b - 0x80) as usize],
            _ => b as char,
        })
        .collect()
}

/// A property value as written in the `.frm`.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    /// A quoted string, with `""` unescaped.
    Str(String),
    /// A number, a `&H...&` colour or anything else unquoted, comment removed.
    Raw(String),
    /// Stored in the `.frx` at this offset. `long` is the `$"..."` form,
    /// which has a 4-byte length.
    Frx {
        file: String,
        offset: u32,
        long: bool,
    },
}

/// A font from a `BeginProperty Font` block.
#[derive(Debug, Clone, PartialEq)]
pub struct Font {
    pub name: String,
    pub size: f32,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
}

impl Default for Font {
    fn default() -> Self {
        Font {
            name: "MS Sans Serif".into(),
            size: 8.25,
            bold: false,
            italic: false,
            underline: false,
        }
    }
}

/// One control (or the form itself, or a menu).
#[derive(Debug, Clone, Default)]
pub struct Control {
    /// `VB.CommandButton`, `MSFlexGridLib.MSFlexGrid`, ...
    pub class: String,
    pub name: String,
    pub props: BTreeMap<String, Value>,
    /// Set when the control has its own `BeginProperty Font`; otherwise it
    /// inherits from its container (see [`Form::font_of`]).
    pub font: Option<Font>,
    pub children: Vec<Control>,
}

impl Control {
    /// The class without its library: `CommandButton`, `MSFlexGrid`.
    pub fn kind(&self) -> &str {
        self.class.rsplit('.').next().unwrap_or(&self.class)
    }

    /// Element of a control array, if it is one.
    pub fn index(&self) -> Option<i32> {
        self.int("Index")
    }

    pub fn str(&self, name: &str) -> Option<&str> {
        match self.props.get(name)? {
            Value::Str(s) | Value::Raw(s) => Some(s),
            Value::Frx { .. } => None,
        }
    }

    pub fn int(&self, name: &str) -> Option<i32> {
        let s = self.str(name)?;
        if let Some(hex) = s.strip_prefix("&H") {
            return u32::from_str_radix(hex.trim_end_matches('&'), 16)
                .ok()
                .map(|v| v as i32);
        }
        s.parse::<f64>().ok().map(|v| v as i32)
    }

    /// VB booleans are -1 (True) and 0 (False); missing means `default`.
    pub fn flag(&self, name: &str, default: bool) -> bool {
        self.int(name).map(|v| v != 0).unwrap_or(default)
    }

    /// A colour property as VB stores it: `&H00BBGGRR&`, or a system colour
    /// `&H800000xx&`.
    pub fn color(&self, name: &str) -> Option<u32> {
        self.int(name).map(|v| v as u32)
    }

    /// Caption with VB's `&` accelerator markers removed (`&&` is a literal `&`).
    pub fn caption(&self) -> String {
        strip_accel(self.str("Caption").unwrap_or(""))
    }

    pub fn left(&self) -> i32 {
        self.int("Left").unwrap_or(0)
    }
    pub fn top(&self) -> i32 {
        self.int("Top").unwrap_or(0)
    }
    pub fn width(&self) -> i32 {
        self.int("Width").unwrap_or(0)
    }
    pub fn height(&self) -> i32 {
        self.int("Height").unwrap_or(0)
    }
    pub fn visible(&self) -> bool {
        self.flag("Visible", true)
    }

    /// Every control under this one, depth first, with this one first.
    pub fn walk(&self) -> Vec<&Control> {
        let mut out = vec![self];
        for c in &self.children {
            out.extend(c.walk());
        }
        out
    }
}

/// Remove `&` accelerators: `"&Deal"` → `"Deal"`, `"A && B"` → `"A & B"`.
pub fn strip_accel(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '&' {
            if chars.peek() == Some(&'&') {
                chars.next();
                out.push('&');
            }
        } else {
            out.push(c);
        }
    }
    out
}

/// A parsed form: its root control and the `Object =` lines (OCXs used).
#[derive(Debug, Clone)]
pub struct Form {
    pub root: Control,
    pub objects: Vec<String>,
    /// `Attribute VB_Name`.
    pub vb_name: String,
}

impl Form {
    /// All controls with this name, in file order (several for an array).
    pub fn all(&self, name: &str) -> Vec<&Control> {
        self.root
            .walk()
            .into_iter()
            .filter(|c| c.name.eq_ignore_ascii_case(name))
            .collect()
    }

    /// The control `name`, or element `index` of the array `name`.
    pub fn get(&self, name: &str, index: Option<i32>) -> Option<&Control> {
        self.all(name).into_iter().find(|c| c.index() == index)
    }

    /// The font a control is drawn with: its own, else its container's.
    pub fn font_of(&self, target: &Control) -> Font {
        fn find<'a>(
            c: &'a Control,
            target: *const Control,
            inherited: &'a Font,
        ) -> Option<&'a Font> {
            let f = c.font.as_ref().unwrap_or(inherited);
            if std::ptr::eq(c, target) {
                return Some(f);
            }
            c.children.iter().find_map(|k| find(k, target, f))
        }
        let default = Font::default();
        find(&self.root, target, &default)
            .cloned()
            .unwrap_or_default()
    }

    /// Absolute position (twips, from the form's client area) of a control.
    pub fn origin_of(&self, target: &Control) -> (i32, i32) {
        fn find(c: &Control, target: *const Control, x: i32, y: i32) -> Option<(i32, i32)> {
            if std::ptr::eq(c, target) {
                return Some((x, y));
            }
            // Children of a frame or picture box are placed inside it;
            // the form's own Left/Top place the window, not its contents.
            let (bx, by) = if c.kind() == "Form" { (0, 0) } else { (x, y) };
            c.children
                .iter()
                .find_map(|k| find(k, target, bx + k.left(), by + k.top()))
        }
        find(&self.root, target, 0, 0).unwrap_or((0, 0))
    }
}

#[derive(Debug)]
pub struct ParseError {
    pub line: usize,
    pub message: String,
}

impl std::fmt::Display for ParseError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "line {}: {}", self.line, self.message)
    }
}

impl std::error::Error for ParseError {}

/// Parse a `.frm` file's bytes (Windows-1252).
pub fn parse(bytes: &[u8]) -> Result<Form, ParseError> {
    parse_str(&decode_cp1252(bytes))
}

pub fn parse_str(text: &str) -> Result<Form, ParseError> {
    let mut objects = Vec::new();
    let mut stack: Vec<Control> = Vec::new();
    let mut root: Option<Control> = None;
    let mut vb_name = String::new();
    // Inside a BeginProperty block: its name (Font or other) and depth, so
    // nested property blocks (e.g. a grid's column formats) are skipped whole.
    let mut prop_block: Vec<String> = Vec::new();
    let mut font: Option<Font> = None;

    for (i, raw) in text.lines().enumerate() {
        let line = raw.trim();
        let err = |m: &str| ParseError {
            line: i + 1,
            message: m.to_string(),
        };
        if line.is_empty() {
            continue;
        }
        if root.is_some() {
            if let Some(rest) = line.strip_prefix("Attribute VB_Name = ") {
                vb_name = unquote(rest);
                break;
            }
            continue;
        }
        if stack.is_empty() {
            if let Some(rest) = line.strip_prefix("Object = ") {
                objects.push(rest.to_string());
            } else if let Some(rest) = line.strip_prefix("Begin ") {
                stack.push(begin(rest));
            }
            continue;
        }
        if !prop_block.is_empty() {
            if line == "EndProperty" {
                let name = prop_block.pop().unwrap();
                if prop_block.is_empty() && name.eq_ignore_ascii_case("Font") {
                    stack.last_mut().unwrap().font = font.take();
                }
            } else if let Some(rest) = line.strip_prefix("BeginProperty ") {
                prop_block.push(first_word(rest));
            } else if prop_block.len() == 1 && prop_block[0].eq_ignore_ascii_case("Font") {
                if let Some((k, v)) = line.split_once('=') {
                    let f = font.get_or_insert_with(Font::default);
                    let v = value(v.trim());
                    let s = match &v {
                        Value::Str(s) | Value::Raw(s) => s.clone(),
                        Value::Frx { .. } => String::new(),
                    };
                    match k.trim() {
                        "Name" => f.name = s,
                        "Size" => f.size = s.parse().unwrap_or(8.25),
                        "Weight" => f.bold = s.parse::<i32>().unwrap_or(400) >= 600,
                        "Italic" => f.italic = s != "0",
                        "Underline" => f.underline = s != "0",
                        _ => {}
                    }
                }
            }
            continue;
        }
        if let Some(rest) = line.strip_prefix("BeginProperty ") {
            prop_block.push(first_word(rest));
            if prop_block[0].eq_ignore_ascii_case("Font") {
                font = Some(Font::default());
            }
        } else if let Some(rest) = line.strip_prefix("Begin ") {
            stack.push(begin(rest));
        } else if line == "End" {
            let done = stack.pop().ok_or_else(|| err("End without Begin"))?;
            match stack.last_mut() {
                Some(parent) => parent.children.push(done),
                None => root = Some(done),
            }
        } else if let Some((k, v)) = line.split_once('=') {
            stack
                .last_mut()
                .unwrap()
                .props
                .insert(k.trim().to_string(), value(v.trim()));
        } else {
            return Err(err(&format!("cannot read {line:?}")));
        }
    }
    let root = root.ok_or(ParseError {
        line: 0,
        message: "no form found".into(),
    })?;
    Ok(Form {
        root,
        objects,
        vb_name,
    })
}

fn first_word(s: &str) -> String {
    s.split_whitespace().next().unwrap_or("").to_string()
}

fn begin(rest: &str) -> Control {
    let mut it = rest.split_whitespace();
    Control {
        class: it.next().unwrap_or("").to_string(),
        name: it.next().unwrap_or("").to_string(),
        ..Default::default()
    }
}

/// `"text"` → text, with `""` unescaped.
fn unquote(s: &str) -> String {
    let s = s.trim();
    let inner = s.strip_prefix('"').unwrap_or(s);
    let inner = inner.strip_suffix('"').unwrap_or(inner);
    inner.replace("\"\"", "\"")
}

fn value(v: &str) -> Value {
    let (long, v) = match v.strip_prefix('$') {
        Some(rest) => (true, rest),
        None => (false, v),
    };
    if let Some(body) = v.strip_prefix('"') {
        // Find the closing quote, skipping doubled quotes.
        let bytes = body.as_bytes();
        let mut i = 0;
        let mut end = body.len();
        while i < bytes.len() {
            if bytes[i] == b'"' {
                if bytes.get(i + 1) == Some(&b'"') {
                    i += 2;
                    continue;
                }
                end = i;
                break;
            }
            i += 1;
        }
        let s = body[..end].replace("\"\"", "\"");
        let after = &body[(end + 1).min(body.len())..];
        if let Some(off) = after.strip_prefix(':') {
            let off = off.split_whitespace().next().unwrap_or("0");
            if let Ok(offset) = u32::from_str_radix(off, 16) {
                return Value::Frx {
                    file: s,
                    offset,
                    long,
                };
            }
        }
        return Value::Str(s);
    }
    // Unquoted: drop a trailing `'comment`.
    let v = v.split('\'').next().unwrap_or("").trim();
    Value::Raw(v.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE: &str = r#"VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PlayScr
   BackColor       =   &H0000C000&
   Caption         =   "Say ""hi"""
   BeginProperty Font
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Weight          =   700
   EndProperty
   Begin VB.Frame f
      Left            =   100
      Top             =   200
      Visible         =   0   'False
      Begin VB.CommandButton b
         Caption         =   "&Deal && go"
         Index           =   1
         Left            =   10
         Top             =   20
      End
   End
   Begin VB.TextBox t
      Text            =   "x.frx":45E2
   End
End
Attribute VB_Name = "PlayScr"
"#;

    #[test]
    fn reads_tree_and_values() {
        let f = parse_str(SAMPLE).unwrap();
        assert_eq!(f.vb_name, "PlayScr");
        assert_eq!(f.objects.len(), 1);
        assert_eq!(f.root.str("Caption"), Some("Say \"hi\""));
        assert_eq!(f.root.color("BackColor"), Some(0xC000));
        let b = f.get("b", Some(1)).unwrap();
        assert_eq!(b.caption(), "Deal & go");
        assert_eq!(f.origin_of(b), (110, 220));
        assert!(!f.get("f", None).unwrap().visible());
        let font = f.font_of(b);
        assert_eq!(font.size, 9.75);
        assert!(font.bold);
        assert_eq!(
            f.get("t", None).unwrap().props["Text"],
            Value::Frx {
                file: "x.frx".into(),
                offset: 0x45E2,
                long: false
            }
        );
    }
}
