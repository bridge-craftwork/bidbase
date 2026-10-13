//! The classic Windows look: VB's colours, 3D borders, fonts.

use egui::{
    Color32, FontData, FontDefinitions, FontFamily, FontId, Painter, Pos2, Rect, Stroke,
    StrokeKind, Vec2,
};

pub const BUTTON_FACE: Color32 = Color32::from_rgb(0xF0, 0xF0, 0xF0);
pub const HIGHLIGHT: Color32 = Color32::from_rgb(0x00, 0x78, 0xD7);
pub const GRAY_TEXT: Color32 = Color32::from_rgb(0x6D, 0x6D, 0x6D);
const LIGHT: Color32 = Color32::WHITE;
const SHADOW: Color32 = Color32::from_rgb(0xA0, 0xA0, 0xA0);
const DARK: Color32 = Color32::from_rgb(0x69, 0x69, 0x69);

/// A VB colour: `&H00BBGGRR&`, or a system colour `&H800000nn&` as
/// Windows 10 shows it.
pub fn vb_color(v: u32) -> Color32 {
    if v & 0x8000_0000 != 0 {
        return match v & 0xFF {
            0x00 => Color32::from_rgb(0xC8, 0xC8, 0xC8), // scroll bar
            0x01 => Color32::from_rgb(0x00, 0x00, 0x00), // desktop
            0x02 => Color32::from_rgb(0x99, 0xB4, 0xD1), // active title
            0x04 => Color32::from_rgb(0xF0, 0xF0, 0xF0), // menu
            0x05 => Color32::WHITE,                      // window
            0x06 => Color32::from_rgb(0x64, 0x64, 0x64), // window frame
            0x07 | 0x08 | 0x12 => Color32::BLACK,        // menu / window / button text
            0x0D => HIGHLIGHT,
            0x0E => Color32::WHITE, // highlight text
            0x0F => BUTTON_FACE,
            0x10 => SHADOW,
            0x11 => GRAY_TEXT,
            0x14 => LIGHT,
            0x15 => DARK,
            0x18 => Color32::from_rgb(0xFF, 0xFF, 0xE1), // tooltip
            _ => BUTTON_FACE,
        };
    }
    Color32::from_rgb(
        (v & 0xFF) as u8,
        ((v >> 8) & 0xFF) as u8,
        ((v >> 16) & 0xFF) as u8,
    )
}

/// VB's default background for a kind of control.
pub fn default_back(kind: &str) -> u32 {
    match kind {
        "TextBox" | "ComboBox" | "ListBox" | "MSFlexGrid" => 0x8000_0005,
        _ => 0x8000_000F,
    }
}

const SANS: &str = "vb-sans";
const SANS_BOLD: &str = "vb-sans-bold";
const MONO: &str = "vb-mono";
const MONO_BOLD: &str = "vb-mono-bold";

/// Set up the classic look: light visuals and the fonts the forms use
/// (Liberation Sans for MS Sans Serif and Arial, Liberation Mono for
/// Courier; SIL Open Font License, fonts/LICENSE). They are the only fonts:
/// egui's own are left out to keep the web build small, so text is limited
/// to what Liberation covers (Latin, the card suits, box arrows).
pub fn install(ctx: &egui::Context) {
    let mut fonts = FontDefinitions::empty();
    let add = |fonts: &mut FontDefinitions, name: &str, bytes: &'static [u8]| {
        fonts
            .font_data
            .insert(name.into(), FontData::from_static(bytes).into());
        fonts
            .families
            .insert(FontFamily::Name(name.into()), vec![name.to_string()]);
    };
    add(
        &mut fonts,
        SANS,
        include_bytes!("../fonts/LiberationSans-Regular.ttf"),
    );
    add(
        &mut fonts,
        SANS_BOLD,
        include_bytes!("../fonts/LiberationSans-Bold.ttf"),
    );
    add(
        &mut fonts,
        MONO,
        include_bytes!("../fonts/LiberationMono-Regular.ttf"),
    );
    add(
        &mut fonts,
        MONO_BOLD,
        include_bytes!("../fonts/LiberationMono-Bold.ttf"),
    );
    // egui's own widgets (menus) in the same faces.
    fonts
        .families
        .insert(FontFamily::Proportional, vec![SANS.into()]);
    fonts
        .families
        .insert(FontFamily::Monospace, vec![MONO.into()]);
    ctx.set_fonts(fonts);
    // VB forms have one look: light, whatever the system or browser
    // prefers (eframe's web runner follows the browser's dark mode).
    ctx.set_theme(egui::ThemePreference::Light);
    ctx.set_visuals_of(egui::Theme::Light, egui::Visuals::light());
    ctx.set_visuals_of(egui::Theme::Dark, egui::Visuals::light());
}

/// The family for a VB font name.
pub fn family(name: &str, bold: bool) -> FontFamily {
    let mono = name.to_ascii_lowercase().contains("courier");
    FontFamily::Name(
        match (mono, bold) {
            (false, false) => SANS,
            (false, true) => SANS_BOLD,
            (true, false) => MONO,
            (true, true) => MONO_BOLD,
        }
        .into(),
    )
}

/// A VB font (size in points at 96 dpi) as an egui font.
pub fn font_id(name: &str, size_pt: f32, bold: bool) -> FontId {
    FontId::new(size_pt * 96.0 / 72.0, family(name, bold))
}

/// A sunken 3D border, as a text box or picture box has.
pub fn sunken(p: &Painter, r: Rect) {
    let r = r.shrink(0.5);
    p.line_segment([r.left_bottom(), r.left_top()], Stroke::new(1.0, SHADOW));
    p.line_segment([r.left_top(), r.right_top()], Stroke::new(1.0, SHADOW));
    p.line_segment([r.right_top(), r.right_bottom()], Stroke::new(1.0, LIGHT));
    p.line_segment([r.right_bottom(), r.left_bottom()], Stroke::new(1.0, LIGHT));
}

/// A push button.
pub fn button(p: &Painter, r: Rect, face: Color32, pressed: bool, hovered: bool) {
    p.rect_filled(r, 0.0, face);
    let (tl, br) = if pressed {
        (DARK, LIGHT)
    } else {
        (LIGHT, DARK)
    };
    let r1 = r.shrink(0.5);
    p.line_segment([r1.left_bottom(), r1.left_top()], Stroke::new(1.0, tl));
    p.line_segment([r1.left_top(), r1.right_top()], Stroke::new(1.0, tl));
    p.line_segment([r1.right_top(), r1.right_bottom()], Stroke::new(1.0, br));
    p.line_segment([r1.right_bottom(), r1.left_bottom()], Stroke::new(1.0, br));
    if hovered && !pressed {
        p.rect_stroke(
            r.shrink(1.0),
            0.0,
            Stroke::new(1.0, Color32::from_rgb(0x7A, 0xB4, 0xE8)),
            StrokeKind::Inside,
        );
    }
}

pub fn check_box(p: &Painter, b: Rect, on: bool) {
    p.rect_filled(b, 0.0, Color32::WHITE);
    p.rect_stroke(
        b,
        0.0,
        Stroke::new(1.0, Color32::from_rgb(0x33, 0x33, 0x33)),
        StrokeKind::Inside,
    );
    if on {
        let pts = [
            Pos2::new(b.left() + 3.0, b.center().y),
            Pos2::new(b.left() + 5.5, b.bottom() - 3.5),
            Pos2::new(b.right() - 3.0, b.top() + 3.0),
        ];
        p.line_segment([pts[0], pts[1]], Stroke::new(2.0, Color32::BLACK));
        p.line_segment([pts[1], pts[2]], Stroke::new(2.0, Color32::BLACK));
    }
}

pub fn option_button(p: &Painter, b: Rect, on: bool) {
    p.circle_filled(b.center(), 6.0, Color32::WHITE);
    p.circle_stroke(
        b.center(),
        6.0,
        Stroke::new(1.0, Color32::from_rgb(0x33, 0x33, 0x33)),
    );
    if on {
        p.circle_filled(b.center(), 3.0, Color32::BLACK);
    }
}

/// A frame's etched border with its caption breaking the top line.
pub fn frame(p: &Painter, r: Rect, caption: &str, font: FontId, color: Color32, back: Color32) {
    let top = r.top() + font.size / 2.0;
    let b = Rect::from_min_max(Pos2::new(r.left(), top), r.max).shrink(0.5);
    p.rect_stroke(b, 0.0, Stroke::new(1.0, SHADOW), StrokeKind::Middle);
    p.rect_stroke(
        b.translate(Vec2::splat(1.0)),
        0.0,
        Stroke::new(1.0, LIGHT),
        StrokeKind::Middle,
    );
    if !caption.trim().is_empty() {
        let g = p.layout_no_wrap(caption.to_string(), font, color);
        let at = Pos2::new(r.left() + 8.0, r.top());
        p.rect_filled(
            Rect::from_min_size(at - Vec2::new(2.0, 0.0), g.size() + Vec2::new(4.0, 0.0)),
            0.0,
            back,
        );
        p.galley(at, g, color);
    }
}

/// Text wrapped to a rectangle. `align`: 0 left, 1 right, 2 centre (VB's
/// `Alignment`). Buttons centre vertically too.
pub fn wrapped_text(p: &Painter, r: Rect, text: &str, font: FontId, color: Color32, align: i32) {
    if text.is_empty() {
        return;
    }
    let mut job =
        egui::text::LayoutJob::simple(text.replace("\r\n", "\n"), font, color, r.width().max(1.0));
    job.halign = match align {
        1 => egui::Align::RIGHT,
        2 => egui::Align::Center,
        _ => egui::Align::LEFT,
    };
    let g = p.layout_job(job);
    let x = match align {
        1 => r.right(),
        2 => r.center().x,
        _ => r.left(),
    };
    let y = if align == 2 && g.size().y < r.height() {
        r.center().y - g.size().y / 2.0
    } else {
        r.top()
    };
    p.with_clip_rect(r.intersect(p.clip_rect()))
        .galley(Pos2::new(x, y), g, color);
}
