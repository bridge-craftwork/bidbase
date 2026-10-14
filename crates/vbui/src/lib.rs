//! Draw a VB6 form with egui, from Nelson's `.frm`, so the Rust ports of
//! his programs look like his and read like his code.
//!
//! A [`Screen`] holds the parsed form and the run-time state VB code would
//! change on its controls (caption, text, visibility, colours, a check box's
//! value, a list's items, a grid's cells). The port changes that state the
//! way his event handlers do (`screen.set_caption("L_Vul", None, "NS")`),
//! calls [`Screen::show`] each frame, and gets back the [`Event`]s (clicks,
//! menu picks, text changes) to handle.
//!
//! Positions are VB's twips at 15 to a pixel, so a form 9,960 twips wide
//! is 664 points wide. Controls are drawn in VB's order: a container's
//! labels and lines first (VB paints them on the container itself), then
//! its other controls, the first in the file on top.
//!
//! Not drawn here: picture boxes' contents (the port paints them through
//! the `custom` hook of [`Screen::show`]), timers and common dialogs.

pub mod grid;
pub mod ico;
pub mod theme;

use std::collections::HashMap;

use bidbase_forms::{frm, frx, Control, Form};
use egui::{Align2, Color32, FontFamily, FontId, Pos2, Rect, Sense, Stroke, StrokeKind, Ui, Vec2};

pub use grid::Grid;
pub use theme::{install, vb_color};

/// Twips per point (egui's logical pixel).
pub const TWIPS: f32 = 15.0;

/// A control by name (case-insensitive, as in VB) and array index.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Key {
    pub name: String,
    pub index: Option<i32>,
}

impl Key {
    pub fn new(name: &str, index: Option<i32>) -> Self {
        Key {
            name: name.to_ascii_lowercase(),
            index,
        }
    }

    pub fn of(c: &Control) -> Self {
        Key::new(&c.name, c.index())
    }

    /// True for `name` (any element of an array).
    pub fn is(&self, name: &str) -> bool {
        self.name.eq_ignore_ascii_case(name)
    }
}

/// What the port did to a control at run time. `None` keeps the `.frm`'s value.
#[derive(Debug, Clone, Default)]
pub struct State {
    pub caption: Option<String>,
    pub text: Option<String>,
    pub visible: Option<bool>,
    pub enabled: Option<bool>,
    pub back: Option<u32>,
    pub fore: Option<u32>,
    /// Check box (0, 1, 2), option button (0, -1), menu checked (0, 1).
    pub value: Option<i32>,
    pub left: Option<i32>,
    pub top: Option<i32>,
    pub width: Option<i32>,
    pub height: Option<i32>,
    pub list: Option<Vec<String>>,
    /// Selected item of a list or combo box; -1 for none.
    pub list_index: Option<i32>,
    pub font_size: Option<f32>,
    pub grid: Option<Grid>,
    /// Set by [`Screen::bring_to_front`]: drawn above its siblings, the
    /// latest on top.
    pub z_front: Option<u64>,
}

/// Something the user did.
#[derive(Debug, Clone, PartialEq)]
pub enum Event {
    Click(Key),
    DblClick(Key),
    /// Text typed in a text box or combo box.
    Change(Key),
    /// Enter pressed in a text box.
    Enter(Key),
    /// A menu item was picked.
    Menu(Key),
    /// A press on a picture box or label, at (x, y) twips inside it.
    MouseDown {
        key: Key,
        x: f32,
        y: f32,
        right: bool,
    },
    /// A click on a grid cell.
    GridClick {
        key: Key,
        row: usize,
        col: usize,
    },
}

/// A message box waiting for an answer.
#[derive(Debug, Clone)]
struct MsgBox {
    id: String,
    title: String,
    text: String,
    buttons: Vec<String>,
}

/// The form and its run-time state.
pub struct Screen {
    pub form: Form,
    frx: Vec<u8>,
    states: HashMap<Key, State>,
    /// Each control's font (its own or inherited) and container.
    fonts: HashMap<Key, bidbase_forms::Font>,
    parents: HashMap<Key, Key>,
    msgbox: Option<MsgBox>,
    z_counter: u64,
    /// Answers to message boxes, for [`Screen::take_answer`].
    answers: Vec<(String, String)>,
}

impl Screen {
    /// Parse `frm` (and its `.frx`, empty if none) and set up grids and
    /// lists from their design-time properties.
    pub fn new(frm_bytes: &[u8], frx_bytes: Vec<u8>) -> Result<Self, frm::ParseError> {
        let form = frm::parse(frm_bytes)?;
        let mut s = Screen {
            form,
            frx: frx_bytes,
            states: HashMap::new(),
            fonts: HashMap::new(),
            parents: HashMap::new(),
            msgbox: None,
            z_counter: 0,
            answers: Vec::new(),
        };
        fn index(
            c: &Control,
            inherited: &bidbase_forms::Font,
            fonts: &mut HashMap<Key, bidbase_forms::Font>,
            parents: &mut HashMap<Key, Key>,
        ) {
            let f = c.font.clone().unwrap_or_else(|| inherited.clone());
            for k in &c.children {
                parents.insert(Key::of(k), Key::of(c));
                index(k, &f, fonts, parents);
            }
            fonts.insert(Key::of(c), f);
        }
        index(
            &s.form.root,
            &Default::default(),
            &mut s.fonts,
            &mut s.parents,
        );
        let mut init = Vec::new();
        for c in s.form.root.walk() {
            let k = Key::of(c);
            if c.kind() == "MSFlexGrid" {
                init.push((k, None, Some(Grid::from_control(c))));
            } else if let Some(frm::Value::Frx { offset, .. }) = c.props.get("List") {
                init.push((k, frx::list(&s.frx, *offset), None));
            }
        }
        for (k, list, grid) in init {
            let st = s.states.entry(k).or_default();
            st.list = list.or(st.list.take());
            st.grid = grid;
        }
        Ok(s)
    }

    /// The form's client size in points.
    pub fn size(&self) -> Vec2 {
        let r = &self.form.root;
        Vec2::new(
            r.int("ClientWidth").unwrap_or(0) as f32 / TWIPS,
            r.int("ClientHeight").unwrap_or(0) as f32 / TWIPS,
        )
    }

    /// The window title (the form's caption).
    pub fn title(&self) -> String {
        self.form.root.str("Caption").unwrap_or("").to_string()
    }

    /// The form's icon (an `.ico` file), if it has one.
    pub fn icon(&self) -> Option<&[u8]> {
        match self.form.root.props.get("Icon")? {
            frm::Value::Frx { offset, .. } => frx::picture(&self.frx, *offset),
            _ => None,
        }
    }

    fn control(&self, k: &Key) -> Option<&Control> {
        self.form.get(&k.name, k.index)
    }

    /// True if the control `k` is `container` or sits inside it.
    pub fn within(&self, k: &Key, container: &str) -> bool {
        let mut at = Some(k);
        while let Some(c) = at {
            if c.is(container) {
                return true;
            }
            at = self.parents.get(c);
        }
        false
    }

    /// The run-time state of a control, created on first use.
    pub fn st(&mut self, name: &str, index: Option<i32>) -> &mut State {
        self.states.entry(Key::new(name, index)).or_default()
    }

    fn state(&self, k: &Key) -> Option<&State> {
        self.states.get(k)
    }

    fn design_text(&self, c: &Control, prop: &str) -> String {
        match c.props.get(prop) {
            Some(frm::Value::Str(s)) | Some(frm::Value::Raw(s)) => s.clone(),
            Some(frm::Value::Frx { offset, long, .. }) => {
                frx::text(&self.frx, *offset, *long).unwrap_or_default()
            }
            None => String::new(),
        }
    }

    // --- VB-style properties -------------------------------------------

    pub fn caption(&self, name: &str, index: Option<i32>) -> String {
        let k = Key::new(name, index);
        if let Some(c) = self.state(&k).and_then(|s| s.caption.clone()) {
            return c;
        }
        self.control(&k)
            .map(|c| self.design_text(c, "Caption"))
            .unwrap_or_default()
    }

    pub fn set_caption(&mut self, name: &str, index: Option<i32>, text: impl Into<String>) {
        self.st(name, index).caption = Some(text.into());
    }

    pub fn text(&self, name: &str, index: Option<i32>) -> String {
        let k = Key::new(name, index);
        if let Some(t) = self.state(&k).and_then(|s| s.text.clone()) {
            return t;
        }
        self.control(&k)
            .map(|c| self.design_text(c, "Text"))
            .unwrap_or_default()
    }

    pub fn set_text(&mut self, name: &str, index: Option<i32>, text: impl Into<String>) {
        self.st(name, index).text = Some(text.into());
    }

    pub fn visible(&self, name: &str, index: Option<i32>) -> bool {
        let k = Key::new(name, index);
        self.state(&k)
            .and_then(|s| s.visible)
            .or_else(|| self.control(&k).map(|c| c.visible()))
            .unwrap_or(false)
    }

    pub fn set_visible(&mut self, name: &str, index: Option<i32>, on: bool) {
        self.st(name, index).visible = Some(on);
    }

    pub fn enabled(&self, name: &str, index: Option<i32>) -> bool {
        let k = Key::new(name, index);
        self.state(&k)
            .and_then(|s| s.enabled)
            .or_else(|| self.control(&k).map(|c| c.flag("Enabled", true)))
            .unwrap_or(true)
    }

    pub fn set_enabled(&mut self, name: &str, index: Option<i32>, on: bool) {
        self.st(name, index).enabled = Some(on);
    }

    /// A check box's value (0 unchecked, 1 checked), an option button's
    /// (-1 when chosen) or a menu item's `Checked` (1).
    pub fn value(&self, name: &str, index: Option<i32>) -> i32 {
        let k = Key::new(name, index);
        self.state(&k)
            .and_then(|s| s.value)
            .or_else(|| {
                self.control(&k)
                    .and_then(|c| c.int("Value").or_else(|| c.int("Checked")))
            })
            .unwrap_or(0)
    }

    pub fn set_value(&mut self, name: &str, index: Option<i32>, v: i32) {
        self.st(name, index).value = Some(v);
    }

    pub fn set_back(&mut self, name: &str, index: Option<i32>, vb: u32) {
        self.st(name, index).back = Some(vb);
    }

    pub fn set_fore(&mut self, name: &str, index: Option<i32>, vb: u32) {
        self.st(name, index).fore = Some(vb);
    }

    /// Items of a list or combo box.
    pub fn list(&mut self, name: &str, index: Option<i32>) -> &mut Vec<String> {
        self.st(name, index).list.get_or_insert_with(Vec::new)
    }

    pub fn list_index(&self, name: &str, index: Option<i32>) -> i32 {
        self.state(&Key::new(name, index))
            .and_then(|s| s.list_index)
            .unwrap_or(-1)
    }

    pub fn set_list_index(&mut self, name: &str, index: Option<i32>, i: i32) {
        self.st(name, index).list_index = Some(i);
    }

    /// A grid's cells (every MSFlexGrid has one from the start).
    pub fn grid(&mut self, name: &str) -> &mut Grid {
        self.st(name, None).grid.get_or_insert_with(Grid::default)
    }

    pub fn grid_ref(&self, name: &str) -> Option<&Grid> {
        self.state(&Key::new(name, None))?.grid.as_ref()
    }

    /// VB's `ZOrder 0`: draw the control above its siblings.
    pub fn bring_to_front(&mut self, name: &str, index: Option<i32>) {
        self.z_counter += 1;
        let z = self.z_counter;
        self.st(name, index).z_front = Some(z);
    }

    pub fn move_to(&mut self, name: &str, index: Option<i32>, left: i32, top: i32) {
        let s = self.st(name, index);
        s.left = Some(left);
        s.top = Some(top);
    }

    pub fn resize(&mut self, name: &str, index: Option<i32>, width: i32, height: i32) {
        let s = self.st(name, index);
        s.width = Some(width);
        s.height = Some(height);
    }

    /// Where a control is now, in twips from its container: (left, top,
    /// width, height).
    pub fn bounds(&self, name: &str, index: Option<i32>) -> (i32, i32, i32, i32) {
        let k = Key::new(name, index);
        let Some(c) = self.control(&k) else {
            return (0, 0, 0, 0);
        };
        let s = self.state(&k);
        (
            s.and_then(|s| s.left).unwrap_or(c.left()),
            s.and_then(|s| s.top).unwrap_or(c.top()),
            s.and_then(|s| s.width).unwrap_or(c.width()),
            s.and_then(|s| s.height).unwrap_or(c.height()),
        )
    }

    // --- message boxes ----------------------------------------------------

    /// Show a message box; the answer comes back through
    /// [`Screen::take_answer`] with this `id`. One at a time: a second call
    /// replaces the first.
    pub fn msgbox(&mut self, id: &str, title: &str, text: &str, buttons: &[&str]) {
        self.msgbox = Some(MsgBox {
            id: id.into(),
            title: title.into(),
            text: text.into(),
            buttons: buttons.iter().map(|b| b.to_string()).collect(),
        });
    }

    pub fn msgbox_open(&self) -> bool {
        self.msgbox.is_some()
    }

    /// The button pressed in the message box `id`, once.
    pub fn take_answer(&mut self, id: &str) -> Option<String> {
        let i = self.answers.iter().position(|(a, _)| a == id)?;
        Some(self.answers.remove(i).1)
    }

    /// All answers given since the last call: (id, button).
    pub fn take_answers(&mut self) -> Vec<(String, String)> {
        std::mem::take(&mut self.answers)
    }

    // --- drawing -----------------------------------------------------------

    /// Draw the menu bar from the form's top-level menus.
    pub fn menu_bar(&mut self, ui: &mut Ui) -> Vec<Event> {
        let mut events = Vec::new();
        let menus: Vec<Control> = self
            .form
            .root
            .children
            .iter()
            .filter(|c| c.kind() == "Menu")
            .cloned()
            .collect();
        // The menu bar's own background (Windows' menu colour).
        ui.painter()
            .rect_filled(ui.max_rect(), 0.0, vb_color(0x8000_0004));
        egui::MenuBar::new().ui(ui, |ui| {
            for m in &menus {
                self.menu(ui, m, &mut events);
            }
        });
        events
    }

    fn menu(&self, ui: &mut Ui, m: &Control, events: &mut Vec<Event>) {
        let k = Key::of(m);
        if !self.visible(&m.name, m.index()) {
            return;
        }
        let caption = frm::strip_accel(&self.caption(&m.name, m.index()));
        let enabled = self.enabled(&m.name, m.index());
        if m.children.iter().any(|c| c.kind() == "Menu") {
            ui.add_enabled_ui(enabled, |ui| {
                ui.menu_button(caption, |ui| {
                    for c in &m.children {
                        self.menu(ui, c, events);
                    }
                });
            });
        } else if caption == "-" {
            ui.separator();
        } else {
            let checked = self.value(&m.name, m.index()) != 0;
            let label = if checked {
                format!("√ {caption}")
            } else {
                format!("    {caption}")
            };
            if ui.add_enabled(enabled, egui::Button::new(label)).clicked() {
                events.push(Event::Menu(k));
                ui.close();
            }
        }
    }

    /// Draw the form at `origin` and return what the user did. `custom` is
    /// called for every picture box, with its rectangle, before its child
    /// controls are drawn, so the port can paint its contents.
    pub fn show(
        &mut self,
        ui: &mut Ui,
        origin: Pos2,
        custom: &mut dyn FnMut(&Ui, &Key, Rect),
    ) -> Vec<Event> {
        let mut events = Vec::new();
        let root = self.form.root.clone();
        let rect = Rect::from_min_size(origin, self.size());
        let back = root.color("BackColor").unwrap_or(0x8000_000F);
        ui.painter().rect_filled(rect, 0.0, vb_color(back));
        let modal = self.msgbox.is_some();
        ui.add_enabled_ui(!modal, |ui| {
            self.children(ui, &root, origin, rect, custom, &mut events);
        });
        self.show_msgbox(ui.ctx());
        events
    }

    fn children(
        &mut self,
        ui: &mut Ui,
        parent: &Control,
        origin: Pos2,
        clip: Rect,
        custom: &mut dyn FnMut(&Ui, &Key, Rect),
        events: &mut Vec<Event>,
    ) {
        let graphical = |c: &Control| matches!(c.kind(), "Label" | "Line" | "Shape" | "Image");
        let front = |c: &Control| {
            self.states
                .get(&Key::of(c))
                .and_then(|s| s.z_front)
                .unwrap_or(0)
        };
        let mut windowed: Vec<&Control> = parent
            .children
            .iter()
            .rev()
            .filter(|c| !graphical(c))
            .collect();
        // Stable: controls brought to the front go last (on top), in the
        // order they were brought there.
        windowed.sort_by_key(|c| front(c));
        let order: Vec<&Control> = parent
            .children
            .iter()
            .rev()
            .filter(|c| graphical(c))
            .chain(windowed)
            .collect();
        for c in order {
            self.control_ui(ui, c, origin, clip, custom, events);
        }
    }

    fn control_ui(
        &mut self,
        ui: &mut Ui,
        c: &Control,
        origin: Pos2,
        clip: Rect,
        custom: &mut dyn FnMut(&Ui, &Key, Rect),
        events: &mut Vec<Event>,
    ) {
        let k = Key::of(c);
        let kind = c.kind();
        if matches!(kind, "Menu" | "Timer" | "CommonDialog" | "Adodc" | "Inet")
            || !self.visible(&c.name, c.index())
        {
            return;
        }
        let (l, t, w, h) = self.bounds(&c.name, c.index());
        let rect = Rect::from_min_size(
            origin + Vec2::new(l as f32, t as f32) / TWIPS,
            Vec2::new(w as f32, h as f32) / TWIPS,
        );
        if kind == "Line" {
            let p = |x: &str, y: &str| {
                origin
                    + Vec2::new(c.int(x).unwrap_or(0) as f32, c.int(y).unwrap_or(0) as f32) / TWIPS
            };
            let color = vb_color(c.color("BorderColor").unwrap_or(0x8000_0008));
            let width = c.int("BorderWidth").unwrap_or(1) as f32;
            ui.painter()
                .with_clip_rect(clip)
                .line_segment([p("X1", "Y1"), p("X2", "Y2")], Stroke::new(width, color));
            return;
        }
        let font = self.fonts.get(&k).cloned().unwrap_or_default();
        let size = self
            .state(&k)
            .and_then(|s| s.font_size)
            .unwrap_or(font.size);
        let font_id = theme::font_id(&font.name, size, font.bold);
        let back = self
            .state(&k)
            .and_then(|s| s.back)
            .or_else(|| c.color("BackColor"))
            .unwrap_or(theme::default_back(kind));
        let fore = self
            .state(&k)
            .and_then(|s| s.fore)
            .or_else(|| c.color("ForeColor"))
            .unwrap_or(0x8000_0012);
        let enabled = self.enabled(&c.name, c.index());
        let id = ui.id().with(("vb", &k.name, k.index));
        let painter = ui.painter().with_clip_rect(clip);
        let visible_rect = rect.intersect(clip);
        match kind {
            "Label" => {
                if c.int("BackStyle").unwrap_or(1) != 0 {
                    painter.rect_filled(rect, 0.0, vb_color(back));
                }
                if c.int("BorderStyle").unwrap_or(0) == 1 {
                    theme::sunken(&painter, rect);
                }
                let text = frm::strip_accel(&self.caption(&c.name, c.index()));
                let align = c.int("Alignment").unwrap_or(0);
                theme::wrapped_text(&painter, rect, &text, font_id, vb_color(fore), align);
                let r = ui.interact(visible_rect, id, Sense::click());
                click_events(&r, &k, rect, events);
            }
            "CommandButton" => {
                let graphical = c.int("Style").unwrap_or(0) == 1;
                let r = ui.interact(visible_rect, id, Sense::click());
                let face = if graphical {
                    vb_color(back)
                } else {
                    theme::BUTTON_FACE
                };
                theme::button(
                    &painter,
                    rect,
                    face,
                    r.is_pointer_button_down_on(),
                    r.hovered(),
                );
                let text = frm::strip_accel(&self.caption(&c.name, c.index()));
                let color = if enabled {
                    vb_color(fore)
                } else {
                    theme::GRAY_TEXT
                };
                theme::wrapped_text(&painter, rect.shrink(2.0), &text, font_id, color, 2);
                if enabled && r.clicked() {
                    events.push(Event::Click(k));
                }
            }
            "CheckBox" | "OptionButton" => {
                painter.rect_filled(rect, 0.0, vb_color(back));
                let r = ui.interact(visible_rect, id, Sense::click());
                let on = self.value(&c.name, c.index()) != 0;
                let bx = Rect::from_min_size(
                    Pos2::new(rect.left() + 1.0, rect.center().y - 6.5),
                    Vec2::splat(13.0),
                );
                if kind == "CheckBox" {
                    theme::check_box(&painter, bx, on);
                } else {
                    theme::option_button(&painter, bx, on);
                }
                let text_rect =
                    Rect::from_min_max(Pos2::new(bx.right() + 4.0, rect.top()), rect.max);
                let text = frm::strip_accel(&self.caption(&c.name, c.index()));
                theme::wrapped_text(&painter, text_rect, &text, font_id, vb_color(fore), 0);
                if enabled && r.clicked() {
                    if kind == "CheckBox" {
                        self.set_value(&c.name, c.index(), if on { 0 } else { 1 });
                    } else {
                        self.choose_option(c);
                    }
                    events.push(Event::Click(k));
                }
            }
            "Frame" => {
                // A frame over other controls takes the clicks that land on it.
                ui.interact(visible_rect, id, Sense::click());
                painter.rect_filled(rect, 0.0, vb_color(back));
                if c.int("BorderStyle").unwrap_or(1) != 0 {
                    let caption = frm::strip_accel(&self.caption(&c.name, c.index()));
                    theme::frame(
                        &painter,
                        rect,
                        &caption,
                        font_id,
                        vb_color(fore),
                        vb_color(back),
                    );
                }
                let inner = rect.intersect(clip);
                self.children(ui, c, rect.min, inner, custom, events);
            }
            "PictureBox" => {
                painter.rect_filled(rect, 0.0, vb_color(back));
                if c.int("BorderStyle").unwrap_or(1) != 0 {
                    if c.int("Appearance").unwrap_or(1) == 0 {
                        painter.rect_stroke(
                            rect,
                            0.0,
                            Stroke::new(1.0, Color32::BLACK),
                            StrokeKind::Inside,
                        );
                    } else {
                        theme::sunken(&painter, rect);
                    }
                }
                let r = ui.interact(visible_rect, id, Sense::click());
                click_events(&r, &k, rect, events);
                let mut sub = ui.new_child(egui::UiBuilder::new().max_rect(visible_rect));
                sub.set_clip_rect(visible_rect);
                custom(&sub, &k, rect);
                let inner = rect.intersect(clip);
                self.children(ui, c, rect.min, inner, custom, events);
            }
            "TextBox" => self.text_box(ui, c, &k, rect, visible_rect, font_id, back, fore, events),
            "ComboBox" | "ListBox" => {
                self.list_box(ui, c, &k, rect, visible_rect, font_id, back, fore, events)
            }
            "MSFlexGrid" => {
                let grid = self
                    .states
                    .entry(k.clone())
                    .or_default()
                    .grid
                    .get_or_insert_with(Grid::default);
                let r = ui.interact(visible_rect, id, Sense::click());
                grid.paint(&painter, rect, &font_id);
                if r.clicked() {
                    if let Some(p) = r.interact_pointer_pos() {
                        if let Some((row, col)) = grid.cell_at(rect, p) {
                            events.push(Event::GridClick { key: k, row, col });
                        }
                    }
                }
            }
            "DataGrid" | "HScrollBar" | "VScrollBar" => {
                painter.rect_filled(rect, 0.0, vb_color(back));
                theme::sunken(&painter, rect);
            }
            _ => {}
        }
    }

    fn choose_option(&mut self, c: &Control) {
        // Option buttons in the same container are a group.
        let parent = self.parents.get(&Key::of(c)).cloned();
        let group: Vec<(String, Option<i32>)> = self
            .form
            .root
            .walk()
            .into_iter()
            .filter(|o| {
                o.kind() == "OptionButton" && self.parents.get(&Key::of(o)) == parent.as_ref()
            })
            .map(|o| (o.name.clone(), o.index()))
            .collect();
        for (n, i) in group {
            self.set_value(&n, i, 0);
        }
        self.set_value(&c.name, c.index(), -1);
    }

    #[allow(clippy::too_many_arguments)]
    fn text_box(
        &mut self,
        ui: &mut Ui,
        c: &Control,
        k: &Key,
        rect: Rect,
        visible_rect: Rect,
        font_id: FontId,
        back: u32,
        fore: u32,
        events: &mut Vec<Event>,
    ) {
        let mut text = self.text(&c.name, c.index());
        let multiline = c.flag("MultiLine", false);
        let painter = ui.painter().with_clip_rect(visible_rect);
        painter.rect_filled(rect, 0.0, vb_color(back));
        if c.int("BorderStyle").unwrap_or(1) != 0 {
            theme::sunken(&painter, rect);
        }
        let align = match c.int("Alignment").unwrap_or(0) {
            1 => egui::Align::RIGHT,
            2 => egui::Align::Center,
            _ => egui::Align::LEFT,
        };
        let inner = rect.shrink(2.0);
        let mut child =
            ui.new_child(egui::UiBuilder::new().max_rect(inner.intersect(visible_rect)));
        child.set_clip_rect(inner.intersect(visible_rect));
        let resp = if multiline {
            egui::ScrollArea::vertical()
                .id_salt(("vbscroll", &k.name, k.index))
                .max_height(inner.height())
                .auto_shrink(false)
                .stick_to_bottom(false)
                .show(&mut child, |ui| {
                    ui.add(
                        egui::TextEdit::multiline(&mut text)
                            .id_salt(("vbtext", &k.name, k.index))
                            .font(font_id.clone())
                            .text_color(vb_color(fore))
                            .frame(egui::Frame::NONE)
                            .horizontal_align(align)
                            .desired_width(inner.width())
                            .desired_rows(1)
                            .margin(Vec2::ZERO),
                    )
                })
                .inner
        } else {
            child.add(
                egui::TextEdit::singleline(&mut text)
                    .id_salt(("vbtext", &k.name, k.index))
                    .font(font_id)
                    .text_color(vb_color(fore))
                    .frame(egui::Frame::NONE)
                    .horizontal_align(align)
                    .desired_width(inner.width())
                    .margin(Vec2::ZERO),
            )
        };
        if resp.changed() {
            self.set_text(&c.name, c.index(), text);
            events.push(Event::Change(k.clone()));
        }
        if resp.lost_focus() && ui.input(|i| i.key_pressed(egui::Key::Enter)) {
            events.push(Event::Enter(k.clone()));
        }
        if resp.clicked() {
            events.push(Event::Click(k.clone()));
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn list_box(
        &mut self,
        ui: &mut Ui,
        c: &Control,
        k: &Key,
        rect: Rect,
        visible_rect: Rect,
        font_id: FontId,
        back: u32,
        fore: u32,
        events: &mut Vec<Event>,
    ) {
        // Combo style 0 (drop-down combo) and 2 (drop-down list) show one
        // line and a drop-down; style 1 (simple combo) and list boxes show
        // the list in place.
        let style = c.int("Style").unwrap_or(0);
        let dropdown = c.kind() == "ComboBox" && style != 1;
        let items = self
            .states
            .get(k)
            .and_then(|s| s.list.clone())
            .unwrap_or_default();
        let selected = self.list_index(&c.name, c.index());
        let painter = ui.painter().with_clip_rect(visible_rect);
        // Rows as tight as the font allows, so a list shows as many lines as
        // Nelson sized it for (the DDA table's five in 90 pixels).
        let row_h = font_id.size;
        if dropdown {
            let rect = Rect::from_min_size(rect.min, Vec2::new(rect.width(), row_h + 6.0));
            painter.rect_filled(rect, 0.0, vb_color(back));
            theme::sunken(&painter, rect);
            let text = if c.kind() == "ComboBox" && style == 0 {
                self.text(&c.name, c.index())
            } else {
                usize::try_from(selected)
                    .ok()
                    .and_then(|i| items.get(i).cloned())
                    .unwrap_or_default()
            };
            let arrow = Rect::from_min_max(
                Pos2::new(rect.right() - 17.0, rect.top() + 2.0),
                rect.max - Vec2::splat(2.0),
            );
            theme::button(&painter, arrow, theme::BUTTON_FACE, false, false);
            painter.text(
                arrow.center(),
                Align2::CENTER_CENTER,
                "▼",
                FontId::proportional(8.0),
                Color32::BLACK,
            );
            painter.text(
                Pos2::new(rect.left() + 3.0, rect.center().y),
                Align2::LEFT_CENTER,
                text,
                font_id.clone(),
                vb_color(fore),
            );
            let id = ui.id().with(("vbcombo", &k.name, k.index));
            let r = ui.interact(rect.intersect(visible_rect), id, Sense::click());
            let popup = egui::Popup::menu(&r).width(rect.width());
            popup.show(|ui| {
                egui::ScrollArea::vertical()
                    .max_height(300.0)
                    .show(ui, |ui| {
                        for (i, item) in items.iter().enumerate() {
                            if ui
                                .selectable_label(
                                    i as i32 == selected,
                                    egui::RichText::new(item).font(font_id.clone()),
                                )
                                .clicked()
                            {
                                self.set_list_index(&c.name, c.index(), i as i32);
                                self.set_text(&c.name, c.index(), item.clone());
                                events.push(Event::Click(k.clone()));
                                ui.close();
                            }
                        }
                    });
            });
            return;
        }
        // A simple combo has its text line above the list.
        let rect = if c.kind() == "ComboBox" {
            let line = Rect::from_min_size(rect.min, Vec2::new(rect.width(), row_h + 4.0));
            painter.rect_filled(line, 0.0, vb_color(back));
            theme::sunken(&painter, line);
            painter.with_clip_rect(line.shrink(2.0)).text(
                Pos2::new(line.left() + 3.0, line.center().y),
                Align2::LEFT_CENTER,
                self.text(&c.name, c.index()),
                font_id.clone(),
                vb_color(fore),
            );
            Rect::from_min_max(Pos2::new(rect.left(), line.bottom() + 1.0), rect.max)
        } else {
            rect
        };
        painter.rect_filled(rect, 0.0, vb_color(back));
        theme::sunken(&painter, rect);
        let inner = rect.shrink(2.0);
        let mut child =
            ui.new_child(egui::UiBuilder::new().max_rect(inner.intersect(visible_rect)));
        child.set_clip_rect(inner.intersect(visible_rect));
        egui::ScrollArea::vertical()
            .id_salt(("vblist", &k.name, k.index))
            .max_height(inner.height())
            .auto_shrink(false)
            .show(&mut child, |ui| {
                ui.spacing_mut().item_spacing.y = 0.0;
                for (i, item) in items.iter().enumerate() {
                    let (r, resp) =
                        ui.allocate_exact_size(Vec2::new(inner.width(), row_h), Sense::click());
                    let sel = i as i32 == selected;
                    if sel {
                        ui.painter().rect_filled(r, 0.0, theme::HIGHLIGHT);
                    }
                    ui.painter().text(
                        Pos2::new(r.left() + 2.0, r.center().y),
                        Align2::LEFT_CENTER,
                        item,
                        font_id.clone(),
                        if sel { Color32::WHITE } else { vb_color(fore) },
                    );
                    if resp.clicked() {
                        self.set_list_index(&c.name, c.index(), i as i32);
                        events.push(Event::Click(k.clone()));
                    }
                    if resp.double_clicked() {
                        events.push(Event::DblClick(k.clone()));
                    }
                }
            });
    }

    fn show_msgbox(&mut self, ctx: &egui::Context) {
        let Some(mb) = self.msgbox.clone() else {
            return;
        };
        let mut answer = None;
        egui::Modal::new(egui::Id::new("vb-msgbox")).show(ctx, |ui| {
            ui.set_max_width(420.0);
            ui.strong(&mb.title);
            ui.add_space(6.0);
            ui.label(&mb.text);
            ui.add_space(8.0);
            ui.horizontal(|ui| {
                for b in &mb.buttons {
                    if ui.button(b).clicked() {
                        answer = Some(b.clone());
                    }
                }
            });
        });
        if let Some(a) = answer {
            self.answers.push((mb.id, a));
            self.msgbox = None;
        }
    }
}

fn click_events(r: &egui::Response, k: &Key, rect: Rect, events: &mut Vec<Event>) {
    let at = |p: Pos2| ((p.x - rect.left()) * TWIPS, (p.y - rect.top()) * TWIPS);
    if r.clicked() || r.secondary_clicked() {
        if let Some(p) = r.interact_pointer_pos() {
            let (x, y) = at(p);
            events.push(Event::MouseDown {
                key: k.clone(),
                x,
                y,
                right: r.secondary_clicked(),
            });
        }
        if r.clicked() {
            events.push(Event::Click(k.clone()));
        }
    }
    if r.double_clicked() {
        events.push(Event::DblClick(k.clone()));
    }
}

/// The font a VB font name maps to.
pub fn family_for(name: &str) -> FontFamily {
    theme::family(name, false)
}
