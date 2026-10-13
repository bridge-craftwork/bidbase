//! An MSFlexGrid: rows and columns of text cells with per-cell colours.
//! Sizes in twips, as VB code sets them (`ColWidth`, `RowHeight`).

use bidbase_forms::Control;
use egui::{Align2, Color32, FontId, Painter, Pos2, Rect, Stroke, Vec2};

use crate::theme::vb_color;
use crate::TWIPS;

#[derive(Debug, Clone, Default)]
pub struct Cell {
    pub text: String,
    /// VB colour; `None` for the grid's.
    pub back: Option<u32>,
    pub fore: Option<u32>,
    pub bold: Option<bool>,
    /// MSFlexGrid's `CellAlignment`: 1 left-centre, 4 centre-centre, 7 right-centre; 0 default.
    pub align: i32,
}

#[derive(Debug, Clone)]
pub struct Grid {
    pub rows: usize,
    pub cols: usize,
    pub fixed_rows: usize,
    pub fixed_cols: usize,
    pub col_width: Vec<i32>,
    pub row_height: Vec<i32>,
    pub back: u32,
    pub back_fixed: u32,
    pub fore: u32,
    pub grid_color: u32,
    pub grid_color_fixed: u32,
    /// 0 none, 1 flat, 2 inset, 3 raised.
    pub grid_lines: i32,
    pub cells: Vec<Vec<Cell>>,
    pub bold: bool,
}

impl Default for Grid {
    fn default() -> Self {
        Grid::new(2, 2)
    }
}

impl Grid {
    pub fn new(rows: usize, cols: usize) -> Self {
        Grid {
            rows,
            cols,
            fixed_rows: 1,
            fixed_cols: 1,
            col_width: vec![960; cols],
            row_height: vec![240; rows],
            back: 0x8000_0005,
            back_fixed: 0x8000_000F,
            fore: 0x8000_0008,
            grid_color: 0xC0C0C0,
            grid_color_fixed: 0,
            grid_lines: 1,
            cells: vec![vec![Cell::default(); cols]; rows],
            bold: false,
        }
    }

    /// A grid as the `.frm` sets it up.
    pub fn from_control(c: &Control) -> Self {
        let rows = c.int("Rows").unwrap_or(2).max(0) as usize;
        let cols = c.int("Cols").unwrap_or(2).max(0) as usize;
        let mut g = Grid::new(rows, cols);
        g.fixed_rows = c.int("FixedRows").unwrap_or(1).max(0) as usize;
        g.fixed_cols = c.int("FixedCols").unwrap_or(1).max(0) as usize;
        if let Some(v) = c.color("BackColor") {
            g.back = v;
        }
        if let Some(v) = c.color("BackColorFixed") {
            g.back_fixed = v;
        }
        if let Some(v) = c.color("ForeColor") {
            g.fore = v;
        }
        if let Some(v) = c.color("GridColor") {
            g.grid_color = v;
        }
        if let Some(v) = c.color("GridColorFixed") {
            g.grid_color_fixed = v;
        }
        g.grid_lines = c.int("GridLines").unwrap_or(1);
        g.bold = c.font.as_ref().is_some_and(|f| f.bold);
        g
    }

    pub fn set_rows(&mut self, rows: usize) {
        self.rows = rows;
        self.row_height
            .resize(rows, *self.row_height.last().unwrap_or(&240));
        self.cells.resize(rows, vec![Cell::default(); self.cols]);
    }

    pub fn set_cols(&mut self, cols: usize) {
        self.cols = cols;
        self.col_width
            .resize(cols, *self.col_width.last().unwrap_or(&960));
        for r in &mut self.cells {
            r.resize(cols, Cell::default());
        }
    }

    pub fn text(&self, row: usize, col: usize) -> &str {
        self.cells
            .get(row)
            .and_then(|r| r.get(col))
            .map(|c| c.text.as_str())
            .unwrap_or("")
    }

    /// `TextMatrix(row, col) = text`.
    pub fn set_text(&mut self, row: usize, col: usize, text: impl Into<String>) {
        if let Some(c) = self.cells.get_mut(row).and_then(|r| r.get_mut(col)) {
            c.text = text.into();
        }
    }

    pub fn cell_mut(&mut self, row: usize, col: usize) -> Option<&mut Cell> {
        self.cells.get_mut(row).and_then(|r| r.get_mut(col))
    }

    /// Clear every cell's text and colours.
    pub fn clear(&mut self) {
        for r in &mut self.cells {
            for c in r {
                *c = Cell::default();
            }
        }
    }

    fn col_x(&self, col: usize) -> f32 {
        self.col_width[..col]
            .iter()
            .map(|&w| w as f32 / TWIPS)
            .sum()
    }

    fn row_y(&self, row: usize) -> f32 {
        self.row_height[..row]
            .iter()
            .map(|&h| h as f32 / TWIPS)
            .sum()
    }

    /// The cell under `p`, if any.
    pub fn cell_at(&self, rect: Rect, p: Pos2) -> Option<(usize, usize)> {
        let v = p - rect.min;
        let row = (0..self.rows).find(|&r| v.y >= self.row_y(r) && v.y < self.row_y(r + 1))?;
        let col = (0..self.cols).find(|&c| v.x >= self.col_x(c) && v.x < self.col_x(c + 1))?;
        Some((row, col))
    }

    pub fn paint(&self, p: &Painter, rect: Rect, font: &FontId) {
        let p = p.with_clip_rect(rect.intersect(p.clip_rect()));
        // The area past the last row/column is the control's dark background.
        p.rect_filled(rect, 0.0, Color32::from_rgb(0x80, 0x80, 0x80));
        let bold_font = FontId::new(font.size, crate::theme::family("", true));
        for r in 0..self.rows {
            for c in 0..self.cols {
                let cell = &self.cells[r][c];
                let fixed = r < self.fixed_rows || c < self.fixed_cols;
                let min = rect.min + Vec2::new(self.col_x(c), self.row_y(r));
                let size = Vec2::new(self.col_width[c] as f32, self.row_height[r] as f32) / TWIPS;
                let cr = Rect::from_min_size(min, size);
                let back = cell
                    .back
                    .unwrap_or(if fixed { self.back_fixed } else { self.back });
                p.rect_filled(cr, 0.0, vb_color(back));
                if self.grid_lines != 0 {
                    let gc = if fixed {
                        self.grid_color_fixed
                    } else {
                        self.grid_color
                    };
                    let line = if self.grid_lines == 2 && !fixed {
                        Color32::from_rgb(0x80, 0x80, 0x80)
                    } else {
                        vb_color(gc)
                    };
                    p.line_segment([cr.right_top(), cr.right_bottom()], Stroke::new(1.0, line));
                    p.line_segment(
                        [cr.left_bottom(), cr.right_bottom()],
                        Stroke::new(1.0, line),
                    );
                    if self.grid_lines == 2 {
                        p.line_segment(
                            [cr.left_top(), cr.right_top()],
                            Stroke::new(1.0, Color32::WHITE),
                        );
                        p.line_segment(
                            [cr.left_top(), cr.left_bottom()],
                            Stroke::new(1.0, Color32::WHITE),
                        );
                    }
                    if fixed {
                        p.line_segment(
                            [cr.left_top(), cr.right_top()],
                            Stroke::new(1.0, Color32::WHITE),
                        );
                        p.line_segment(
                            [cr.left_top(), cr.left_bottom()],
                            Stroke::new(1.0, Color32::WHITE),
                        );
                    }
                }
                if !cell.text.is_empty() {
                    let f = if cell.bold.unwrap_or(self.bold) {
                        bold_font.clone()
                    } else {
                        font.clone()
                    };
                    let color = vb_color(cell.fore.unwrap_or(self.fore));
                    let (anchor, x) = match cell.align {
                        1 => (Align2::LEFT_CENTER, cr.left() + 2.0),
                        7 => (Align2::RIGHT_CENTER, cr.right() - 2.0),
                        4 => (Align2::CENTER_CENTER, cr.center().x),
                        _ if fixed => (Align2::CENTER_CENTER, cr.center().x),
                        _ => (Align2::LEFT_CENTER, cr.left() + 2.0),
                    };
                    p.with_clip_rect(cr.intersect(p.clip_rect())).text(
                        Pos2::new(x, cr.center().y),
                        anchor,
                        &cell.text,
                        f,
                        color,
                    );
                }
            }
        }
    }
}
