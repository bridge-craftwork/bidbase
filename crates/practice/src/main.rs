//! `bidbase-practice`: Nelson Ford's BidBase Bidding Practice program,
//! ported from VB6 (aeyec-2026/source/BidBase Bidding Practice.frm).

// A window, not a console, on Windows (release builds).
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() -> eframe::Result {
    let (title, size, icon) = bidbase_practice::window();
    let mut viewport = egui::ViewportBuilder::default()
        .with_title(title)
        .with_inner_size(size)
        .with_resizable(false);
    if let Some(icon) = icon {
        viewport = viewport.with_icon(icon);
    }
    eframe::run_native(
        "bidbase-practice",
        eframe::NativeOptions {
            viewport,
            ..Default::default()
        },
        Box::new(|cc| Ok(Box::new(bidbase_practice::Shell::new(cc)))),
    )
}
