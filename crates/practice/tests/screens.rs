//! The port's screens, rendered offscreen, for comparing with Nelson's
//! (his notes' "BidBase Practice.jpg"). Set BIDBASE_SHOTS to a directory
//! to write PNGs; without it the screens are still drawn, so a panic while
//! drawing fails the test.

use bidbase_practice::{App, Settings};
use bidbase_vbui::{Event, Key};
use egui_kittest::Harness;

fn app(deal: u32) -> App {
    let mut a = App::with_settings(Settings {
        deal_number: deal,
        ..Default::default()
    });
    a.wait_for_dda();
    a
}

fn click(a: &mut App, name: &str, index: Option<i32>) {
    a.handle(Event::Click(Key::new(name, index)));
}

fn shoot(name: &str, mut app: App) {
    let size = app.window_size();
    let mut first = true;
    let mut harness = Harness::builder()
        .with_size(size)
        .with_pixels_per_point(1.0)
        .build_ui(move |ui| {
            // Fonts apply from the next frame on.
            if first {
                bidbase_vbui::install(ui.ctx());
                first = false;
                return;
            }
            app.ui(ui);
        });
    harness.run_steps(4);
    let Ok(dir) = std::env::var("BIDBASE_SHOTS") else {
        return;
    };
    let image = harness.render().expect("offscreen render");
    image
        .save(std::path::Path::new(&dir).join(format!("{name}.png")))
        .expect("write screenshot");
}

#[test]
fn opening_screen() {
    // Start-up makes the dealer's first call.
    let a = app(31118);
    assert_eq!(a.calls.len(), 1);
    assert!(a.dda.is_some());
    shoot("opening", a);
}

#[test]
fn bidding_to_the_end() {
    let mut a = app(24);
    for _ in 0..40 {
        click(&mut a, "b_Ok", None);
    }
    // The auction ended and the last click showed "Bidding's over."
    assert!(a.calls.len() >= 4);
    assert!(a.screen.visible("f_Msg", None));
    assert!(
        a.screen.text("t_Disclose", None).contains("Final contract")
            || a.calls.iter().all(|c| c.token == "P")
    );
    shoot("auction-over", a);
}

#[test]
fn next_deal_moves_the_dealer() {
    let mut a = app(100);
    let d = a.settings.dealer;
    click(&mut a, "b_DealBackFwd", Some(1));
    assert_eq!(a.settings.deal_number, 101);
    assert_ne!(a.settings.dealer, d);
    a.handle(Event::Menu(Key::new("mHide", None)));
    a.wait_for_dda();
    shoot("hidden", a);
}

#[test]
fn bid_box_and_prompts() {
    let mut a = app(31118);
    // Clicking a call other than BidBase's last asks whether to replace it.
    a.handle(Event::GridClick {
        key: Key::new("g_BidBox", None),
        row: 6,
        col: 4,
    });
    assert!(a.screen.visible("f_Msg", None));
    shoot("msgbox", a);
    let mut b = app(31118);
    click(&mut b, "b_Deal", None);
    assert!(b.screen.visible("f_Inbox", None));
    shoot("inbox", b);
}
