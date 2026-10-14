//! Nelson Ford's BidBase Bidding Practice program, ported from VB6.
//!
//! The screen is his form (`BidBase Bidding Practice.frm`) drawn by
//! bidbase-vbui; the deals are his dealer's (`dealer`); the calls come
//! from bidbase-engine with his 2026 rules; the DDA table from
//! bridge-solver where he calls dds.dll. docs/PRACTICE-PROGRAM.md describes
//! his program; the end of that file lists where this port departs from it.
//!
//! The handlers below follow his (`b_Deal_Click`, `g_BidBox_Click`, ...).
//! His message boxes (`MyMsgBox`, the `f_Msg` frame) and prompts
//! (`f_Inbox`) wait in a loop; here they show the frame and remember what
//! to do with the answer ([`Pending`]).

pub mod cards;
pub mod dda;
pub mod dealer;

use bidbase_engine::convert::{call_from_bidbase, call_to_bidbase, hand_to_bidbase, vul_for};
use bidbase_engine::lookup::{Request, Scoring};
use bidbase_engine::table::{self, Missing};
use bidbase_engine::{Bidder, RuleBase};
use bidbase_vbui::{Event, Key, Screen};
use bridge_types::{Auction, Card, DdTable, Deal, Direction, Suit, Vulnerability};
use egui::{Rect, Vec2};
use serde::{Deserialize, Serialize};

use dealer::{VbRnd, MAX_DEAL};

const FRM: &[u8] = include_bytes!("../../../aeyec-2026/source/BidBase Bidding Practice.frm");
const FRX: &[u8] = include_bytes!("../../../aeyec-2026/source/BidBase Bidding Practice.frx");
const BIDS: &str = include_str!("../../../data/2026/Bids.csv");
const NAMES: &str = include_str!("../../../data/2026/BidNames.csv");

const HELP_URL: &str = "http://www.aeyec.com/bidbase/Notes/CardShark_BidBase_Practice_Program.htm";
const CONVENTIONS_URL: &str = "http://www.aeyec.com/bidbase/Notes/Index.htm";

const MENU_HEIGHT: f32 = 22.0;

/// The program's seats: `p_Hand(1..4)` and the auction grid's columns.
pub fn seat_of_box(index: i32) -> Direction {
    dealer::SEATS[(index.clamp(1, 4) - 1) as usize]
}

fn seat_number(d: Direction) -> usize {
    dealer::SEATS.iter().position(|&s| s == d).unwrap()
}

/// Suits in the order his `DisplayHand` draws them.
pub const DISPLAY_SUITS: [Suit; 4] = [Suit::Spades, Suit::Hearts, Suit::Diamonds, Suit::Clubs];

/// A hand's cards in display order: by suit, then high to low.
pub fn display_order(hand: &bridge_types::Hand) -> Vec<Card> {
    let mut out = Vec::new();
    for s in DISPLAY_SUITS {
        let mut cs = hand.cards_in_suit(s);
        cs.sort_by_key(|c| std::cmp::Reverse(c.rank));
        out.extend(cs);
    }
    out
}

/// The window's title, size (menu bar plus form) and icon, from the form.
pub fn window() -> (String, Vec2, Option<egui::IconData>) {
    let screen = Screen::new(FRM, FRX.to_vec()).expect("Nelson's practice form parses");
    let icon = screen
        .icon()
        .and_then(bidbase_vbui::ico::decode)
        .map(|(width, height, rgba)| egui::IconData {
            rgba,
            width,
            height,
        });
    (
        screen.title(),
        screen.size() + Vec2::new(0.0, MENU_HEIGHT),
        icon,
    )
}

/// `L_Vul` as he labels it, from South's side.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum Vul {
    #[default]
    None,
    /// North-South vulnerable.
    Unfav,
    /// East-West vulnerable.
    Fav,
    Both,
}

impl Vul {
    fn label(self) -> &'static str {
        match self {
            Vul::None => "",
            Vul::Unfav => "Unfav.",
            Vul::Fav => "Fav.",
            Vul::Both => "Both",
        }
    }
    fn next(self) -> Vul {
        match self {
            Vul::None => Vul::Unfav,
            Vul::Unfav => Vul::Fav,
            Vul::Fav => Vul::Both,
            Vul::Both => Vul::None,
        }
    }
    fn board(self) -> Vulnerability {
        match self {
            Vul::None => Vulnerability::None,
            Vul::Unfav => Vulnerability::NorthSouth,
            Vul::Fav => Vulnerability::EastWest,
            Vul::Both => Vulnerability::Both,
        }
    }
}

/// What is kept between sessions (his `BBPractice.cfg`).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct Settings {
    pub deal_number: u32,
    /// N, E, S or W.
    pub dealer: char,
    pub keep_dealer: bool,
    pub hide_hands: bool,
    pub auto_bid: bool,
    pub vul: Vul,
}

impl Default for Settings {
    fn default() -> Self {
        Settings {
            deal_number: 1,
            dealer: 'S',
            keep_dealer: false,
            hide_hands: false,
            auto_bid: true,
            vul: Vul::None,
        }
    }
}

/// What to do with the answer to a message box or prompt.
#[derive(Debug, Clone)]
enum Pending {
    /// Just a message.
    Nothing,
    DealNumber,
    Dealer,
    /// Bid box, case 1: replace BidBase's last call with `mine`?
    Supersede {
        mine: String,
    },
    /// Bid box, case 2: BidBase's call `bb` (made, not yet recorded) differs from `mine`.
    ChangeBb {
        mine: String,
        bb: Made,
    },
    /// Bid w/ Analysis: the call(s) to watch.
    Analysis,
    MixSeats,
}

/// One call as made, kept so the grid, the auction and the disclosure agree.
#[derive(Debug, Clone)]
pub struct Made {
    pub seat: Direction,
    /// BidBase token: "1C", "P", "4N@30".
    pub token: String,
    pub rule: Option<i64>,
    pub bon: String,
    pub disclosure: String,
    /// Why BidBase passed for want of a rule.
    pub missing: Option<Missing>,
    /// The engine's key for the lookup (for messages).
    pub key: String,
}

/// The double-dummy table: worked out off the drawing thread on native.
enum DdaJob {
    Idle,
    #[cfg(not(target_arch = "wasm32"))]
    Running(std::sync::mpsc::Receiver<(u64, DdTable)>),
    #[cfg(target_arch = "wasm32")]
    Due,
}

pub struct App {
    pub screen: Screen,
    bidder: Bidder,
    pub settings: Settings,
    pub deal: Deal,
    rnd: VbRnd,
    pub calls: Vec<Made>,
    pending: Option<Pending>,
    dda_job: DdaJob,
    /// Bumped with every new layout, so a late DDA for an old one is dropped.
    generation: u64,
    pub dda: Option<DdTable>,
    /// Seats last chosen for Mix Up Cards.
    mixed: Vec<Direction>,
    /// A link to open (the help pages), or "quit:", taken by the shell.
    pub open_url: Option<String>,
    /// Stay on Top changed: the shell applies it.
    pub on_top_changed: Option<bool>,
}

impl Default for App {
    fn default() -> Self {
        Self::new()
    }
}

impl App {
    pub fn new() -> Self {
        Self::with_settings(Settings::default())
    }

    /// Start as his `Form_Load` does: the saved deal, its DDA, and the
    /// dealer's first call.
    pub fn with_settings(settings: Settings) -> Self {
        let screen = Screen::new(FRM, FRX.to_vec()).expect("Nelson's practice form parses");
        let base = RuleBase::from_csv(BIDS, NAMES).expect("the 2026 rules load");
        let mut bidder = Bidder::new(base);
        bidder.seed(seed_from_clock());
        let mut app = App {
            screen,
            bidder,
            settings,
            deal: Deal::new(),
            rnd: VbRnd::default(),
            calls: Vec::new(),
            pending: None,
            dda_job: DdaJob::Idle,
            generation: 0,
            dda: None,
            mixed: Vec::new(),
            open_url: None,
            on_top_changed: None,
        };
        app.set_up_screen();
        let n = app.settings.deal_number.clamp(1, MAX_DEAL);
        app.new_deal(n);
        app.bid_next();
        app
    }

    fn set_up_screen(&mut self) {
        // SetUpBiddingGrid (frm:1950).
        let g = self.screen.grid("g_BidBox");
        for c in 0..5 {
            g.col_width[c] = 500;
        }
        for r in 0..7 {
            g.row_height[r] = 255;
        }
        g.row_height[7] = 300;
        for r in 0..7 {
            for c in 0..5 {
                g.set_text(r, c, format!("{}{}", r + 1, &"CDHSN"[c..c + 1]));
                g.cell_mut(r, c).unwrap().align = 4;
            }
        }
        for (c, t) in ["Pass", "Dbl", "Rdbl", "BID", "BOX"].iter().enumerate() {
            g.set_text(7, c, *t);
            g.cell_mut(7, c).unwrap().align = 4;
        }
        // SetUpBidDisplayGrid (frm:1928).
        let g = self.screen.grid("g_Bidding");
        g.row_height[0] = 260;
        g.col_width = vec![530, 530, 530, 535];
        for (c, t) in ["S", "W", "N", "E"].iter().enumerate() {
            g.set_text(0, c, *t);
        }
        self.screen.set_visible("p_BidBoxLbl", None, false);
        // The convention-practice list sits just below the window; his
        // program shows it only when it enlarges the form for practice.
        self.screen.set_visible("cmb_FindDeals", None, false);
        self.screen.set_text("cmb_DDA", None, dda::CAPTION);
        let vul = if self.settings.vul == Vul::None {
            "none"
        } else {
            self.settings.vul.label()
        };
        self.screen.set_caption("L_Vul", None, vul);
        self.screen
            .set_value("ck_KeepDealer", None, self.settings.keep_dealer as i32);
        self.screen.set_visible("ck_KeepDealer", None, true);
        self.screen
            .set_value("mAutoBid", None, self.settings.auto_bid as i32);
        self.screen
            .set_value("mHide", None, self.settings.hide_hands as i32);
        // Features not in this port yet are greyed out in the menus.
        for m in [
            "mLoadDeal",
            "mLoadPbn",
            "mLoadPbnDealNum",
            "mRunEditor",
            "mAverageDDA",
            "m100DealAvg",
            "mCreateConv",
            "mFindDeals",
            "mPractice",
            "mUpdateConventionFiles",
        ] {
            self.screen.set_enabled(m, None, false);
        }
        self.screen.set_enabled("mScramble", Some(1), false);
    }

    /// Menu bar plus the form.
    pub fn window_size(&self) -> Vec2 {
        self.screen.size() + Vec2::new(0.0, MENU_HEIGHT)
    }

    fn dealer(&self) -> Direction {
        Direction::from_char(self.settings.dealer).unwrap_or(Direction::South)
    }

    fn set_dealer(&mut self, d: Direction) {
        self.settings.dealer = d.to_char();
        self.screen
            .set_caption("L_Dealer", None, d.to_char().to_string());
    }

    // --- deals -------------------------------------------------------------

    /// Deal number `n`, show it, clear the auction, start the DDA
    /// (`FinishDealSetup`, frm:2140).
    fn new_deal(&mut self, n: u32) {
        self.settings.deal_number = n;
        self.deal = dealer::deal(n);
        // Mix Up Cards goes on from a generator seeded by the deal.
        self.rnd = VbRnd::default();
        self.rnd.reseed(-(n as f32));
        self.screen.set_caption("L_Deal", None, n.to_string());
        let d = self.dealer();
        self.set_dealer(d);
        self.restart();
        self.start_dda();
    }

    /// Repeat: same deal, clear the auction (`b_RestartDeal_Click`).
    fn restart(&mut self) {
        self.calls.clear();
        self.screen.set_text("t_Disclose", None, "");
        self.screen
            .set_caption("b_ShowAnalysis", None, "Bid w/ &Analysis");
        self.show_auction();
    }

    fn start_dda(&mut self) {
        self.generation += 1;
        self.dda = None;
        let list = self.screen.list("cmb_DDA", None);
        list.clear();
        list.push("Calculating DDA...".into());
        #[cfg(not(target_arch = "wasm32"))]
        {
            let (tx, rx) = std::sync::mpsc::channel();
            let deal = self.deal.clone();
            let generation = self.generation;
            std::thread::spawn(move || {
                let _ = tx.send((generation, bridge_solver::solve_dd_table(&deal)));
            });
            self.dda_job = DdaJob::Running(rx);
        }
        #[cfg(target_arch = "wasm32")]
        {
            self.dda_job = DdaJob::Due;
        }
    }

    /// Pick up a finished DDA. True while one is still being worked out.
    fn poll_dda(&mut self) -> bool {
        let done = match &self.dda_job {
            DdaJob::Idle => return false,
            #[cfg(not(target_arch = "wasm32"))]
            DdaJob::Running(rx) => match rx.try_recv() {
                Ok((g, t)) if g == self.generation => Some(t),
                Ok(_) => return true,
                Err(std::sync::mpsc::TryRecvError::Empty) => return true,
                Err(_) => None,
            },
            #[cfg(target_arch = "wasm32")]
            DdaJob::Due => Some(bridge_solver::solve_dd_table(&self.deal)),
        };
        self.dda_job = DdaJob::Idle;
        if let Some(t) = done {
            let lines = dda::lines(&self.deal, &t);
            *self.screen.list("cmb_DDA", None) = lines;
            self.dda = Some(t);
        }
        false
    }

    /// For tests: wait for the DDA.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn wait_for_dda(&mut self) {
        while self.poll_dda() {
            std::thread::sleep(std::time::Duration::from_millis(5));
        }
    }

    // --- the auction -------------------------------------------------------

    fn tokens(&self) -> Vec<String> {
        self.calls.iter().map(|c| c.token.clone()).collect()
    }

    fn auction(&self) -> Auction {
        let mut a = Auction::new(self.dealer());
        for c in &self.calls {
            a.add_call(call_from_bidbase(&c.token).unwrap_or(bridge_types::Call::Pass));
        }
        a
    }

    /// The engine's call for the seat now to call (`MakeBid`, frm:6061),
    /// not yet recorded.
    fn bb_call(&mut self, trace: bool) -> (Made, Vec<bidbase_engine::lookup::Tried>) {
        let auction = self.auction();
        let tokens = self.tokens();
        let seat = auction.next_caller();
        let template = Request {
            hand: hand_to_bidbase(self.deal.hand(seat)),
            vul: vul_for(self.settings.vul.board(), seat),
            scoring: Scoring::Matchpoints,
            disclose: true,
            trace,
            ..Default::default()
        };
        let (key, _) = table::key_for(&tokens);
        let (tc, tried) = table::next_call_with(&mut self.bidder, &auction, &tokens, template);
        let bon = tc
            .found
            .as_ref()
            .and_then(|f| self.bidder.base.by_id(f.rule_id))
            .map(|r| r.bid_order.clone())
            .unwrap_or_default();
        let made = Made {
            seat,
            token: tc.token.clone(),
            rule: tc.found.as_ref().map(|f| f.rule_id),
            bon,
            disclosure: tc.found.map(|f| f.disclosure).unwrap_or_default(),
            missing: tc.missing,
            key,
        };
        (made, tried)
    }

    fn auction_over(&self) -> bool {
        self.calls.len() >= 4 && self.calls.iter().rev().take(3).all(|c| c.token == "P")
    }

    /// Record a call and show it (`MakeBid` steps 5-8).
    fn record(&mut self, m: Made) {
        let mut text = String::new();
        match m.missing {
            Some(Missing::UnknownAuction) => {
                text = format!("No BidBase entries for \"{}\".", m.key)
            }
            Some(Missing::Illegal) => {
                text = "BidBase's call was not legal here, so it passed.".into()
            }
            _ => {}
        }
        if m.token != "P" && !m.disclosure.trim().is_empty() {
            if !text.is_empty() {
                text.push('\n');
            }
            text.push_str(m.disclosure.trim_end());
        }
        self.calls.push(m);
        self.show_auction();
        if self.auction_over() {
            let line = match self.auction().final_contract() {
                Some(fc) => format!(
                    "Final contract: {} by {}",
                    fc.to_pbn(),
                    fc.declarer.to_char()
                ),
                None => "Passed out.".into(),
            };
            text = if text.is_empty() {
                line
            } else {
                format!("{line}\n\n{text}")
            };
        }
        self.set_disclosure(&text);
    }

    fn set_disclosure(&mut self, text: &str) {
        self.screen
            .set_text("t_Disclose", None, text.replace("\r\n", "\n"));
    }

    /// BB Bid (`b_Ok_Click`, `ProcessBid`).
    fn bid_next(&mut self) {
        if self.auction_over() {
            self.msgbox(
                "Three consecutive passes.\nBidding's over.",
                MsgKind::Ok,
                Pending::Nothing,
            );
            return;
        }
        let (m, _) = self.bb_call(false);
        self.record(m);
    }

    /// Calls as the grid shows them: "4N@30" as "4N".
    fn shown(token: &str) -> String {
        token.chars().take(2).collect()
    }

    /// `DisplayBid` (frm:6752): the dealer's column on row 1, left to
    /// right S W N E. Long auctions scroll: the last six rows show.
    fn show_auction(&mut self) {
        let first = seat_number(self.dealer());
        let rows_needed = (first + self.calls.len()).div_ceil(4).max(1);
        let skip = rows_needed.saturating_sub(6);
        let tokens: Vec<String> = self.calls.iter().map(|c| Self::shown(&c.token)).collect();
        let g = self.screen.grid("g_Bidding");
        for r in 1..g.rows {
            for c in 0..4 {
                g.set_text(r, c, "");
                if let Some(cell) = g.cell_mut(r, c) {
                    cell.align = 4;
                }
            }
        }
        for (i, t) in tokens.iter().enumerate() {
            let slot = first + i;
            let row = slot / 4;
            if row >= skip {
                g.set_text(row - skip + 1, slot % 4, t);
            }
        }
    }

    // --- the bid box -------------------------------------------------------

    /// A click on the bid box (`g_BidBox_Click`, frm:3475).
    fn bid_box(&mut self, row: usize, col: usize) {
        let mine = match (row, col) {
            (7, 0) => "P".to_string(),
            (7, 1) => "D".to_string(),
            (7, 2) => "R".to_string(),
            (7, _) => return, // the BID BOX label
            (r, c) => format!("{}{}", r + 1, &"CDHSN"[c..c + 1]),
        };
        // Case 1: a call differing from BidBase's last one may replace it.
        if let Some(last) = self.calls.last() {
            if Self::shown(&last.token) != mine && !self.auction_over() {
                let bb = Self::shown(&last.token);
                self.msgbox(
                    &format!("Do you want your bid of {mine}\nto replace BB's bid of {bb}?"),
                    MsgKind::YesNo,
                    Pending::Supersede { mine },
                );
                return;
            }
        }
        self.my_call(mine);
    }

    /// Bid box, case 2: the user's call for the seat now to call.
    fn my_call(&mut self, mine: String) {
        if self.auction_over() {
            self.msgbox(
                "Three consecutive passes.\nBidding's over.",
                MsgKind::Ok,
                Pending::Nothing,
            );
            return;
        }
        let Some(call) = call_from_bidbase(&mine) else {
            return;
        };
        if !self.auction().is_legal(&call) {
            self.msgbox(
                &format!("{mine} is not a legal call here."),
                MsgKind::Ok,
                Pending::Nothing,
            );
            return;
        }
        let (bb, _) = self.bb_call(false);
        if Self::shown(&bb.token) == mine {
            let disclosure = bb.disclosure.clone();
            self.record(bb);
            let t = format!(
                "         *** Your bid is the same as BB's. ***\n\n            Click a bid for the next player or\n            click \"BB Bid\" for BB to do it\n            or click right arrow for next deal.\n\n{disclosure}"
            );
            self.set_disclosure(&t);
        } else {
            let text = format!(
                "Did you want to change BidBase's bid of {}\nto your bid of {mine} in the database?\n(Click No to keep your bid for now.)",
                Self::shown(&bb.token)
            );
            self.msgbox(&text, MsgKind::YesNoCancel, Pending::ChangeBb { mine, bb });
        }
    }

    /// The user's call in place of BidBase's, keeping the auction and the
    /// grid in step (the VB changed only the grid).
    fn user_made(&self, mine: &str) -> Made {
        Made {
            seat: self.auction().next_caller(),
            token: call_from_bidbase(mine)
                .map(|c| call_to_bidbase(&c))
                .unwrap_or_else(|| mine.to_string()),
            rule: None,
            bon: String::new(),
            disclosure: String::new(),
            missing: None,
            key: String::new(),
        }
    }

    // --- Edit Bid, analysis ------------------------------------------------

    /// Edit Bid (`b_Edit_Click`): what the BidBase Editor would import
    /// (`EditPractice.txt`), and the last call taken back so the next BB
    /// Bid makes it again.
    fn edit_bid(&mut self) {
        let Some(last) = self.calls.pop() else { return };
        let prior: Vec<String> = self
            .calls
            .iter()
            .map(|c| c.token.replace('?', ""))
            .collect();
        let mut source = format!(
            "Deal #{}-{}",
            self.settings.deal_number,
            self.dealer().to_char()
        );
        if let Some(id) = last.rule {
            source.push_str(&format!(" ID={id} BON={}", last.bon));
        }
        let text = format!(
            "For the BidBase Editor (EditPractice.txt):\n\n{}\n{}\n{}",
            hand_to_bidbase(self.deal.hand(last.seat)),
            prior.join("-"),
            source
        );
        self.show_auction();
        self.set_disclosure(&text);
    }

    /// Bid w/ Analysis: the last call made again, with the rules tried
    /// for the watched call(s).
    fn analysis(&mut self, watch: &str) {
        if self.calls.pop().is_none() {
            self.msgbox("No call to analyze yet.", MsgKind::Ok, Pending::Nothing);
            return;
        }
        let (m, tried) = self.bb_call(true);
        let words: Vec<&str> = watch.split_whitespace().collect();
        let all = words.contains(&"ALL");
        let mut lines = vec![format!(
            "Analysis of {}'s call after \"{}\":",
            m.seat.to_char(),
            m.key
        )];
        for t in &tried {
            let bid = t.bid.trim();
            if all || words.iter().any(|w| w.eq_ignore_ascii_case(bid)) {
                let why = match &t.failed {
                    Some((test, detail)) => format!("{test}: {detail}"),
                    None => "fits".to_string(),
                };
                lines.push(format!("#{} {bid} (BON {}): {why}", t.rule_id, t.bid_order));
            }
        }
        lines.push(format!("BidBase bids {}.", Self::shown(&m.token)));
        self.record(m);
        self.set_disclosure(&lines.join("\n"));
        self.screen
            .set_caption("b_ShowAnalysis", None, "Rebid w/Analysis");
    }

    // --- options -----------------------------------------------------------

    /// Rotate so the hand at `from` sits South; the dealer moves with the
    /// cards (`mRotate_Click`), done on one seat array.
    fn rotate(&mut self, from: Direction) {
        let k = seat_number(from);
        let old = self.deal.clone();
        for i in 0..4 {
            self.deal.set_hand(
                dealer::SEATS[i],
                old.hand(dealer::SEATS[(i + k) % 4]).clone(),
            );
        }
        let d = dealer::SEATS[(seat_number(self.dealer()) + 4 - k) % 4];
        self.set_dealer(d);
        self.restart();
        self.start_dda();
    }

    /// Mix Up Cards: pool the seats' cards, shuffle, deal them back.
    fn mix(&mut self, seats: &[Direction]) {
        let mut pool: Vec<Card> = seats
            .iter()
            .flat_map(|s| self.deal.hand(*s).cards().to_vec())
            .collect();
        dealer::scramble(&mut self.rnd, &mut pool);
        let mut at = 0;
        for s in seats {
            let n = self.deal.hand(*s).len();
            self.deal.set_hand(
                *s,
                bridge_types::Hand::from_cards(pool[at..at + n].to_vec()),
            );
            at += n;
        }
        self.mixed = seats.to_vec();
        self.screen.set_enabled("mScramble", Some(1), true);
        self.restart();
        self.start_dda();
    }

    /// Save Deal: the deal as his `Notes\SavedDeals.txt` gets it.
    pub fn saved_deal_text(&mut self) -> String {
        let h = |d| hand_to_bidbase(self.deal.hand(d));
        let mut out = format!(
            "\r\n\r\nDeal #{}\r\nN:  {}\r\nS:  {}\r\nW:  {}\r\nE:  {}\r\nDealer: {}\r\nBidding:\r\n",
            self.settings.deal_number,
            h(Direction::North),
            h(Direction::South),
            h(Direction::West),
            h(Direction::East),
            self.dealer().to_char()
        );
        for chunk in self.calls.chunks(4) {
            let line: Vec<String> = chunk.iter().map(|c| c.token.clone()).collect();
            out.push_str(&line.join("-"));
            out.push_str("\r\n");
        }
        for l in self.screen.list("cmb_DDA", None).iter() {
            out.push_str(l);
            out.push_str("\r\n");
        }
        out.push_str("\r\n");
        out
    }

    // --- message boxes ------------------------------------------------------

    /// His `MyMsgBox` (frm:8247) on the `f_Msg` frame, 200 pixels down.
    fn msgbox(&mut self, text: &str, kind: MsgKind, then: Pending) {
        let lines = text.lines().count().max(1) as i32;
        let longest = text.lines().map(|l| l.chars().count()).max().unwrap_or(0) as f32;
        // TextPixelWidth at 12 pt bold: about 9 pixels a character.
        let label_w = (longest * 9.0 * 15.8) as i32;
        let frame_w_px = ((label_w + 400) as f32 / 15.3) as i32;
        let form_w_px = self.screen.size().x as i32;
        let frame_h_px = 20 * lines + 29 + 45;
        let s = &mut self.screen;
        s.set_caption("L_Caption", None, "  CardShark BidBase");
        s.set_caption("L_MsgBox", None, text);
        let (_, label_top, _, _) = s.bounds("L_MsgBox", None);
        s.resize("L_MsgBox", None, label_w, 300 * lines);
        let button_top = label_top + 300 * lines + 50;
        s.move_to(
            "f_Msg",
            None,
            ((form_w_px - frame_w_px) / 2).max(0) * 15,
            200 * 15,
        );
        s.resize("f_Msg", None, frame_w_px * 15, frame_h_px * 15);
        s.resize("L_Caption", None, frame_w_px * 15, 330);
        s.resize(
            "L_MsgBackground",
            None,
            label_w + 200,
            (frame_h_px as f32 * 14.67) as i32 - 385,
        );
        for i in 0..3 {
            s.set_visible("b_MsgInput", Some(i), false);
            let (bl, _, _, _) = s.bounds("b_MsgInput", Some(i));
            s.move_to("b_MsgInput", Some(i), bl, button_top);
        }
        s.set_visible("t_MsgInput", None, false);
        s.set_visible("b_MsgInput", Some(0), true);
        match kind {
            MsgKind::Ok => s.set_caption("b_MsgInput", Some(0), "OK"),
            MsgKind::YesNo | MsgKind::YesNoCancel => {
                s.set_caption("b_MsgInput", Some(0), "Yes");
                s.set_visible("b_MsgInput", Some(1), true);
                s.set_visible("b_MsgInput", Some(2), kind == MsgKind::YesNoCancel);
            }
            MsgKind::Input(ref default) => {
                s.set_caption("b_MsgInput", Some(0), "Ok");
                let (tl, _, _, th) = s.bounds("t_MsgInput", None);
                s.move_to("t_MsgInput", None, tl, button_top);
                s.resize("t_MsgInput", None, (label_w - tl + 200).max(1500), th);
                s.set_text("t_MsgInput", None, default.clone());
                s.set_visible("t_MsgInput", None, true);
            }
        }
        s.set_visible("t_Calculating", None, true);
        s.bring_to_front("t_Calculating", None);
        s.set_visible("f_Msg", None, true);
        s.bring_to_front("f_Msg", None);
        self.pending = Some(then);
    }

    /// His `f_Inbox` prompt.
    fn inbox(&mut self, label: &str, default: &str, then: Pending) {
        self.screen.set_caption("L_InboxLbl", None, label);
        self.screen.set_text("t_Inbox", None, default);
        self.screen.set_visible("f_Inbox", None, true);
        self.pending = Some(then);
    }

    fn close_boxes(&mut self) {
        self.screen.set_visible("f_Msg", None, false);
        self.screen.set_visible("f_Inbox", None, false);
        self.screen.set_visible("t_Calculating", None, false);
    }

    fn modal(&self) -> Option<&'static str> {
        if self.screen.visible("f_Msg", None) {
            Some("f_Msg")
        } else if self.screen.visible("f_Inbox", None) {
            Some("f_Inbox")
        } else {
            None
        }
    }

    /// An answer: "yes", "no", "cancel", "" for OK, or the text typed.
    fn answer(&mut self, answer: String) {
        self.close_boxes();
        let Some(p) = self.pending.take() else { return };
        match p {
            Pending::Nothing => {}
            Pending::DealNumber => match answer.trim().parse::<u64>() {
                Ok(n) if (1..=MAX_DEAL as u64).contains(&n) => self.new_deal(n as u32),
                Ok(n) if n > MAX_DEAL as u64 => self.msgbox(
                    "Deal number cannot be greater than 9,999,999.",
                    MsgKind::Ok,
                    Pending::Nothing,
                ),
                _ => {}
            },
            Pending::Dealer => {
                let d = answer
                    .trim()
                    .chars()
                    .next()
                    .and_then(|c| Direction::from_char(c.to_ascii_uppercase()))
                    .unwrap_or(Direction::South);
                self.set_dealer(d);
                self.restart();
            }
            Pending::Supersede { mine } => {
                if answer == "yes" {
                    self.calls.pop();
                    let m = self.user_made(&mine);
                    let legal =
                        call_from_bidbase(&m.token).is_some_and(|c| self.auction().is_legal(&c));
                    if legal {
                        self.record(m);
                        self.set_disclosure("Your call replaces BidBase's.\n(Adding it to the rules needs the BidBase Editor.)");
                    } else {
                        self.show_auction();
                        self.msgbox(
                            &format!("{mine} is not a legal call here."),
                            MsgKind::Ok,
                            Pending::Nothing,
                        );
                    }
                } else {
                    self.my_call(mine);
                }
            }
            Pending::ChangeBb { mine, bb } => match answer.as_str() {
                "yes" => {
                    self.record(bb);
                    self.edit_bid();
                }
                "no" => {
                    let m = self.user_made(&mine);
                    self.record(m);
                }
                _ => self.record(bb),
            },
            Pending::Analysis => {
                let mut w = answer.to_ascii_uppercase();
                if w.contains("ALL") {
                    w = "ALL".into();
                } else if let Some(c) = w.chars().next().filter(|c| c.is_ascii_digit()) {
                    // A rule bidding "2x" may be the 2S asked about.
                    w = format!("{w} {c}x");
                }
                if !w.trim().is_empty() {
                    self.analysis(&w);
                }
            }
            Pending::MixSeats => {
                let mut seats: Vec<Direction> = Vec::new();
                for d in answer
                    .to_ascii_uppercase()
                    .chars()
                    .filter_map(Direction::from_char)
                {
                    if !seats.contains(&d) {
                        seats.push(d);
                    }
                }
                match seats.len() {
                    0 => {}
                    1 => self.msgbox("You must specify at least 2 hands to Scramble.", MsgKind::Ok, Pending::Nothing),
                    4 => self.msgbox(
                        "You cannot Scramble more than 3 hands.\nTo Scramble all 4, just click on Deal.",
                        MsgKind::Ok,
                        Pending::Nothing,
                    ),
                    _ => self.mix(&seats),
                }
            }
        }
    }

    fn not_yet(&mut self, what: &str) {
        self.msgbox(
            &format!("{what}\nis not in this port yet."),
            MsgKind::Ok,
            Pending::Nothing,
        );
    }

    // --- events -------------------------------------------------------------

    pub fn handle(&mut self, e: Event) {
        if let Some(frame) = self.modal() {
            let key = match &e {
                Event::Click(k) | Event::Enter(k) | Event::Change(k) | Event::DblClick(k) => {
                    Some(k)
                }
                Event::GridClick { key, .. } | Event::MouseDown { key, .. } => Some(key),
                Event::Menu(_) => None,
            };
            if !key.is_some_and(|k| self.screen.within(k, frame)) {
                return;
            }
        }
        match e {
            Event::Click(k) => self.click(&k),
            Event::Enter(k) if k.is("t_Inbox") || k.is("t_MsgInput") => {
                let t = self.screen.text(&k.name, None);
                self.answer(t);
            }
            Event::Change(k) if k.is("t_Inbox") => {
                // The dealer prompt closes on S, W, N or E (frm:5705).
                if matches!(self.pending, Some(Pending::Dealer)) {
                    let t = self
                        .screen
                        .text("t_Inbox", None)
                        .trim()
                        .to_ascii_uppercase();
                    if matches!(t.as_str(), "S" | "W" | "N" | "E") {
                        self.answer(t);
                    }
                }
            }
            Event::GridClick { key, row, col } if key.is("g_BidBox") => self.bid_box(row, col),
            Event::Menu(k) => self.menu(&k),
            _ => {}
        }
    }

    fn click(&mut self, k: &Key) {
        match k.name.as_str() {
            "b_msginput" => {
                let a = match k.index {
                    Some(1) => "no".to_string(),
                    Some(2) => "cancel".to_string(),
                    _ if self.screen.visible("t_MsgInput", None) => {
                        self.screen.text("t_MsgInput", None)
                    }
                    _ if self.screen.caption("b_MsgInput", Some(0)) == "Yes" => "yes".to_string(),
                    _ => String::new(),
                };
                self.answer(a);
            }
            "b_donedealnum" => {
                let t = self.screen.text("t_Inbox", None);
                self.answer(t);
            }
            "b_deal" => {
                let n = self.settings.deal_number.to_string();
                self.inbox(" Enter a deal number:", &n, Pending::DealNumber);
            }
            "b_dealbackfwd" => {
                let forward = k.index == Some(1);
                let n = self.settings.deal_number;
                let n = if forward {
                    n % MAX_DEAL + 1
                } else if n <= 1 {
                    MAX_DEAL
                } else {
                    n - 1
                };
                if self.screen.value("ck_KeepDealer", None) == 0 {
                    let d = self.dealer();
                    self.set_dealer(if forward { d.next() } else { d.prev() });
                }
                self.new_deal(n);
                self.bid_next();
            }
            "b_restartdeal" => self.restart(),
            "b_vul" => {
                self.settings.vul = self.settings.vul.next();
                self.screen
                    .set_caption("L_Vul", None, self.settings.vul.label());
            }
            "b_dealer" => {
                self.screen.set_value("ck_KeepDealer", None, 0);
                self.settings.keep_dealer = false;
                self.inbox(" Enter dealer\n (N S E W)", "", Pending::Dealer);
            }
            "ck_keepdealer" => {
                self.settings.keep_dealer = self.screen.value("ck_KeepDealer", None) != 0
            }
            "b_ok" => self.bid_next(),
            "b_edit" => self.edit_bid(),
            "b_showanalysis" => self.msgbox(
                "Enter bid to watch for or ALL:",
                MsgKind::Input("ALL".into()),
                Pending::Analysis,
            ),
            "b_enterhands" => self.not_yet("Enter Hands"),
            _ => {}
        }
    }

    fn menu(&mut self, k: &Key) {
        match k.name.as_str() {
            "msavedeal" => {
                let text = self.saved_deal_text();
                #[cfg(not(target_arch = "wasm32"))]
                if let Some(path) = save_deal_file(&text) {
                    self.msgbox(
                        &format!("Saved to\n{}", path.display()),
                        MsgKind::Ok,
                        Pending::Nothing,
                    );
                }
                self.set_disclosure(text.trim());
            }
            "mexit" => self.open_url = Some("quit:".into()),
            "mautobid" => {
                self.settings.auto_bid = !self.settings.auto_bid;
                self.screen
                    .set_value("mAutoBid", None, self.settings.auto_bid as i32);
            }
            "mhide" => {
                self.settings.hide_hands = !self.settings.hide_hands;
                self.screen
                    .set_value("mHide", None, self.settings.hide_hands as i32);
            }
            "mstayontop" => {
                let on = self.screen.value("mStayOnTop", None) == 0;
                self.screen.set_value("mStayOnTop", None, on as i32);
                self.on_top_changed = Some(on);
            }
            "menterhands" => self.not_yet("Enter Hands"),
            "mscramble" => {
                if k.index == Some(1) {
                    if self.mixed.is_empty() {
                        self.msgbox(
                            "No record of hands selected.\nPlease enter again.",
                            MsgKind::Ok,
                            Pending::Nothing,
                        );
                    } else {
                        let seats = self.mixed.clone();
                        self.mix(&seats);
                    }
                } else {
                    self.msgbox(
                        "Enter the hands (N,S,E,W) to scramble.",
                        MsgKind::Input(String::new()),
                        Pending::MixSeats,
                    );
                }
            }
            "mrotate" => {
                let from = match k.index {
                    Some(1) => Direction::West,
                    Some(2) => Direction::North,
                    _ => Direction::East,
                };
                self.rotate(from);
            }
            "mshowpraccnvdocs" => self.open_url = Some(format!("{HELP_URL}#C1")),
            "mshowconvdocs" => self.open_url = Some(CONVENTIONS_URL.into()),
            "mhelp" => self.open_url = Some(HELP_URL.into()),
            _ => {}
        }
    }

    // --- drawing ------------------------------------------------------------

    pub fn ui(&mut self, ui: &mut egui::Ui) {
        if self.poll_dda() {
            ui.ctx()
                .request_repaint_after(std::time::Duration::from_millis(30));
        }
        let top = ui.max_rect().min;
        let menu_rect = Rect::from_min_size(top, Vec2::new(self.screen.size().x, MENU_HEIGHT));
        let mut menu_ui = ui.new_child(egui::UiBuilder::new().max_rect(menu_rect));
        let mut events = self.screen.menu_bar(&mut menu_ui);
        let origin = top + Vec2::new(0.0, MENU_HEIGHT);
        let App {
            screen,
            deal,
            settings,
            ..
        } = self;
        let hide = settings.hide_hands;
        events.extend(screen.show(ui, origin, &mut |ui, key, rect| {
            if key.is("p_Hand") {
                paint_hand(ui, key, rect, deal, hide);
            }
        }));
        // His accelerators: Enter is BB Bid (the default button), Alt+key
        // the underlined letters.
        let typing = ui.ctx().memory(|m| m.focused().is_some());
        let keys = ui.ctx().input(|i| {
            let mut v = Vec::new();
            if i.key_pressed(egui::Key::Enter) && !i.modifiers.alt {
                v.push("enter");
            }
            for (k, name) in [
                (egui::Key::D, "b_Deal"),
                (egui::Key::R, "b_RestartDeal"),
                (egui::Key::V, "b_Vul"),
                (egui::Key::L, "b_Dealer"),
                (egui::Key::B, "b_Ok"),
                (egui::Key::E, "b_Edit"),
                (egui::Key::A, "b_ShowAnalysis"),
            ] {
                if i.modifiers.alt && i.key_pressed(k) {
                    v.push(name);
                }
            }
            v
        });
        for k in keys {
            if k == "enter" {
                if self.modal().is_none() && !typing && events.is_empty() {
                    events.push(Event::Click(Key::new("b_Ok", None)));
                }
            } else {
                events.push(Event::Click(Key::new(k, None)));
            }
        }
        for e in events {
            self.handle(e);
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
enum MsgKind {
    Ok,
    YesNo,
    YesNoCancel,
    Input(String),
}

/// `DisplayHand` (dealer:178): N and S fan right 21 pixels a card, W and E
/// down 27; spades, hearts, diamonds, clubs; a thin double edge where the
/// diamonds start after a major; W, N and E face down when hidden.
fn paint_hand(ui: &egui::Ui, key: &Key, rect: Rect, deal: &Deal, hide: bool) {
    let seat = seat_of_box(key.index.unwrap_or(1));
    let p = ui.painter();
    let across = matches!(seat, Direction::North | Direction::South);
    let (step, edge) = if across {
        (Vec2::new(21.0, 0.0), Vec2::new(1.0, 0.0))
    } else {
        (Vec2::new(0.0, 27.0), Vec2::new(0.0, 1.0))
    };
    let mut after_major = false;
    for (i, c) in display_order(deal.hand(seat)).iter().enumerate() {
        let at = rect.min + step * i as f32;
        if hide && seat != Direction::South {
            cards::back(p, at);
            continue;
        }
        if c.suit == Suit::Diamonds && after_major {
            cards::face(p, at - edge * 2.0, *c);
            cards::face(p, at - edge, *c);
        }
        after_major = c.suit.is_major();
        cards::face(p, at, *c);
    }
}

fn seed_from_clock() -> u64 {
    web_time::SystemTime::now()
        .duration_since(web_time::UNIX_EPOCH)
        .map(|d| d.as_nanos() as u64)
        .unwrap_or(1)
}

/// Append to `SavedDeals.txt` in the user's BidBase folder.
#[cfg(not(target_arch = "wasm32"))]
fn save_deal_file(text: &str) -> Option<std::path::PathBuf> {
    use std::io::Write;
    let home = std::env::var_os("HOME").or_else(|| std::env::var_os("USERPROFILE"))?;
    let dir = std::path::Path::new(&home).join("BidBase");
    std::fs::create_dir_all(&dir).ok()?;
    let path = dir.join("SavedDeals.txt");
    let mut f = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
        .ok()?;
    f.write_all(text.as_bytes()).ok()?;
    Some(path)
}

const SETTINGS_KEY: &str = "bidbase-practice";

/// The window (native) or canvas (web) around the app.
pub struct Shell {
    app: App,
}

impl Shell {
    pub fn new(cc: &eframe::CreationContext<'_>) -> Self {
        bidbase_vbui::install(&cc.egui_ctx);
        let settings = cc
            .storage
            .and_then(|s| eframe::get_value::<Settings>(s, SETTINGS_KEY))
            .unwrap_or_default();
        Shell {
            app: App::with_settings(settings),
        }
    }
}

impl eframe::App for Shell {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default()
            .frame(egui::Frame::NONE)
            .show(ui, |ui| self.app.ui(ui));
        let ctx = ui.ctx().clone();
        if let Some(url) = self.app.open_url.take() {
            if url == "quit:" {
                ctx.send_viewport_cmd(egui::ViewportCommand::Close);
            } else {
                ctx.open_url(egui::OpenUrl::new_tab(url));
            }
        }
        if let Some(on) = self.app.on_top_changed.take() {
            let level = if on {
                egui::WindowLevel::AlwaysOnTop
            } else {
                egui::WindowLevel::Normal
            };
            ctx.send_viewport_cmd(egui::ViewportCommand::WindowLevel(level));
        }
    }

    fn save(&mut self, storage: &mut dyn eframe::Storage) {
        eframe::set_value(storage, SETTINGS_KEY, &self.app.settings);
    }
}
