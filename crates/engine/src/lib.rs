//! A Rust port of the bidding engine in Nelson Ford's CardShark BidBase
//! (`BidBaseDB_Bidder.cls`, VB6, September 2018 SourceForge release).
//!
//! BidBase keeps its bidding knowledge in a database of rules (`Bids`):
//! each names the auction so far, a call, and what the hand needs. The
//! engine analyses the hand once, then tries the auction's rules in
//! `Bid_Order_Number` order and makes the call of the first rule the hand
//! passes. This crate reads the rules from the CSV export of the
//! database (`data/2018`, made by `tools/export-db.sh`).
//!
//! ```no_run
//! use bidbase_engine::{Bidder, Request, RuleBase, Vul};
//! let base = RuleBase::load(std::path::Path::new("data/2018")).unwrap();
//! let mut bidder = Bidder::new(base);
//! let out = bidder.lookup(&Request {
//!     hand: "AK2-AK2-AK2-K432".into(),
//!     vul: Vul::Neither,
//!     ..Default::default()
//! });
//! println!("{:?}", out.result.map(|f| f.call));
//! ```

pub mod auction;
pub mod card;
pub mod checks;
pub mod convert;
pub mod hand;
pub mod lookup;
pub mod rules;
pub mod selftest;
pub mod vb;

pub use auction::AuctionFacts;
pub use hand::{HandError, HandFacts};
pub use lookup::{Bidder, Found, LookupError, Outcome, Request, Scoring, Tried, Vul};
pub use rules::{Convention, Rule, RuleBase};
