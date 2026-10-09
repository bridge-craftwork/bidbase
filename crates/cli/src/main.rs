//! `bidbase`: ask the BidBase port for a call, look at a rule, or check
//! every rule against its own test hand.

use std::collections::BTreeMap;
use std::path::PathBuf;
use std::process::ExitCode;

use bidbase_engine::selftest::{check_rule, Verdict};
use bidbase_engine::{Bidder, Request, Rule, RuleBase, Scoring, Vul};
use bridge_types::{Call, Hand};
use clap::{Parser, Subcommand};

#[derive(Parser)]
#[command(
    name = "bidbase",
    version,
    about = "Rust port of the CardShark BidBase bidding engine"
)]
struct Cli {
    /// Directory with the CSV export of BidBase.accdb (Bids.csv, BidNames.csv).
    #[arg(long, global = true, env = "BIDBASE_DATA")]
    data: Option<PathBuf>,
    /// Ignore the database's convention switches (BidNames.Selected) and
    /// use every convention.
    #[arg(long, global = true)]
    all_conventions: bool,
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// Find the call for a hand.
    Bid {
        /// The hand: "AK432-64-QJT3-T2" or PBN "AK432.64.QJT3.T2" (spades first).
        hand: String,
        /// Calls so far, dealer first: "P-1C-1N" or "Pass 1C 1NT" (X, XX, D, R).
        #[arg(short, long, default_value = "")]
        auction: String,
        /// Vulnerability from the bidder's side: none, we, they, both.
        #[arg(short, long, default_value = "none")]
        vul: String,
        /// Scoring: mp or imp (rules marked M or I apply only then).
        #[arg(short, long)]
        scoring: Option<String>,
        /// Trump suit (C, D, H, S) for ace-asking replies.
        #[arg(short, long)]
        trump: Option<char>,
        /// Seed for rules used a percentage of the time.
        #[arg(long, default_value_t = 1)]
        seed: u64,
        /// List every rule tried and why it was passed over.
        #[arg(short, long)]
        why: bool,
        /// Print what the call discloses to the table.
        #[arg(short, long)]
        disclose: bool,
    },
    /// Show a rule's fields.
    Rule {
        /// The rule's ID.
        id: i64,
    },
    /// Check every rule against its test hand, as the BidBase Editor does
    /// when an entry is saved.
    Selftest {
        /// List the rules that fail (own: reject their test hand; first:
        /// an earlier rule takes it; all: both).
        #[arg(short, long)]
        list: Option<String>,
    },
}

fn main() -> ExitCode {
    let cli = Cli::parse();
    let dir = cli
        .data
        .unwrap_or_else(|| PathBuf::from(concat!(env!("CARGO_MANIFEST_DIR"), "/../../data/2018")));
    let mut base = match RuleBase::load(&dir) {
        Ok(b) => b,
        Err(e) => {
            eprintln!("bidbase: cannot read rules from {}: {e}", dir.display());
            return ExitCode::FAILURE;
        }
    };
    if cli.all_conventions {
        base.switch_all_on();
    }
    match cli.command {
        Command::Bid {
            hand,
            auction,
            vul,
            scoring,
            trump,
            seed,
            why,
            disclose,
        } => bid(
            base,
            &hand,
            &auction,
            &vul,
            scoring.as_deref(),
            trump,
            seed,
            why,
            disclose,
        ),
        Command::Rule { id } => show_rule(&base, id),
        Command::Selftest { list } => selftest(base, list.as_deref()),
    }
}

#[allow(clippy::too_many_arguments)]
fn bid(
    base: RuleBase,
    hand: &str,
    auction: &str,
    vul: &str,
    scoring: Option<&str>,
    trump: Option<char>,
    seed: u64,
    why: bool,
    disclose: bool,
) -> ExitCode {
    let hand = match normalize_hand(hand) {
        Some(h) => h,
        None => {
            eprintln!("bidbase: cannot read hand {hand:?}");
            return ExitCode::FAILURE;
        }
    };
    let Some(prior_calls) = normalize_auction(auction) else {
        eprintln!("bidbase: cannot read auction {auction:?}");
        return ExitCode::FAILURE;
    };
    let Some(vul) = Vul::parse(vul) else {
        eprintln!("bidbase: unknown vulnerability {vul:?} (none, we, they, both)");
        return ExitCode::FAILURE;
    };
    let scoring = match scoring.map(str::to_ascii_lowercase).as_deref() {
        None => Scoring::Unspecified,
        Some("mp" | "m" | "matchpoints") => Scoring::Matchpoints,
        Some("imp" | "imps" | "i") => Scoring::Imps,
        Some(s) => {
            eprintln!("bidbase: unknown scoring {s:?} (mp, imp)");
            return ExitCode::FAILURE;
        }
    };
    let trump = trump.map_or(0, |c| {
        bidbase_engine::vb::suit_index(c.to_ascii_uppercase())
    });
    let mut bidder = Bidder::new(base);
    bidder.seed(seed);
    let out = bidder.lookup(&Request {
        hand,
        prior_calls,
        vul,
        scoring,
        trump,
        disclose,
        trace: why,
        ..Default::default()
    });
    for t in &out.tried {
        match &t.failed {
            Some((test, reason)) => {
                println!(
                    "  #{:<6} {:<7} {:<3} {test}: {reason}",
                    t.rule_id, t.bid_order, t.bid
                )
            }
            None => println!(
                "  #{:<6} {:<7} {:<3} matches",
                t.rule_id, t.bid_order, t.bid
            ),
        }
    }
    match out.result {
        Ok(found) => {
            println!("{}  (rule #{})", found.call, found.rule_id);
            if disclose && !found.disclosure.is_empty() {
                println!("\n{}", found.disclosure.trim_end());
            }
            ExitCode::SUCCESS
        }
        Err(e) => {
            println!("no call: {e} (error {})", e.code());
            ExitCode::from(2)
        }
    }
}

/// BidBase "S-H-D-C" with dashes, or PBN with dots.
fn normalize_hand(hand: &str) -> Option<String> {
    let hand = hand.trim().to_ascii_uppercase().replace("10", "T");
    if hand.contains('.') {
        let parsed = Hand::from_pbn(&hand)?;
        return Some(bidbase_engine::convert::hand_to_bidbase(&parsed));
    }
    Some(hand)
}

/// "P-1C-1N", "Pass 1C 1NT", "p 1c x": BidBase tokens joined by dashes.
fn normalize_auction(auction: &str) -> Option<String> {
    let tokens: Vec<&str> = auction
        .split(|c: char| c == '-' || c.is_whitespace() || c == ',')
        .filter(|t| !t.is_empty())
        .collect();
    let mut out = Vec::new();
    for t in tokens {
        let up = t.to_ascii_uppercase();
        let token = match up.as_str() {
            "P" | "PASS" => "P".to_string(),
            "D" | "X" | "DBL" => "D".to_string(),
            "R" | "XX" | "RDBL" => "R".to_string(),
            _ => {
                let call =
                    Call::from_pbn(&up).or_else(|| Call::from_pbn(&up.replace('N', "NT")))?;
                bidbase_engine::convert::call_to_bidbase(&call)
            }
        };
        out.push(token);
    }
    Some(out.join("-"))
}

fn show_rule(base: &RuleBase, id: i64) -> ExitCode {
    let Some(rule) = base.by_id(id) else {
        eprintln!("bidbase: no rule #{id}");
        return ExitCode::FAILURE;
    };
    let text = format!("{rule:#?}");
    for line in text.lines() {
        if !line.ends_with(": \"\",") {
            println!("{line}");
        }
    }
    if let Some(c) = base.conventions.get(&rule.convention()) {
        println!(
            "convention #{}: {:?} selected={:?} sub={:?}",
            c.id, c.name, c.selected, c.sub_selected
        );
    }
    ExitCode::SUCCESS
}

fn selftest(base: RuleBase, list: Option<&str>) -> ExitCode {
    let rules: Vec<Rule> = base
        .rules
        .iter()
        .filter(|r| !r.is_section_end() && r.has_test_hand())
        .cloned()
        .collect();
    let total = base.rules.iter().filter(|r| !r.is_section_end()).count();
    let mut bidder = Bidder::new(base);
    let (mut ok, mut own, mut shadowed, mut off, mut bad, mut errors) = (0, 0, 0, 0, 0, 0);
    let mut own_by_test: BTreeMap<String, usize> = BTreeMap::new();
    let show_own = matches!(list, Some("own" | "all"));
    let show_first = matches!(list, Some("first" | "all"));
    for rule in &rules {
        let head = format!(
            "#{:<6} {:<22} {:<3} {:<17}",
            rule.id,
            format!("{:?}", rule.prior_bids),
            rule.bid,
            rule.test_hand.trim()
        );
        match check_rule(&mut bidder, rule) {
            Verdict::Ok => ok += 1,
            Verdict::OwnFails(why) => {
                own += 1;
                let test = why.split(':').next().unwrap_or("?").to_string();
                *own_by_test.entry(test).or_default() += 1;
                if show_own {
                    println!("own   {head} {why}");
                }
            }
            Verdict::Shadowed(by, call) => {
                shadowed += 1;
                if show_first {
                    println!("first {head} taken by #{by} ({call})");
                }
            }
            Verdict::Off => off += 1,
            Verdict::BadHand(e) => {
                bad += 1;
                if list.is_some() {
                    println!("hand  {head} {e}");
                }
            }
            Verdict::Error(e) => {
                errors += 1;
                if list.is_some() {
                    println!("error {head} {e}");
                }
            }
        }
    }
    println!(
        "rules: {total}, with test hands: {}\n  first match is the rule itself: {ok}\n  \
         switched off, accepts its test hand: {off}\n  \
         an earlier rule takes the hand: {shadowed}\n  rule rejects its own test hand: {own}\n  \
         test hand unreadable: {bad}\n  errors: {errors}",
        rules.len()
    );
    for (test, n) in &own_by_test {
        println!("    rejected by {test}: {n}");
    }
    ExitCode::SUCCESS
}
