//! `bidbase`: ask the BidBase port for a call, look at a rule, or check
//! every rule against its own test hand.

mod bid_pbn;

use std::collections::BTreeMap;
use std::path::PathBuf;
use std::process::ExitCode;

use bidbase_engine::card::{Applied, CardMap, Switch};
use bidbase_engine::selftest::{check_rule, Verdict};
use bidbase_engine::{Bidder, Request, Rule, RuleBase, Scoring, Vul};
use bridge_card::Card;
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
    /// Set the convention switches from a convention card (convention-card's
    /// native JSON), through card/bidbase-map.toml.
    #[arg(long, global = true)]
    card: Option<PathBuf>,
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
    /// Bid every deal of a PBN file at all four seats; write the auctions
    /// in bba-cli's layout. The card (--card) plays both sides.
    BidPbn {
        #[arg(short, long)]
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
        /// Seed for rules used a percentage of the time.
        #[arg(long, default_value_t = 1)]
        seed: u64,
    },
    /// Walk the auctions of another engine's PBN (bba-cli layout) and ask
    /// BidBase for every call in the same position: how often its rules
    /// fire, and how often they agree. `-v` lists every disagreement.
    Replay {
        #[arg(short, long)]
        input: PathBuf,
        #[arg(short, long)]
        verbose: bool,
    },
    /// Show which conventions a card switches on and off (--card), or with
    /// no card, what the map covers.
    Card {
        /// List each convention, not just the totals.
        #[arg(short, long)]
        verbose: bool,
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
    let applied = match &cli.card {
        Some(path) => match apply_card(path, &mut base) {
            Ok(a) => Some(a),
            Err(e) => {
                eprintln!("bidbase: {e}");
                return ExitCode::FAILURE;
            }
        },
        None => None,
    };
    match cli.command {
        Command::Card { verbose } => show_card(&base, applied.as_ref(), verbose),
        Command::Replay { input, verbose } => run_replay(base, &input, verbose),
        Command::BidPbn {
            input,
            output,
            seed,
        } => {
            let card = cli
                .card
                .as_ref()
                .and_then(|p| p.file_stem())
                .map_or("BidBase database switches".to_string(), |s| {
                    s.to_string_lossy().into_owned()
                });
            run_bid_pbn(base, &input, &output, seed, &card)
        }
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
    let hand = match bidbase_engine::convert::normalize_hand(hand) {
        Some(h) => h,
        None => {
            eprintln!("bidbase: cannot read hand {hand:?}");
            return ExitCode::FAILURE;
        }
    };
    let Some(prior_calls) = bidbase_engine::convert::normalize_auction(auction) else {
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
            println!(
                "{}  (rule #{}) {}",
                found.call, found.rule_id, found.meaning
            );
            if disclose {
                println!("\n{} can show (every enabled rule for it):", found.call);
                for (id, m) in &found.meanings {
                    println!("  #{id:<6} {m}");
                }
            }
            ExitCode::SUCCESS
        }
        Err(e) => {
            println!("no call: {e} (error {})", e.code());
            ExitCode::from(2)
        }
    }
}

fn run_replay(base: RuleBase, input: &std::path::Path, verbose: bool) -> ExitCode {
    let text = match std::fs::read_to_string(input) {
        Ok(t) => t,
        Err(e) => {
            eprintln!("bidbase: {}: {e}", input.display());
            return ExitCode::FAILURE;
        }
    };
    let mut bidder = Bidder::new(base);
    let calls = match bid_pbn::replay(&text, &mut bidder) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("bidbase: {e}");
            return ExitCode::FAILURE;
        }
    };
    // Split by whether the reference call is a pass, since "no rule fits"
    // is BidBase passing.
    #[derive(Default)]
    struct Tally {
        calls: usize,
        unknown: usize,
        no_fit: usize,
        fired: usize,
        agree: usize,
    }
    let mut by_kind: BTreeMap<&str, Tally> = BTreeMap::new();
    for c in &calls {
        let kind = if c.theirs == "P" {
            "their pass"
        } else {
            "their bid/double"
        };
        for key in [kind, "all"] {
            let t = by_kind.entry(key).or_default();
            t.calls += 1;
            match &c.ours {
                Some(ours) => {
                    t.fired += 1;
                    t.agree += usize::from(*ours == c.theirs);
                }
                None if c.known => t.no_fit += 1,
                None => t.unknown += 1,
            }
        }
        if verbose {
            if let Some(ours) = c.ours.as_ref().filter(|o| **o != c.theirs) {
                println!(
                    "board {:>3} {:<28} theirs {:<3} ours {:<3} (#{})",
                    c.board,
                    format!("{:?}", c.prior),
                    c.theirs,
                    ours,
                    c.rule_id.unwrap_or(0)
                );
            }
        }
    }
    let pct = |a: usize, b: usize| 100.0 * a as f64 / b.max(1) as f64;
    println!(
        "{:<17} {:>6} {:>15} {:>14} {:>12} {:>16}",
        "", "calls", "no rules (seq)", "rules, no fit", "rule fired", "fired and agree"
    );
    for (kind, t) in by_kind.iter().rev() {
        println!(
            "{kind:<17} {:>6} {:>14.0}% {:>13.0}% {:>11.0}% {:>9} = {:.0}%",
            t.calls,
            pct(t.unknown, t.calls),
            pct(t.no_fit, t.calls),
            pct(t.fired, t.calls),
            t.agree,
            pct(t.agree, t.fired)
        );
    }
    ExitCode::SUCCESS
}

fn run_bid_pbn(
    base: RuleBase,
    input: &std::path::Path,
    output: &std::path::Path,
    seed: u64,
    card: &str,
) -> ExitCode {
    let text = match std::fs::read_to_string(input) {
        Ok(t) => t,
        Err(e) => {
            eprintln!("bidbase: {}: {e}", input.display());
            return ExitCode::FAILURE;
        }
    };
    let mut bidder = Bidder::new(base);
    bidder.seed(seed);
    let started = std::time::Instant::now();
    match bid_pbn::bid_pbn(&text, &mut bidder, card, card) {
        Ok((out, st)) => {
            if let Err(e) = std::fs::write(output, out) {
                eprintln!("bidbase: {}: {e}", output.display());
                return ExitCode::FAILURE;
            }
            eprintln!(
                "{} deals, {} calls in {:.1}s: {} with no rule (passed; {} in sequences with \
                 no rules), {} illegal (passed), {} 4NT replies by ace-asking key, {} runaway",
                st.deals,
                st.calls,
                started.elapsed().as_secs_f64(),
                st.no_rule,
                st.unknown_auction,
                st.illegal,
                st.slam_keys,
                st.runaway
            );
            ExitCode::SUCCESS
        }
        Err(e) => {
            eprintln!("bidbase: {e}");
            ExitCode::FAILURE
        }
    }
}

fn apply_card(path: &std::path::Path, base: &mut RuleBase) -> Result<Applied, String> {
    let (map, vocab) = CardMap::standard()?;
    let text = std::fs::read_to_string(path).map_err(|e| format!("{}: {e}", path.display()))?;
    let (card, report) =
        Card::from_json(&vocab, &text).map_err(|e| format!("{}: {e}", path.display()))?;
    if !report.is_clean() {
        eprintln!("bidbase: {}: {report:?}", path.display());
    }
    Ok(map.apply(&card, base))
}

/// With a card: the switches it set. Without: how much of the database the
/// map covers, by rules.
fn show_card(base: &RuleBase, applied: Option<&Applied>, verbose: bool) -> ExitCode {
    let mut rules_of: BTreeMap<i64, usize> = BTreeMap::new();
    for r in base.rules.iter().filter(|r| !r.is_section_end()) {
        *rules_of.entry(r.convention()).or_default() += 1;
    }
    let count = |ids: &[i64]| -> (usize, usize) {
        let used: Vec<_> = ids.iter().filter(|id| rules_of.contains_key(id)).collect();
        (used.len(), used.iter().map(|id| rules_of[id]).sum())
    };
    let line = |label: &str, ids: &[i64]| {
        let (n, rules) = count(ids);
        println!("{label:<34} {n:>4} conventions, {rules:>5} rules");
        if verbose {
            for id in ids.iter().filter(|id| rules_of.contains_key(id)) {
                let name = base.conventions.get(id).map_or("", |c| c.name.as_str());
                println!("    {id:>4} {:>4}  {name}", rules_of[id]);
            }
        }
    };
    match applied {
        Some(a) => {
            line("on (card or always)", &a.on);
            line("off (card)", &a.off);
            line("not in the map: database setting", &a.kept);
        }
        None => {
            let Ok((map, _)) = CardMap::standard() else {
                eprintln!("bidbase: the card map doesn't load");
                return ExitCode::FAILURE;
            };
            let ids: Vec<i64> = rules_of.keys().copied().filter(|id| *id != 0).collect();
            let (always, carded): (Vec<i64>, Vec<i64>) = ids
                .iter()
                .filter(|id| map.entries.contains_key(id))
                .partition(|id| map.entries[id].switch == Switch::Always);
            let unmapped: Vec<i64> = ids
                .iter()
                .copied()
                .filter(|id| !map.entries.contains_key(id))
                .collect();
            line("set by card fields", &carded);
            line("always on (BidBase categories)", &always);
            line("not in the map", &unmapped);
        }
    }
    ExitCode::SUCCESS
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
