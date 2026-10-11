//! The rule database: the `Bids` and `BidNames` tables, read from the CSV
//! export that `tools/export-db.sh` makes of `BidBase.accdb`.

use std::collections::HashMap;
use std::path::Path;

use serde::Deserialize;

use crate::vb::{ival, val};

/// One row of the `Bids` table: a call, the auction it applies to, and
/// the hand it needs. The conditions stay as the text Nelson typed; the
/// checks in `checks.rs` read them the way the VB6 engine did.
///
/// Columns missing from an export read as empty: the 2026 database
/// dropped `AgainstID`, `ValSystem`, `SuitVal`, `NT_Val`, `Strength_Discl`
/// and the disclosed lengths (`xLen`..`CLen`).
#[derive(Debug, Clone, Default, Deserialize, serde::Serialize)]
#[serde(default)]
pub struct Rule {
    #[serde(rename = "ID")]
    pub id: i64,
    #[serde(rename = "Pct_Used")]
    pub pct_used: String,
    #[serde(rename = "RoundOfBidding")]
    pub round: String,
    #[serde(rename = "Player")]
    pub player: String,
    #[serde(rename = "Prior_Bids")]
    pub prior_bids: String,
    #[serde(rename = "Bid_Order_Number")]
    pub bid_order: String,
    #[serde(rename = "Bid")]
    pub bid: String,
    #[serde(rename = "ConvID")]
    pub conv_id: String,
    #[serde(rename = "AgainstID")]
    pub against_id: String,
    #[serde(rename = "HCP")]
    pub hcp: String,
    #[serde(rename = "ValSystem")]
    pub val_system: String,
    #[serde(rename = "SuitVal")]
    pub suit_val: String,
    #[serde(rename = "NT_Val")]
    pub nt_val: String,
    #[serde(rename = "Cust_Pts")]
    pub cust_pts: String,
    #[serde(rename = "A_Cust")]
    pub a_cust: String,
    #[serde(rename = "K_Cust")]
    pub k_cust: String,
    #[serde(rename = "Q_Cust")]
    pub q_cust: String,
    #[serde(rename = "J_Cust")]
    pub j_cust: String,
    #[serde(rename = "T_Cust")]
    pub t_cust: String,
    #[serde(rename = "Dist_Pts")]
    pub dist_pts: String,
    #[serde(rename = "Winrs")]
    pub winners: String,
    #[serde(rename = "Losers")]
    pub losers: String,
    #[serde(rename = "QTs")]
    pub quick_tricks: String,
    #[serde(rename = "Stopped_Suits")]
    pub stopped_suits: String,
    #[serde(rename = "Shape")]
    pub shape: String,
    #[serde(rename = "Intermed")]
    pub intermed: String,
    #[serde(rename = "xNum")]
    pub x_num: String,
    #[serde(rename = "xHCP")]
    pub x_hcp: String,
    #[serde(rename = "sNum")]
    pub s_num: String,
    #[serde(rename = "sHCP")]
    pub s_hcp: String,
    #[serde(rename = "hNum")]
    pub h_num: String,
    #[serde(rename = "hHCP")]
    pub h_hcp: String,
    #[serde(rename = "dNum")]
    pub d_num: String,
    #[serde(rename = "dHCP")]
    pub d_hcp: String,
    #[serde(rename = "cNum")]
    pub c_num: String,
    #[serde(rename = "cHCP")]
    pub c_hcp: String,
    #[serde(rename = "cmNum")]
    pub cm_num: String,
    #[serde(rename = "cmHCP")]
    pub cm_hcp: String,
    #[serde(rename = "Vul")]
    pub vul: String,
    #[serde(rename = "Pos")]
    pub pos: String,
    #[serde(rename = "Artificial")]
    pub artificial: String,
    #[serde(rename = "Alert")]
    pub alert: String,
    #[serde(rename = "Explanation")]
    pub explanation: String,
    #[serde(rename = "Test_Hand")]
    pub test_hand: String,
    #[serde(rename = "Strength_Discl")]
    pub strength_discl: String,
    #[serde(rename = "xLen")]
    pub x_len: String,
    #[serde(rename = "SLen")]
    pub s_len: String,
    #[serde(rename = "HLen")]
    pub h_len: String,
    #[serde(rename = "DLen")]
    pub d_len: String,
    #[serde(rename = "CLen")]
    pub c_len: String,
    #[serde(rename = "Asks")]
    pub asks: String,
    #[serde(rename = "DateOfEntry")]
    pub date_of_entry: String,
    #[serde(rename = "Best_Contract")]
    pub best_contract: String,
    #[serde(rename = "Num_Used")]
    pub num_used: String,
    #[serde(rename = "Source")]
    pub source: String,
}

impl Rule {
    /// The suit-length spec for suit 1-4 (C, D, H, S).
    pub fn suit_num(&self, s: usize) -> &str {
        match s {
            1 => &self.c_num,
            2 => &self.d_num,
            3 => &self.h_num,
            4 => &self.s_num,
            _ => "",
        }
    }

    /// The suit-strength spec for suit 1-4 (C, D, H, S).
    pub fn suit_hcp(&self, s: usize) -> &str {
        match s {
            1 => &self.c_hcp,
            2 => &self.d_hcp,
            3 => &self.h_hcp,
            4 => &self.s_hcp,
            _ => "",
        }
    }

    /// `BidNames.IDnum` of the rule's convention, 0 when it has none.
    pub fn convention(&self) -> i64 {
        ival(&self.conv_id) as i64
    }

    /// The "ZZZZZZ" row that closes an auction's section.
    pub fn is_section_end(&self) -> bool {
        self.bid_order == "ZZZZZZ"
    }

    /// Whether the test hand is a whole hand ("AK2-AK2-AK2-K432").
    pub fn has_test_hand(&self) -> bool {
        self.test_hand.trim().len() == 16
    }

    /// Trump suit 1-4 named at the start of `Source`, as "<H...". Slam
    /// rules use it to say which suit their test hand is playing in.
    pub fn source_trump(&self) -> usize {
        match self.source.get(..2) {
            Some("<C") => 1,
            Some("<D") => 2,
            Some("<H") => 3,
            Some("<S") => 4,
            _ => 0,
        }
    }
}

/// One row of `BidNames`: a convention that can be switched on and off.
#[derive(Debug, Clone, Deserialize, serde::Serialize)]
pub struct Convention {
    #[serde(rename = "IDnum")]
    pub id: i64,
    /// "" = on, "0" = off, "W" = only for our side, "T" = only for theirs.
    #[serde(rename = "Selected")]
    pub selected: String,
    #[serde(rename = "SubCat_Selected")]
    pub sub_selected: String,
    #[serde(rename = "Bid_Name")]
    pub name: String,
    #[serde(rename = "Bid_Description")]
    pub description: String,
}

/// Which engine the rules were written for. The rules' meaning changed
/// between the 2018 release and Nelson's 2026 engine (docs/ENGINE-2026.md);
/// an export without the `Strength_Discl` column (dropped after 2018) is
/// taken as 2026.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Dialect {
    #[default]
    V2018,
    V2026,
}

/// HCP ranges of the notrump openings. Responses to them say "P24-25"
/// (partnership total) and the engine subtracts the opening's minimum.
#[derive(Debug, Clone, Copy, Default)]
pub struct NtRanges {
    pub one_min: i32,
    pub one_max: i32,
    pub two_min: i32,
    pub two_max: i32,
}

/// The whole rule set, indexed for lookup.
#[derive(Debug, Clone)]
pub struct RuleBase {
    pub rules: Vec<Rule>,
    pub conventions: HashMap<i64, Convention>,
    pub nt: NtRanges,
    pub dialect: Dialect,
    by_prior: HashMap<String, Vec<usize>>,
    by_id: HashMap<i64, usize>,
}

pub type LoadError = Box<dyn std::error::Error + Send + Sync>;

impl RuleBase {
    /// Read `Bids.csv` and `BidNames.csv` from an export directory.
    pub fn load(dir: &Path) -> Result<Self, LoadError> {
        let read = |name: &str| {
            let path = dir.join(name);
            std::fs::read_to_string(&path).map_err(|e| format!("{}: {e}", path.display()))
        };
        Self::from_csv(&read("Bids.csv")?, &read("BidNames.csv")?)
    }

    /// The rule set from the text of `Bids.csv` and `BidNames.csv`. An
    /// export without `Strength_Discl` is read as the 2026 dialect.
    pub fn from_csv(bids: &str, names: &str) -> Result<Self, LoadError> {
        let rules: Vec<Rule> = read_csv(bids, "Bids.csv")?;
        let names: Vec<Convention> = read_csv(names, "BidNames.csv")?;
        let header = csv::Reader::from_reader(bids.as_bytes()).headers()?.clone();
        let mut base = Self::new(rules, names);
        if !header.iter().any(|h| h == "Strength_Discl") {
            base.dialect = Dialect::V2026;
        }
        Ok(base)
    }

    pub fn new(rules: Vec<Rule>, names: Vec<Convention>) -> Self {
        let mut by_prior: HashMap<String, Vec<usize>> = HashMap::new();
        let mut by_id = HashMap::new();
        for (i, r) in rules.iter().enumerate() {
            by_prior.entry(r.prior_bids.clone()).or_default().push(i);
            by_id.insert(r.id, i);
        }
        // `ORDER BY Bid_Order_Number` on a text column: Access compares
        // text without case. Rows with equal numbers keep stored order.
        for list in by_prior.values_mut() {
            list.sort_by(|&a, &b| {
                let ka = rules[a].bid_order.to_ascii_uppercase();
                let kb = rules[b].bid_order.to_ascii_uppercase();
                ka.cmp(&kb).then(a.cmp(&b))
            });
        }
        let conventions = names
            .into_iter()
            .map(|mut c| {
                c.selected = c.selected.trim().to_string();
                c.sub_selected = c.sub_selected.trim().to_string();
                (c.id, c)
            })
            .collect();
        let mut base = RuleBase {
            rules,
            conventions,
            nt: NtRanges::default(),
            dialect: Dialect::default(),
            by_prior,
            by_id,
        };
        base.nt = base.find_nt_ranges();
        base
    }

    /// Switch every convention on. The `Selected` flags in a copy of the
    /// database are the owner's settings, not part of the rules: the
    /// 2018 copy has 1,424 rules in conventions switched off, including
    /// natural notrump responses and four-card 1-over-1 responses to 1C.
    pub fn switch_all_on(&mut self) {
        for c in self.conventions.values_mut() {
            c.selected.clear();
            c.sub_selected.clear();
        }
    }

    /// Rules for an auction (prior bids without opening passes), in
    /// priority order.
    pub fn for_auction(&self, prior_bids: &str) -> &[usize] {
        self.by_prior
            .get(prior_bids)
            .map(Vec::as_slice)
            .unwrap_or(&[])
    }

    pub fn by_id(&self, id: i64) -> Option<&Rule> {
        self.by_id.get(&id).map(|&i| &self.rules[i])
    }

    pub fn index_of(&self, id: i64) -> Option<usize> {
        self.by_id.get(&id).copied()
    }

    /// Distinct auctions the rules cover.
    pub fn auctions(&self) -> impl Iterator<Item = &str> {
        self.by_prior.keys().map(String::as_str)
    }

    /// The 2018 engine read these from fixed `Bid_Order_Number`s
    /// ("130000" for 1NT, "110000" for 2NT), which by then pointed at
    /// other rules. Find the opening rules themselves: the first enabled
    /// 1NT and 2NT openings with a plain "lo-hi" range.
    fn find_nt_ranges(&self) -> NtRanges {
        let range = |bid: &str| {
            self.for_auction("")
                .iter()
                .map(|&i| &self.rules[i])
                .filter(|r| r.bid == bid && r.pct_used.trim().is_empty())
                .find_map(|r| {
                    let (lo, hi) = r.hcp.trim().split_once('-')?;
                    Some((lo.parse().ok()?, hi.parse().ok()?))
                })
                .unwrap_or((0, 0))
        };
        let (one_min, one_max) = range("1N");
        let (two_min, two_max) = range("2N");
        NtRanges {
            one_min,
            one_max,
            two_min,
            two_max,
        }
    }

    /// Whether a convention allows the rule for a bidder at `pos_num`
    /// (1-4, counted from the opener: 1 and 3 are the opening side).
    /// Returns the reason when it doesn't.
    pub fn convention_allows(&self, rule: &Rule, pos_num: i32) -> Result<(), String> {
        let Some(c) = self.conventions.get(&rule.convention()) else {
            return Ok(());
        };
        let ours = pos_num == 1 || pos_num == 3;
        let allows = |sel: &str| sel.is_empty() || (sel == "W" && ours) || (sel == "T" && !ours);
        if !allows(&c.selected) {
            return Err(format!(
                "convention {} \"{}\" is off ({:?})",
                c.id, c.name, c.selected
            ));
        }
        if !allows(&c.sub_selected) {
            return Err(format!(
                "convention {} \"{}\" sub-option is off ({:?})",
                c.id, c.name, c.sub_selected
            ));
        }
        Ok(())
    }
}

fn read_csv<T: for<'de> Deserialize<'de>>(text: &str, name: &str) -> Result<Vec<T>, LoadError> {
    let mut reader = csv::Reader::from_reader(text.as_bytes());
    let mut out = Vec::new();
    for row in reader.deserialize() {
        out.push(row.map_err(|e| format!("{name}: {e}"))?);
    }
    Ok(out)
}

/// Read `Pct_Used` as a percentage (0 when it is blank, a letter or "0").
pub fn pct_number(pct: &str) -> f64 {
    val(pct)
}
