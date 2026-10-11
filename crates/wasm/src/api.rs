//! The JSON API behind the browser build. Every function takes a JSON
//! request string and returns a JSON response string, `{"ok": true, ...}`
//! or `{"ok": false, "error": "..."}`. Plain Rust, tested natively.
//!
//! An engine is a rule set with its convention switches; `createEngine`
//! returns an id the other calls take. The rule sets compiled in are
//! Nelson's 2026 rules ("2026", the default) and the 2018 SourceForge
//! release ("2018"); a request may instead pass `{"bids": ..., "bidnames":
//! ...}` with the CSV text of an export.

use std::cell::RefCell;
use std::collections::HashMap;

use bidbase_engine::card::CardMap;
use bidbase_engine::convert::{normalize_auction, normalize_hand};
use bidbase_engine::rules::Dialect;
use bidbase_engine::table::{self, Missing};
use bidbase_engine::{meaning, AuctionFacts, Bidder, Request, RuleBase, Scoring, Vul};
use bridge_types::{Direction, Hand, Vulnerability};
use serde_json::{json, Value};

const BIDS_2026: &str = include_str!("../../../data/2026/Bids.csv");
const NAMES_2026: &str = include_str!("../../../data/2026/BidNames.csv");
const BIDS_2018: &str = include_str!("../../../data/2018/Bids.csv");
const NAMES_2018: &str = include_str!("../../../data/2018/BidNames.csv");

/// The API's version; bump it when a request or response changes shape.
pub const API: u32 = 1;

struct Engine {
    bidder: Bidder,
    source: String,
}

thread_local! {
    static ENGINES: RefCell<(u32, HashMap<u32, Engine>)> = RefCell::new((0, HashMap::new()));
}

fn ok(mut v: Value) -> String {
    v["ok"] = Value::Bool(true);
    v.to_string()
}

fn err(msg: impl std::fmt::Display) -> String {
    json!({"ok": false, "error": msg.to_string()}).to_string()
}

fn parse(request: &str) -> Result<Value, String> {
    serde_json::from_str(request).map_err(|e| format!("request is not JSON: {e}"))
}

fn str_field<'a>(req: &'a Value, name: &str) -> &'a str {
    req.get(name).and_then(Value::as_str).unwrap_or("")
}

fn with_engine<T>(
    req: &Value,
    f: impl FnOnce(&mut Engine) -> Result<T, String>,
) -> Result<T, String> {
    let id = req
        .get("engine")
        .and_then(Value::as_u64)
        .ok_or("no \"engine\" id")? as u32;
    ENGINES.with(|e| {
        let mut e = e.borrow_mut();
        let engine = e.1.get_mut(&id).ok_or(format!("no engine {id}"))?;
        f(engine)
    })
}

/// `{}` → `{ok, api, version, rule_sets}`
pub fn info() -> String {
    ok(json!({
        "api": API,
        "version": env!("CARGO_PKG_VERSION"),
        "rule_sets": ["2026", "2018"],
    }))
}

/// `{rules?: "2026"|"2018"|{bids, bidnames}, card?: <card JSON>,
/// all_conventions?: bool, seed?: n}` → `{ok, engine, source, dialect,
/// rules, conventions, nt, card?}`
pub fn create_engine(request: &str) -> String {
    let req = match parse(request) {
        Ok(r) => r,
        Err(e) => return err(e),
    };
    let (source, loaded) = match req.get("rules") {
        None | Some(Value::Null) => (
            "2026".to_string(),
            RuleBase::from_csv(BIDS_2026, NAMES_2026),
        ),
        Some(Value::String(s)) if s == "2026" => {
            ("2026".into(), RuleBase::from_csv(BIDS_2026, NAMES_2026))
        }
        Some(Value::String(s)) if s == "2018" => {
            ("2018".into(), RuleBase::from_csv(BIDS_2018, NAMES_2018))
        }
        Some(Value::Object(o)) => {
            let bids = o.get("bids").and_then(Value::as_str).unwrap_or("");
            let names = o.get("bidnames").and_then(Value::as_str).unwrap_or("");
            ("custom".into(), RuleBase::from_csv(bids, names))
        }
        Some(other) => return err(format!("unknown rule set {other}")),
    };
    let mut base = match loaded {
        Ok(b) => b,
        Err(e) => return err(format!("cannot read the rules: {e}")),
    };
    if req.get("all_conventions").and_then(Value::as_bool) == Some(true) {
        base.switch_all_on();
    }
    let mut card_report = Value::Null;
    if let Some(card) = req.get("card").filter(|c| !c.is_null()) {
        let text = match card {
            Value::String(s) => s.clone(),
            other => other.to_string(),
        };
        let (map, vocab) = match CardMap::standard() {
            Ok(m) => m,
            Err(e) => return err(format!("card map: {e}")),
        };
        let (card, _) = match bridge_card::Card::from_json(&vocab, &text) {
            Ok(c) => c,
            Err(e) => return err(format!("card: {e}")),
        };
        let applied = map.apply(&card, &mut base);
        card_report =
            json!({"on": applied.on.len(), "off": applied.off.len(), "kept": applied.kept.len()});
    }
    let summary = json!({
        "source": source,
        "dialect": match base.dialect { Dialect::V2018 => "2018", Dialect::V2026 => "2026" },
        "rules": base.rules.iter().filter(|r| !r.is_section_end()).count(),
        "conventions": base.conventions.len(),
        "nt": {"one": [base.nt.one_min, base.nt.one_max], "two": [base.nt.two_min, base.nt.two_max]},
        "card": card_report,
    });
    let mut bidder = Bidder::new(base);
    bidder.seed(req.get("seed").and_then(Value::as_u64).unwrap_or(1));
    let id = ENGINES.with(|e| {
        let mut e = e.borrow_mut();
        e.0 += 1;
        let id = e.0;
        e.1.insert(id, Engine { bidder, source });
        id
    });
    let mut out = summary;
    out["engine"] = json!(id);
    ok(out)
}

/// `{engine}` → `{ok}`
pub fn free_engine(request: &str) -> String {
    let req = match parse(request) {
        Ok(r) => r,
        Err(e) => return err(e),
    };
    let id = req.get("engine").and_then(Value::as_u64).unwrap_or(0) as u32;
    ENGINES.with(|e| e.borrow_mut().1.remove(&id));
    ok(json!({}))
}

fn vul_of(s: &str) -> Result<Vul, String> {
    Vul::parse(s).ok_or(format!(
        "unknown vulnerability {s:?} (none, we, they, both)"
    ))
}

fn trump_of(s: &str) -> usize {
    s.chars().next().map_or(0, |c| {
        bidbase_engine::vb::suit_index(c.to_ascii_uppercase())
    })
}

/// `{engine, hand, auction?, vul?, trump?, scoring?}` → `{ok, call, rule,
/// meaning, meanings: [{rule, meaning}], tried: [{rule, bid, bon, test?,
/// why?}], error?}`. With no call, `error` says why (no rules for the
/// auction, or none fits the hand) and `tried` still lists the rules.
pub fn bid(request: &str) -> String {
    let req = match parse(request) {
        Ok(r) => r,
        Err(e) => return err(e),
    };
    let result = with_engine(&req, |engine| {
        let hand = normalize_hand(str_field(&req, "hand")).ok_or("cannot read the hand")?;
        let prior =
            normalize_auction(str_field(&req, "auction")).ok_or("cannot read the auction")?;
        let vul = vul_of(req.get("vul").and_then(Value::as_str).unwrap_or("none"))?;
        let scoring = match str_field(&req, "scoring").to_ascii_lowercase().as_str() {
            "mp" | "m" => Scoring::Matchpoints,
            "imp" | "i" => Scoring::Imps,
            _ => Scoring::Unspecified,
        };
        let out = engine.bidder.lookup(&Request {
            hand,
            prior_calls: prior.clone(),
            vul,
            scoring,
            trump: trump_of(str_field(&req, "trump")),
            disclose: true,
            trace: true,
            ..Default::default()
        });
        let tried: Vec<Value> = out
            .tried
            .iter()
            .map(|t| match &t.failed {
                Some((test, why)) => json!({"rule": t.rule_id, "bid": t.bid, "bon": t.bid_order, "test": test, "why": why}),
                None => json!({"rule": t.rule_id, "bid": t.bid, "bon": t.bid_order}),
            })
            .collect();
        let key = AuctionFacts::new(&prior).prior_bids;
        Ok(match out.result {
            Ok(f) => json!({
                "call": f.call,
                "rule": f.rule_id,
                "meaning": f.meaning,
                "meanings": f.meanings.iter().map(|(id, m)| json!({"rule": id, "meaning": m})).collect::<Vec<_>>(),
                "key": key,
                "tried": tried,
            }),
            Err(e) => {
                json!({"call": null, "error": e.to_string(), "code": e.code(), "key": key, "tried": tried})
            }
        })
    });
    match result {
        Ok(v) => ok(v),
        Err(e) => err(e),
    }
}

/// `{engine, id}` → `{ok, rule: {column: value...}, convention?, meaning}`;
/// the rule's non-empty fields under the database's column names.
pub fn rule(request: &str) -> String {
    let req = match parse(request) {
        Ok(r) => r,
        Err(e) => return err(e),
    };
    let id = req.get("id").and_then(Value::as_i64).unwrap_or(0);
    let result = with_engine(&req, |engine| {
        let base = &engine.bidder.base;
        let r = base.by_id(id).ok_or(format!("no rule #{id}"))?;
        let mut fields = serde_json::to_value(r).map_err(|e| e.to_string())?;
        if let Value::Object(map) = &mut fields {
            map.retain(|_, v| !matches!(v, Value::String(s) if s.is_empty()));
        }
        let conv = base.conventions.get(&r.convention()).map(|c| {
            json!({"id": c.id, "name": c.name, "selected": c.selected, "sub_selected": c.sub_selected})
        });
        Ok(json!({
            "rule": fields,
            "convention": conv,
            "meaning": meaning::describe(base, r, &r.prior_bids, 0),
        }))
    });
    match result {
        Ok(v) => ok(v),
        Err(e) => err(e),
    }
}

/// `{engine, auction}` → `{ok, key, rules: [{id, bon, bid, enabled,
/// convention, meaning, test_hand}]}`: the auction's section in priority
/// order, as the Editor's grid shows it.
pub fn section(request: &str) -> String {
    let req = match parse(request) {
        Ok(r) => r,
        Err(e) => return err(e),
    };
    let result = with_engine(&req, |engine| {
        let raw = str_field(&req, "auction");
        // Keys like "4N@30-P" or "RKCB: 4N-P" are taken as written.
        let prior = if raw.contains('@') || raw.contains(':') {
            raw.trim().to_string()
        } else {
            normalize_auction(raw).ok_or("cannot read the auction")?
        };
        let facts = AuctionFacts::new(&prior);
        let key = if raw.contains(':') {
            raw.trim().to_string()
        } else {
            facts.prior_bids.clone()
        };
        let bidder = &engine.bidder;
        let base = &bidder.base;
        let rules: Vec<Value> = base
            .for_auction(&key)
            .iter()
            .map(|&i| &base.rules[i])
            .take_while(|r| !r.is_section_end())
            .map(|r| {
                let conv = base
                    .conventions
                    .get(&r.convention())
                    .map(|c| c.name.trim().to_string());
                json!({
                    "id": r.id,
                    "bon": r.bid_order,
                    "bid": r.bid,
                    "enabled": bidder.is_enabled(r, facts.pos_num),
                    "convention": conv,
                    "meaning": meaning::describe(base, r, &key, 0),
                    "test_hand": r.test_hand.trim(),
                })
            })
            .collect();
        Ok(json!({"key": key, "rules": rules}))
    });
    match result {
        Ok(v) => ok(v),
        Err(e) => err(e),
    }
}

/// `{engine}` → `{ok, conventions: [{id, name, on, selected, sub_selected,
/// rules}]}`, the conventions some rule uses, by name.
pub fn conventions(request: &str) -> String {
    let req = match parse(request) {
        Ok(r) => r,
        Err(e) => return err(e),
    };
    let result = with_engine(&req, |engine| {
        let base = &engine.bidder.base;
        let mut count: HashMap<i64, usize> = HashMap::new();
        for r in base.rules.iter().filter(|r| !r.is_section_end()) {
            *count.entry(r.convention()).or_default() += 1;
        }
        let mut list: Vec<Value> = base
            .conventions
            .values()
            .filter(|c| count.contains_key(&c.id))
            .map(|c| {
                json!({
                    "id": c.id,
                    "name": c.name.trim(),
                    "on": c.selected.is_empty() && c.sub_selected.is_empty(),
                    "selected": c.selected,
                    "sub_selected": c.sub_selected,
                    "rules": count[&c.id],
                })
            })
            .collect();
        list.sort_by(|a, b| a["name"].as_str().cmp(&b["name"].as_str()));
        Ok(json!({"source": engine.source, "conventions": list}))
    });
    match result {
        Ok(v) => ok(v),
        Err(e) => err(e),
    }
}

/// `{engine, id, on}` → `{ok}`: switch one convention on or off.
pub fn set_convention(request: &str) -> String {
    let req = match parse(request) {
        Ok(r) => r,
        Err(e) => return err(e),
    };
    let id = req.get("id").and_then(Value::as_i64).unwrap_or(-1);
    let on = req.get("on").and_then(Value::as_bool).unwrap_or(true);
    let result = with_engine(&req, |engine| {
        let c = engine
            .bidder
            .base
            .conventions
            .get_mut(&id)
            .ok_or(format!("no convention {id}"))?;
        if on {
            c.selected.clear();
            c.sub_selected.clear();
        } else {
            c.selected = "0".into();
        }
        Ok(json!({}))
    });
    match result {
        Ok(v) => ok(v),
        Err(e) => err(e),
    }
}

/// `{engine, hands: {N, E, S, W}, dealer?: "N", vul?: "None"|"NS"|"EW"|"All"}`
/// → `{ok, calls: [{seat, call, rule?, meaning?, missing?}], contract,
/// declarer}`. Hands are PBN ("AKQ.JT9.876.5432") or BidBase form.
pub fn bid_deal(request: &str) -> String {
    let req = match parse(request) {
        Ok(r) => r,
        Err(e) => return err(e),
    };
    let result = with_engine(&req, |engine| {
        let hands_v = req.get("hands").ok_or("no \"hands\"")?;
        let mut hands: [String; 4] = Default::default();
        for (i, seat) in ["N", "E", "S", "W"].iter().enumerate() {
            let h = hands_v
                .get(*seat)
                .and_then(Value::as_str)
                .ok_or(format!("no hand for {seat}"))?;
            hands[i] = normalize_hand(h).ok_or(format!("cannot read {seat}'s hand"))?;
        }
        let dealer = str_field(&req, "dealer")
            .chars()
            .next()
            .and_then(Direction::from_char)
            .unwrap_or(Direction::North);
        let vul = Vulnerability::from_pbn(req.get("vul").and_then(Value::as_str).unwrap_or("None"))
            .ok_or("unknown vulnerability (None, NS, EW, All)")?;
        let (auction, calls, _) = table::bid_deal(
            &mut engine.bidder,
            &hands,
            dealer,
            vul,
            Scoring::Matchpoints,
            60,
        );
        let calls: Vec<Value> = calls
            .iter()
            .map(|c| {
                let mut v = json!({"seat": c.seat.to_char().to_string(), "call": c.call.to_pbn()});
                if let Some(f) = &c.found {
                    v["rule"] = json!(f.rule_id);
                    v["meaning"] = json!(f.meaning);
                }
                if let Some(m) = c.missing {
                    v["missing"] = json!(match m {
                        Missing::UnknownAuction => "no rules for this auction",
                        Missing::NoFit => "no rule fits the hand",
                        Missing::Illegal => "the rule's call is not legal here",
                    });
                }
                v
            })
            .collect();
        let (contract, declarer) = match auction.final_contract() {
            Some(fc) => (fc.to_pbn(), fc.declarer.to_char().to_string()),
            None => ("Pass".into(), String::new()),
        };
        Ok(json!({"calls": calls, "contract": contract, "declarer": declarer}))
    });
    match result {
        Ok(v) => ok(v),
        Err(e) => err(e),
    }
}

/// `{seed?}` → `{ok, hands: {N, E, S, W}}`: a random deal (PBN hands), for
/// pages without a dealer of their own.
pub fn deal(request: &str) -> String {
    let req = parse(request).unwrap_or(json!({}));
    let mut state = req
        .get("seed")
        .and_then(Value::as_u64)
        .unwrap_or(0x2545_F491_4F6C_DD1D)
        | 1;
    let mut next = || {
        state ^= state >> 12;
        state ^= state << 25;
        state ^= state >> 27;
        state.wrapping_mul(0x2545_F491_4F6C_DD1D)
    };
    let mut cards: Vec<u8> = (0..52).collect();
    for i in (1..52).rev() {
        let j = (next() % (i as u64 + 1)) as usize;
        cards.swap(i, j);
    }
    let mut out = serde_json::Map::new();
    for (k, seat) in ["N", "E", "S", "W"].iter().enumerate() {
        let hand = Hand::from_cards(
            cards[k * 13..(k + 1) * 13]
                .iter()
                .filter_map(|&c| bridge_types::Card::from_index(c))
                .collect(),
        );
        out.insert(seat.to_string(), json!(hand.to_pbn()));
    }
    ok(json!({"hands": out}))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn engine(extra: &str) -> u64 {
        let r: Value = serde_json::from_str(&create_engine(extra)).unwrap();
        assert_eq!(r["ok"], true, "{r}");
        r["engine"].as_u64().unwrap()
    }

    #[test]
    fn bids_with_the_2026_rules() {
        let id = engine("{}");
        let r: Value = serde_json::from_str(&bid(&format!(
            r#"{{"engine":{id},"hand":"Q8432.32.K32.J32","auction":"1NT Pass"}}"#
        )))
        .unwrap();
        assert_eq!(r["ok"], true, "{r}");
        assert_eq!(r["call"], "2H");
        assert!(r["meaning"].as_str().unwrap().contains("Jacoby"));
        assert!(!r["tried"].as_array().unwrap().is_empty());
    }

    #[test]
    fn the_2018_rules_and_sections() {
        let id = engine(r#"{"rules":"2018"}"#);
        let s: Value =
            serde_json::from_str(&section(&format!(r#"{{"engine":{id},"auction":"1N P"}}"#)))
                .unwrap();
        assert_eq!(s["key"], "1N-P");
        assert!(s["rules"].as_array().unwrap().len() > 20);
        let r: Value =
            serde_json::from_str(&rule(&format!(r#"{{"engine":{id},"id":1859}}"#))).unwrap();
        // Fields go by the database's column names.
        assert_eq!(r["rule"]["Bid"], "2H");
        assert!(r["rule"].get("HCP").is_none(), "empty fields are left out");
    }

    #[test]
    fn deals_and_bids_a_deal() {
        let id = engine("{}");
        let d: Value = serde_json::from_str(&deal(r#"{"seed":7}"#)).unwrap();
        let req = json!({"engine": id, "hands": d["hands"], "dealer": "S", "vul": "NS"});
        let r: Value = serde_json::from_str(&bid_deal(&req.to_string())).unwrap();
        assert_eq!(r["ok"], true, "{r}");
        assert!(r["calls"].as_array().unwrap().len() >= 4);
    }

    #[test]
    fn errors_are_json() {
        let r: Value = serde_json::from_str(&bid(r#"{"engine":999}"#)).unwrap();
        assert_eq!(r["ok"], false);
        let r: Value = serde_json::from_str(&create_engine("not json")).unwrap();
        assert_eq!(r["ok"], false);
    }
}
