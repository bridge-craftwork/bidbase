//! Convention card to BidBase switches: `card/bidbase-map.toml` says which
//! card fields (convention-card's native JSON) turn each `BidNames`
//! convention on or off.

use std::collections::BTreeMap;

use bridge_card::{Card, Value, Vocabulary};

use crate::rules::RuleBase;

/// The map as shipped with this repository.
pub const MAP: &str = include_str!("../../../card/bidbase-map.toml");

/// When a convention is on.
#[derive(Debug, Clone, PartialEq)]
pub enum Cond {
    /// A bool field is true (or false, when `want` is false).
    Flag { path: String, want: bool },
    /// Every field has its value.
    All(Vec<(String, Value)>),
    /// Any of these holds.
    Any(Vec<Cond>),
}

#[derive(Debug, Clone, PartialEq)]
pub enum Switch {
    /// A BidBase category every system has.
    Always,
    /// Decided by the card.
    On(Cond),
}

#[derive(Debug, Clone)]
pub struct Entry {
    pub id: i64,
    pub name: String,
    pub switch: Switch,
}

/// The parsed, checked map.
#[derive(Debug, Clone)]
pub struct CardMap {
    pub entries: BTreeMap<i64, Entry>,
}

/// What applying a card did to each convention.
#[derive(Debug, Clone, Default)]
pub struct Applied {
    pub on: Vec<i64>,
    pub off: Vec<i64>,
    /// In the database but not in the map: left as the database has it.
    pub kept: Vec<i64>,
}

impl CardMap {
    /// The repository's map, checked against convention-card's spec.
    pub fn standard() -> Result<(CardMap, Vocabulary), String> {
        let vocab = bridge_card::standard::vocabulary().map_err(|e| e.to_string())?;
        let map = CardMap::parse(MAP, &vocab)?;
        Ok((map, vocab))
    }

    /// Parse a map, checking each path and value against the card's
    /// field registry.
    pub fn parse(text: &str, vocab: &Vocabulary) -> Result<CardMap, String> {
        let table: toml::Table = text.parse().map_err(|e| format!("bidbase-map.toml: {e}"))?;
        let mut entries = BTreeMap::new();
        for (key, v) in &table {
            let id: i64 = key
                .parse()
                .map_err(|_| format!("{key}: key is not a BidNames ID"))?;
            let t = v.as_table().ok_or_else(|| format!("{key}: not a table"))?;
            let name = t
                .get("name")
                .and_then(|n| n.as_str())
                .unwrap_or("")
                .to_string();
            let switch = match (t.get("always"), t.get("on")) {
                (Some(toml::Value::Boolean(true)), None) => Switch::Always,
                (None, Some(on)) => {
                    Switch::On(parse_cond(on, vocab).map_err(|e| format!("{key} ({name}): {e}"))?)
                }
                _ => {
                    return Err(format!(
                        "{key} ({name}): needs `on = ...` or `always = true`"
                    ))
                }
            };
            entries.insert(id, Entry { id, name, switch });
        }
        Ok(CardMap { entries })
    }

    /// Set the rule base's convention switches from a card.
    pub fn apply(&self, card: &Card, base: &mut RuleBase) -> Applied {
        let mut applied = Applied::default();
        for (id, conv) in base.conventions.iter_mut() {
            let on = match self.entries.get(id).map(|e| &e.switch) {
                None => {
                    applied.kept.push(*id);
                    continue;
                }
                Some(Switch::Always) => true,
                Some(Switch::On(cond)) => holds(cond, card),
            };
            if on {
                conv.selected.clear();
                conv.sub_selected.clear();
                applied.on.push(*id);
            } else {
                conv.selected = "0".into();
                applied.off.push(*id);
            }
        }
        applied.on.sort();
        applied.off.sort();
        applied.kept.sort();
        applied
    }
}

fn parse_cond(v: &toml::Value, vocab: &Vocabulary) -> Result<Cond, String> {
    let registry = vocab.registry();
    match v {
        toml::Value::String(s) => {
            let (path, want) = match s.strip_prefix('!') {
                Some(p) => (p, false),
                None => (s.as_str(), true),
            };
            let field = registry
                .get(path)
                .ok_or_else(|| format!("no card field {path:?}"))?;
            if field.normalize(Value::Bool(want)).is_err() {
                return Err(format!("{path:?} is not a yes/no field"));
            }
            Ok(Cond::Flag {
                path: field.path.clone(),
                want,
            })
        }
        toml::Value::Table(t) => {
            let mut all = Vec::new();
            for (path, want) in t {
                let field = registry
                    .get(path)
                    .ok_or_else(|| format!("no card field {path:?}"))?;
                let value = match want {
                    toml::Value::Boolean(b) => Value::Bool(*b),
                    toml::Value::Integer(i) => Value::Int(*i),
                    toml::Value::String(s) => Value::Text(s.clone()),
                    other => return Err(format!("{path:?}: unsupported value {other}")),
                };
                let value = field
                    .normalize(value)
                    .map_err(|e| format!("{path:?}: {e}"))?;
                all.push((field.path.clone(), value));
            }
            Ok(Cond::All(all))
        }
        toml::Value::Array(items) => Ok(Cond::Any(
            items
                .iter()
                .map(|i| parse_cond(i, vocab))
                .collect::<Result<_, _>>()?,
        )),
        other => Err(format!("unsupported condition {other}")),
    }
}

fn holds(cond: &Cond, card: &Card) -> bool {
    match cond {
        Cond::Flag { path, want } => card.is_on(path) == *want,
        Cond::All(all) => all
            .iter()
            .all(|(path, want)| match (card.effective(path), want) {
                // An unset bool reads as false.
                (None, Value::Bool(false)) => true,
                (have, want) => have == Some(want),
            }),
        Cond::Any(any) => any.iter().any(|c| holds(c, card)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_map_loads_and_checks_against_the_spec() {
        let (map, _) = CardMap::standard().unwrap();
        assert!(map.entries.len() > 300);
        assert_eq!(map.entries[&510].switch, Switch::Always);
    }

    #[test]
    fn bad_paths_and_values_are_errors() {
        let vocab = bridge_card::standard::vocabulary().unwrap();
        assert!(CardMap::parse("1 = { on = \"no.such.field\" }", &vocab).is_err());
        assert!(CardMap::parse("1 = { on = { \"overcalls.jump\" = \"huge\" } }", &vocab).is_err());
        assert!(CardMap::parse("1 = { name = \"x\" }", &vocab).is_err());
        assert!(CardMap::parse(
            "1 = { on = [\"notrump.stayman.play\", \"!slam.ropi\"] }",
            &vocab
        )
        .is_ok());
    }

    #[test]
    fn conditions_read_the_card() {
        let vocab = bridge_card::standard::vocabulary().unwrap();
        let mut card = Card::new(&vocab);
        card.set("notrump.stayman.play", Value::Bool(true)).unwrap();
        let cond = |s: &str| parse_cond(&s.parse::<toml::Table>().unwrap()["on"], &vocab).unwrap();
        assert!(holds(&cond("on = \"notrump.stayman.play\""), &card));
        assert!(holds(&cond("on = \"!slam.ropi\""), &card));
        // general.system_category defaults to two_over_one.
        assert!(holds(
            &cond("on = { \"general.system_category\" = \"two_over_one\" }"),
            &card
        ));
        assert!(!holds(
            &cond("on = { \"general.system_category\" = \"precision\" }"),
            &card
        ));
        assert!(holds(
            &cond("on = [\"slam.ropi\", \"notrump.stayman.play\"]"),
            &card
        ));
    }
}
