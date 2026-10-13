//! Every form in Nelson's 2018 and 2026 sources parses, and the values the
//! ports rely on come out as expected.

use std::path::{Path, PathBuf};

use bidbase_forms::{frm, frx};

fn root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
}

fn forms_under(dir: &Path, out: &mut Vec<PathBuf>) {
    for e in std::fs::read_dir(dir).unwrap().flatten() {
        let p = e.path();
        if p.is_dir() {
            forms_under(&p, out);
        } else if p.extension().is_some_and(|x| x.eq_ignore_ascii_case("frm")) {
            out.push(p);
        }
    }
}

#[test]
fn all_forms_parse() {
    let mut files = Vec::new();
    forms_under(&root().join("aeyec-2026/source"), &mut files);
    forms_under(&root().join("sourceforge-2018"), &mut files);
    assert!(files.len() >= 10, "found only {} forms", files.len());
    for f in files {
        let form = frm::parse(&std::fs::read(&f).unwrap())
            .unwrap_or_else(|e| panic!("{}: {e}", f.display()));
        assert_eq!(form.root.kind(), "Form", "{}", f.display());
        assert!(!form.vb_name.is_empty(), "{}", f.display());
    }
}

#[test]
fn practice_form() {
    let dir = root().join("aeyec-2026/source");
    let form =
        frm::parse(&std::fs::read(dir.join("BidBase Bidding Practice.frm")).unwrap()).unwrap();
    assert_eq!(form.vb_name, "PlayScr");
    assert_eq!(form.root.int("ClientWidth"), Some(9960));
    // 52 card pictures in the deck frame, four hands.
    assert_eq!(form.all("Card").len(), 52);
    assert_eq!(form.all("p_Hand").len(), 4);
    let grid = form.get("g_BidBox", None).unwrap();
    assert_eq!((grid.int("Rows"), grid.int("Cols")), (Some(8), Some(5)));
    let deal = form.get("b_Deal", None).unwrap();
    assert_eq!(deal.caption(), "Deal");
    // b_Deal sits in p_Bottom, which starts at (30, 6405).
    assert_eq!(form.origin_of(deal), (30 + 2520, 6405 + 75));

    let bytes = std::fs::read(dir.join("BidBase Bidding Practice.frx")).unwrap();
    let t = form.get("t_Calculating", None).unwrap();
    let frm::Value::Frx { offset, long, .. } = t.props["Text"].clone() else {
        panic!("t_Calculating text not in frx")
    };
    assert!(frx::text(&bytes, offset, long)
        .unwrap()
        .contains("Calculating DDA"));
    let frm::Value::Frx { offset, .. } = form.root.props["Icon"].clone() else {
        panic!("no icon")
    };
    let icon = frx::picture(&bytes, offset).unwrap();
    assert_eq!(&icon[..4], &[0, 0, 1, 0], "an .ico file");
}

#[test]
fn editor_record_source() {
    let bytes = std::fs::read(root().join("aeyec-2026/source/BidBase Editor.frx")).unwrap();
    let sql = frx::text(&bytes, 0xD3C9, true).unwrap();
    assert!(sql.starts_with("SELECT "), "{sql}");
    assert!(sql.contains("Bids.LHO"));
}
