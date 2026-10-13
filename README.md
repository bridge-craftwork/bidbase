# bidbase

A Rust port of the bidding engine of Nelson Ford's
[CardShark BidBase](http://www.aeyec.com/bidbase/), a rule-based bridge
bidder whose knowledge lives in an Access database of rules (5,465 in the
2018 release, 16,786 in 2026), together with Nelson's original Visual
Basic 6 programs and rules it is ported from.

```
./dev-build.sh run -q -p bidbase-cli -- bid AK2-KQ2-A432-J32
1N  (rule #13) Opening 1N HCP Specification: ...
./dev-build.sh run -q -p bidbase-cli -- --data data/2026 bid Q8432-32-K32-J32 -a "1NT Pass" --why
./dev-build.sh run -q -p bidbase-cli -- --data data/2026 selftest
```

- `crates/engine`: the engine, reading both the 2018 rules and Nelson's
  2026 rules (`docs/PORTING.md`, `docs/ENGINE-2026.md`)
- `crates/cli`: the `bidbase` command: bid a hand (with why each rule
  matched or not), check every rule against its test hand, bid PBN deals,
  replay another engine's auctions
- `crates/wasm`: the engine for the browser (JSON API)
- `crates/practice`: Nelson's Bidding Practice program, ported from VB6:
  his screen (read from his form file), his deals, the engine's calls and
  a double-dummy table, for Windows, macOS, Linux and the browser
  (`docs/PRACTICE-PROGRAM.md`). `crates/forms` reads VB6 form files;
  `crates/vbui` draws them with egui; `crates/practice-web` is the
  browser build.
- `card/`: a map from a convention card ([convention-card](https://github.com/bridge-craftwork/convention-card)
  JSON) to BidBase's convention switches
- `data/2018`, `data/2026`: the rules, exported to CSV (`tools/export-db.sh`)
- `sourceforge-2018/`: Nelson's September 2018 SourceForge release, untouched
- `aeyec-2026/`: Nelson's current programs and rules (October 2026)
- `runs/random-500`: BidBase, BBA and rusty-bidding-bot bidding the same
  500 deals, scored double-dummy

Build with `./dev-build.sh` (see CLAUDE.md). The practice program:
`./dev-build.sh run --release -p bidbase-practice`; builds for Windows,
macOS and Linux come from GitHub Actions (`.github/workflows/practice.yml`:
artifacts on every run, a release for each `practice-v*` tag).

## License

The Rust port, tools and documentation in this repository are licensed
under either of [MIT](LICENSE-MIT) or [Apache-2.0](LICENSE-APACHE), at
your option.

BidBase itself (the programs, source and rules in `sourceforge-2018/`,
`aeyec-2026/` and `data/`) is Nelson Ford's work. His web site says: "All
the databases, text files and programs making up BidBase are free public
domain", and BidBaseDB.DLL is free for use in any bridge software. On
13 October 2026 he gave this project his specific permission to publish
and distribute his files, as they are or modified: see
[PERMISSION.md](PERMISSION.md).
