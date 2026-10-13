# bidbase

Workspace for Nelson Ford's **CardShark BidBase**, a rule-based bridge
bidding program: Visual Basic 6 code plus an Access database of bidding
rules. All of it is public domain (his license page, below). Started
2026-10-09 from a rusty-bidding-bot session. Git repo, private on GitHub:
`bridge-craftwork/bidbase`.

## Goals (Rick, 2026-10-09)

1. **Understand the application**: engine, database, editor, practice program.
2. **Repair or port it so it runs again.** Nelson's programs stopped working
   in late September 2026 (error below). Either fix the VB6/Access setup or
   port the engine (Rust is the obvious choice, next to rusty-bidding-bot).
3. **Measure it against BBA and Rusty** on the same deals.
4. **Improve Rusty's rules from his ruleset**: coverage gaps, treatments,
   cited sources.

## People and status

- **Nelson Ford** (`www.aeyec.com`), author since about 2000;
  the project started earlier in Visual Basic. He posted on the bridge-dev
  forum (groups.io/g/bridge-dev) and wants to share his code.
- 2026-10-09: Rick offered to get it running. 2026-10-11: Nelson's
  current files arrived through his web site and are in **`aeyec-2026/`**
  (README there), rules exported to `data/2026`. His FTP paths map to the
  web from the `AIC` folder down (`AIC` is the web root);
  listing is off, so files are fetched by name, and IIS refuses unknown
  types (`.cls`, `.frm`, `.vbp`...: HTTP 404.3, a missing name gives
  404.0), hence his `BidBase.zip`. A forgotten **2023** copy of the rules
  was at `http://www.aeyec.com/CSBB/BidBase.accdb` (12,200 rules; kept
  locally in the gitignored `nelson-2026/`, not published).
- Publishing: the repo was made **public on 2026-10-11** so Nelson could
  see it first; on **2026-10-13 he gave written permission** to do anything
  with his files, including publishing or distributing them as-is or
  modified (PERMISSION.md quotes it). Before that, history was rewritten
  to drop his web-host account number and his handle. Keep personal
  details (accounts, email addresses, FTP paths) out of commits.
- **`Adodc1` error: found (2026-10-11).** The Editor's five ADO data
  controls `Data1(0..4)` (all captioned "Adodc1") run SQL (`RecordSource`
  in `BidBase Editor.frx` at 0xD3C9, 0xD6DE, 0xD9F0, 0xDD05, 0xE01A) that
  selects `Bids.LHO`, `Bids.Pard`, `Bids.RHO`. The 2023 database has those
  columns (Long Integer, plus `MySpecs`); the **2026 one doesn't**, so
  Access takes them for parameters: "No value given for one or more
  required parameters". The engine uses them too (`BidBaseDB_Bidder.cls`
  SQL at line 6492, `!LHO = lng_LHO` at 8541). **Fix sent to Nelson: add
  `LHO`, `Pard`, `RHO`, `MySpecs` (Number, Long Integer) back to `Bids`.**
  Only 3-5 rows had values in 2023.
- **2026 engine compared and ported (2026-10-11):** docs/ENGINE-2026.md
  lists every change from 2018 that decides a match. The port picks the
  dialect from the export (no `Strength_Discl` column = 2026) and follows
  the 2026 code (`checks2026.rs`). Self-test on his rules
  (`--data data/2026`): 13,133 of 16,673 first match, 840 reject their own
  test hand (was 1,681), the rest checked to match his code as written. On
  runs/random-500: 6.39 IMPs/board from par (2018 rules 6.61, BBA 4.32,
  Rusty 4.80). Not ported: `ComputeOptBid`, the double-dummy fallback when
  no rule fits.
- Nelson's design notes (email 2026-10-10): rules are in sections by the
  prior bids; within a section the first entry that matches the hand and
  the conventions in use wins (as the port does). Docs start at
  http://www.aeyec.com/bidbase/ (design overview first).

## What's here

```
crates/engine/               bidbase-engine: the Rust port of the VB6 engine (docs/PORTING.md)
crates/cli/                  bidbase-cli: the `bidbase` binary (bid, rule, selftest)
crates/forms/                bidbase-forms: reads VB6 .frm/.frx (control tree, properties, lists, pictures)
crates/vbui/                 bidbase-vbui: draws a parsed form with egui (classic look, Liberation fonts)
crates/practice/             bidbase-practice: the Bidding Practice program port (docs/PRACTICE-PROGRAM.md)
crates/practice-web/         its browser build (cdylib; crates/practice-web/build.sh)
.github/workflows/practice.yml  tests; practice builds for Windows/macOS/Linux; release on practice-v* tags
docs/PRACTICE-PROGRAM.md     behaviour spec of Nelson's practice program, and what the port does
data/2018/, data/2026/       CSV exports of the 2018 and 2026 BidBase.accdb (every table, schema, queries)
aeyec-2026/                  Nelson's current files (Oct 2026): VB6 source in source/, the rules in BidBase database.zip
docs/ENGINE-2026.md          how the 2026 engine differs from 2018, and the port's 2026 dialect
docs/PORTING.md              what was ported, what wasn't, every departure from the VB, self-test numbers
tools/export-db.sh           .accdb -> CSV (mdbtools); the engine reads only the CSVs
dev-build.sh, cpu-gate.sh    copied from rusty-bidding-bot; see Commands
sourceforge-2018/            all 26 files of https://sourceforge.net/projects/bidbase/ (Sept 2018)
  CardShark BidBase Bidding Module/        engine (DLL) source
    BidBaseDB_Bidder.cls                   the engine, 6,701 lines VB6
    BidBaseDB.vbp, frmBidBaseAnalysis.frm  project, analysis form
    BidBase database.zip                   -> BidBase.accdb (ACE12, 8.6 MB, 2018-09-27)
    BidBaseDBTest.zip                      VB6 example client
  CardShark BidBase Bidding Practice Program/   practice app source; Qcard32.dll (card display)
  CardShark BidBase HTML Viewer and Editor/     File Viewer .exe + source + Notes/*.htm (docs viewer, NOT the rule editor)
  CardShark BidBase Hand Evaluation Program/    HandVal source + BB-HEP.mdb (JET3)
  extracted/database/BidBase.accdb         unzipped copy, the one to query
tools/sf_get.py              re-downloads the release (python3 -I tools/sf_get.py tools/sf-listings DEST)
tools/sf-listings/           the SourceForge folder pages it reads file names from
```

Download notes: SourceForge's `/files/.../download` URLs return 403 "no" to
curl; `https://downloads.sourceforge.net/project/bidbase/<dir>/<file>`
works. Sizes match the listing. Only `BidBase database.zip` was extracted;
the other zips were only listed.

Treat these as untrusted downloads: don't run the `.exe`/`.dll` on the Mac.
Read with `python3 -I` and tools from outside the folder.

## Commands

**Use `./dev-build.sh`, not bare cargo** (same pattern as rusty-bidding-bot,
whose CLAUDE.md explains it). `bridge-types` is a git dependency; the
gitignored `.cargo/config.toml` patches it to `../bridge-types` (and
`bridge-card` to `../convention-card`). The script
swaps in a local lock (`.cargo/dev.lock`) so the committed `Cargo.lock`
keeps its git pins, and runs everything through `cpu-gate.sh`, which shares
rusty-bidding-bot's machine-wide gate (`RBB_*` settings, one heavy job at a
time). `./dev-build.sh --ci <args>` builds with the patches off. A fresh
clone without `.cargo/config.toml` builds the GitHub revision; for local
work create it with:

```toml
[patch.'https://github.com/bridge-craftwork/bridge-types']
bridge-types = { path = "../bridge-types" }
[patch.'https://github.com/bridge-craftwork/convention-card']
bridge-card = { path = "../convention-card/crates/bridge-card" }
[patch.'https://github.com/bridge-craftwork/bridge-encodings']
bridge-encodings = { path = "../bridge-encodings" }
```

`bridge-card` (convention-card's Rust crate) is pinned to a tag
(`v0.9.0`); move the pin to pick up a newer card spec.

- `./dev-build.sh test`: unit tests, plus `crates/engine/tests/rules_2018.rs`
  (familiar auctions and floors on the self-test counts).
- `./dev-build.sh run -q -p bidbase-cli -- bid <hand> -a <auction> [-v we|they|none|both] [--why] [-d]`:
  a call. Hands as `AK432-64-QJT3-T2` or PBN dots; auctions `P-1C-1N` or
  `Pass 1C 1NT`. `--why` lists every rule tried and the test it failed;
  `-d` prints the disclosure.
- `... -- rule <ID>`: a rule's fields and its convention's switches.
- `... -- selftest [--list own|first|all]`: every rule against its
  `Test_Hand`, as the Editor checks an entry.
- `--all-conventions` (any command) ignores the `BidNames.Selected`
  switches. In the 2018 copy 1,424 rules are switched off, including
  natural NT responses and four-card 1-over-1 responses to 1C.
- `--card <card.json>` (any command) sets the switches from a convention
  card in convention-card's native JSON, through `card/bidbase-map.toml`:
  each `BidNames` ID is on when its card fields say so, `always` for
  BidBase categories every system has, and unlisted IDs keep the
  database's setting. `bidbase card [-v]` shows the map's coverage;
  `bidbase --card X card -v` what a card switched. The map is checked
  against convention-card's `spec/fields.toml` when it loads (paths and
  enum values), so a spec rename fails the build's tests. Example card:
  `card/examples/21GF-DEFAULT.card.json` (from convention-card's goldens).
- `... --card X.json -- bid-pbn -i deals.pbn -o out.pbn`: bid every deal
  at all four seats, written in bba-cli's layout. No rule means Pass (as
  the 2018 practice program); after partner's 4NT the "RKCB: 4N-x" /
  "Blackwood: 4N-x" keys are tried.
- `python3 -I tools/compare_runs.py DD.jsonl NAME=FILE.pbn ...`: score
  engines' contracts double-dummy against par (`rbb dd` output), with
  agreement and head-to-head IMPs. `runs/random-500/` is the first run
  (README there): BidBase 6.6 IMPs/board from par, BBA 4.3, Rusty 4.8;
  BidBase's gap is rule coverage (no rule for 68% of calls).
- `./dev-build.sh run --release -p bidbase-practice`: the practice program
  port (window 664 x 598, his form). `BIDBASE_SHOTS=<dir> ./dev-build.sh
  test -p bidbase-practice --test screens` writes its screens as PNGs
  (offscreen egui_kittest; capturing the real window on the Mac was
  unreliable). The GUI crates use eframe with `glow` and no egui default
  fonts (vbui bundles Liberation Sans/Mono), which keeps the web build
  near 3 MB gzipped.
- **Web page: https://bridge-craftwork.com/bidbase/** (live 2026-10-12;
  Cloudflare Pages project `bidbase`, mounted by bridge-craftwork-site's
  router, `TOOL_ORIGINS['/bidbase']`). `web/build.sh` builds `web/dist`
  (engine from `crates/wasm` with the 2018 and 2026 rules compiled in, ~1 MB
  compressed); deploy with `npx --yes wrangler@4 pages deploy --branch
  main` from the repo root. Tabs: bid a hand (call, meaning, every rule
  tried), bid a deal, rules of an auction, conventions. The hand tab's
  inputs live in the URL fragment, so links show a given bid.
- `crates/wasm/src/api.rs`: the JSON API (createEngine, bid, rule,
  section, conventions, setConvention, bidDeal, deal), tested natively.
- `tools/export-db.sh <BidBase.accdb> data/<name>`: export a new copy;
  `--data data/<name>` (or `BIDBASE_DATA`) points the CLI at it.

## Reading the databases

`mdbtools` is installed (Homebrew) and reads both files: `mdb-ver`,
`mdb-tables -1`, `mdb-schema -T <table>`, `mdb-export <file> <table>`
(CSV), `mdb-queries`. Rick's `../bridge-parsers` Rust crate wraps the same
tools (`src/bws/reader.rs`: public `list_tables`, private `export_table`,
`read_bws` is Bridgemate-specific). Use it or call `mdb-export` directly.

`BidBase.accdb` (ACE12):

| Table | Rows | What |
|---|---|---|
| `Bids` | 6,314 | the rules |
| `BidNames` | 471 | conventions: `IDnum`, `Selected`, `SubCat_Selected`, `Bid_Name`, `Bid_Description` |
| `Garbage` | 348 | retired rules (same layout as `Bids`, older column names, `GarbargeID`) |
| `DupePBs` | 202 | groups of prior-bid strings treated alike (`Description`, `pb1..pb9`) |

Saved queries: `BidNamesQuery`, `BidNames Query Backup`, `GridQuery`.

`Bids` columns (most are text ranges like `22+`, `P23-25`, `5-7`):
`Pct_Used`, `RoundOfBidding`, `Player`, `PB_Sort_Order`, `Prior_Bids`
(e.g. `1N-2C`), `Bid_Order_Number` (priority within an auction; `ZZZZZZ`
ends a section), `Bid`, `ConvID` and `AgainstID` (-> `BidNames.IDnum`),
`HCP`, `ValSystem`, `SuitVal`, `NT_Val`, `Cust_Pts` with `A_Cust..T_Cust`
(custom honor values), `Dist_Pts`, `Winrs`, `Losers`, `QTs`,
`Stopped_Suits`, `Shape`, `Intermed`, per-suit `xNum/xHCP sNum/sHCP
hNum/hHCP dNum/dHCP cNum/cHCP` (x = the bid suit), `cmNum`/`cmHCP`
(combinations), `Vul`, `Pos`, `Artificial`, `Alert`, `Explanation`,
`Test_Hand` (e.g. `AK2-AK2-AK2-K432`), `Strength_Discl`, `xLen SLen HLen
DLen CLen` (lengths disclosed to partner), `Asks`, `DateOfEntry` (2004
onward), `Best_Contract`, `Num_Used`, `Source`, `ID`.

`BB-HEP.mdb` (JET3): `HandVal` (valuation systems: per-feature points by
NT/suit, position, fit, cards, length) and `HandVal_BidReq` (per valuation
system, the points needed for openings, responses, game, slam, NT ranges).

## How the engine works (from the 2018 source)

- Client sets properties (`DBPathAndName`, `PlayersHand`, `PriorCalls`,
  `Vul`, `ScoringMethod`, `TrumpSuit`, ...) and reads `Bid`, `id`,
  `Disclosure`, `ErrMsg`.
- `LookUpBid`: `AnalyzeHand` computes HCP, winners, losing-trick count,
  quick tricks, stoppers, shape and position. Then
  `SELECT * FROM Bids WHERE Prior_Bids='<auction>' ORDER BY
  Bid_Order_Number`, and the **first** row whose tests all pass wins.
  The tests run in order: `PctUsed_Ok` (random use by frequency),
  `TotalHCPs_Ok`, `CustPoints_Ok`, `DistribPts_Ok`, `Winners_Ok`,
  `Losers_Ok`, `QuickTricks_Ok`, `Stoppers_Ok`, `Shape_Ok`,
  `Intermed_Ok`, `Vul_Ok`, `Position_Ok`, `Suits_Ok` (the longest),
  `Eval_Pts_Ok`.
- Opening passes are ignored in the prior-bids string. Nelson's notes
  describe shortcuts (`1M-P`/`1m-P` for both majors/minors, `@` for
  ace-asking, `DupePBs` for auctions sharing rules), but **neither the
  2018 engine nor the 2018 rules use them**; slam replies are keyed by
  pseudo-auctions like `Blackwood: 4N-P`, `RKCB: 4N-P`, `Gerber: 4C-P`.
  Probably Editor-side or newer.
- `GetOtherPlayersData` / `GetOPsBid` (partner's and the opponents' suits,
  fits) are skipped on the bidding path (`GoTo SkipPard`), so the 2018
  engine models no other hand. It looks only at the bidder's hand,
  the auction key, seat and vulnerability.
- The 2018 code has real bugs (NT ranges read from the wrong rules, "P"
  HCP specs never converted, seats counted from the opener...). The port
  does what was meant; docs/PORTING.md lists each case with VB line
  numbers.
- `GetDiscl` / `AddToDisclosure` build the description of a bid shown to
  the user.

## Compared with rusty-bidding-bot

Same core idea: the engine holds no bidding knowledge; rules are data
giving auction, call, hand requirements and priority, and the first match
wins. Conventions can be switched on and off (`BidNames.Selected`, like our
card fields). Each rule carries an explanation, alert and artificial flags,
a test hand and a source, much like our `.test` files and notes.

Differences: rules are database rows edited in a GUI, not text in git. The
conditions are a fixed set of ~59 columns, not an expression language over
named terms. Rules are keyed by the exact prior-bids string (plus the
shortcuts above), not named auction patterns with systems reused from
every entry. There's no judgment layer after the rule match, though
`Pct_Used` can pick randomly. Hand valuation is a separate database. The
rule set grew one auction at a time since 2004.

## Status of goal 2: the port (2026-10-09)

Done for the 2018 engine, without hand valuation (3 rules need it). Of
5,258 rules with test hands: 3,440 are the first match for their own hand,
1,454 more are switched off but accept it, 289 are taken by an earlier
rule, 66 reject their own hand (mostly 2004 test hands that don't fit
their rules), 9 test hands have typos. Details and the open questions in
docs/PORTING.md. Next: port Nelson's current engine and data when his zip
arrives (diff its `.cls` against 2018 first), and choose a convention
profile for comparisons.

## Nelson's GUIs: Rust ports, not VB6 builds (2026-10-13)

Rick decided against building Nelson's VB6 programs in CI: VB6 is
licensed, not on GitHub's runners and not installable from a public
workflow; Rick has no licence; twinBASIC's command-line build is
experimental and the Editor's ADODC/DataGrid binding is unproven there.
Instead the GUIs are ported to Rust as look-alikes ("don't have to be 100%
picture perfect"), native and web, laid out from Nelson's own `.frm`
files. Order: the Bidding Practice program first (done as an MVP; section
9 of docs/PRACTICE-PROGRAM.md), then the Editor, which will edit the CSVs
under `data/` (Rick's choice; nothing writes `.accdb`), then BidSim and the
File Viewer. His screenshots of every program are in the SourceForge File
Viewer zip (`Notes/*.jpg`), the targets to match.

## Ideas for goals 3-4 (unverified)

- **Measure:** have the port bid PBN deals and write auctions in bba-cli's
  PBN layout. Then score with rusty-bidding-bot's tools: `rbb dd`/par, and
  `compare`, which reads PBS-shaped directories of bba-cli output, as
  `probes/tools/conv_ab.py` does. Check those tools' current flags before
  relying on this. BidBase plays its own system (SAYC/2/1-ish, check
  `BidNames`), so pick scenarios and Rusty cards that match, or compare
  only the auctions both systems define.
- **Improve Rusty:** map `Prior_Bids` + `Bid` coverage against Rusty's
  rules to find auctions BidBase defines and Rusty doesn't; read its
  `Source` citations and the 100+ notes pages (`www.aeyec.com/bidbase/Notes/Index.htm`,
  also in the File Viewer zip) for treatments. Changes to Rusty follow its
  own rules (CLAUDE.md there: notes with Sources, `.test` cases, fast lane).

## Ground rules

- BidBase is public domain and unrelated to BBA/EPBot, so rusty-bidding-bot's
  clean-room rule (no decompiled BBA) does not apply to it. Cite BidBase
  (and Nelson) as a source wherever its rules inform Rusty's.
- His files are kept byte for byte (`.gitattributes`: no line-ending
  conversion). Don't run his `.exe`/`.dll` here; read the source.
- The repo is public. Port, tools and docs: MIT OR Apache-2.0. Nelson's
  files (`sourceforge-2018/`, `aeyec-2026/`, `data/`): his public-domain
  work, published with his permission (PERMISSION.md).

## Sources

- http://www.aeyec.com/bidbase/ (license: "All the databases, text files
  and programs making up BidBase are free public domain"; BidBaseDB.DLL
  free for any bridge software)
- http://www.aeyec.com/bidbase/CardShark_BidBase_Programming_Notes.htm
- http://www.aeyec.com/bidbase/Notes/Index.htm, incl.
  CardShark_BidBase_Editor_Overview.htm (how the database is structured)
- https://sourceforge.net/projects/bidbase/ (Sept 2018 release)
- Nelson's emails to Rick, 2026-10-09
