# bidbase

Workspace for Nelson Ford's **CardShark BidBase**, a rule-based bridge
bidding program: Visual Basic 6 code plus an Access database of bidding
rules. All of it is public domain (his license page, below). Started
2026-10-09 from a rusty-bidding-bot session. Not a git repo yet.

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
- 2026-10-09: Rick offered to get it running. Nelson uploaded his
  **current** code and data with FileZilla to his website but sent no URL.
  Google Drive upload failed for him. **Rick is asking him for the exact
  URL.** Folder listing on aeyec.com is off (IIS 403.14), so we can't
  browse. `http://www.aeyec.com/bidbase/source/` and `/bidbase/Source/`
  exist (403, not 404), so the upload is probably there. Guessed names
  (`BidBase.zip`, `source.zip`, ...) all 404.
- Until that arrives we have only the **September 2018 SourceForge
  release** (below).

## The error Nelson sees

A message box titled **`Adodc1`** saying **"No value given for one or more
required parameters."** (screenshot in his email; he says nobody online has
a fix that works for him).

- `Adodc1` is the default name of a VB6 **ADO Data Control** on a form.
  The message is the Access engine's error 0x80040E10: a SQL statement
  names a column or table the engine can't find, so it treats the name as
  a parameter and has no value for it.
- So: some query no longer matches the database it opens. Likely causes,
  most likely first: (a) the database changed (a renamed or dropped column,
  an older copy restored); (b) the program opens a different file than
  intended (a config path, or Windows VirtualStore redirection of files
  under Program Files); (c) a Windows or Office update changed the Access
  database engine (ACE) that reads `.accdb`. No 2026 ACE fix found matching
  this exact error, but (c) fits "stopped two weeks ago, nothing changed".
- **No `Adodc1` in the 2018 source.** The failing program is probably his
  BidBase Editor (not in the SourceForge release) or a newer build, so the
  diagnosis needs his current zip.
- Leads in the 2018 code: `InitializeBidBaseDLL` opens a saved query
  `GridQuery`. A join at `BidBaseDB_Bidder.cls` ~line 4792 names
  `BidNames.Selected`, `BidNames.SubCat_Selected`, `UseAgainst.Bid_Name`,
  `Bids.DateOfEntry`, ... **The column names have changed before:** the
  `Garbage` table (same layout as `Bids`) has `Round_of_Bidding` and
  `Date_of_Entry` where `Bids` has `RoundOfBidding` and `DateOfEntry`.
- **Plan once his zip arrives:** pull every SQL string and every
  `RecordSource` (from the `.frm` files and code) and every saved query
  (`mdb-queries`, `mdb-schema`). Check each column name against the
  current `.accdb` schema; the one that doesn't resolve is the bug. Also
  diff his schema against the 2018 one. A fix can be found on the Mac; it
  needs Windows (VM or Nelson) to test the VB6 build.

## What's here

```
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
- Opening passes are ignored in the prior-bids string. Shortcuts: `1M-P` /
  `1m-P` cover both majors or minors, NT-overcall auctions reuse `1N-P`
  ("systems on"), `@` in place of the prior bids marks ace-asking and
  control bidding, and `DupePBs` lists auctions that share rules.
- `GetOtherPlayersData` / `GetOPsBid` look at other players' calls; how
  far this goes toward modelling partner's hand is **not yet read**.
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

## Ideas for goals 2-4 (unverified)

- **Port:** a Rust reader for `BidBase.accdb`, via `mdb-export` or
  bridge-parsers, turning each `Bids` row into a typed rule. Then a
  matcher reproducing `LookUpBid` and the `*_Ok` tests from the VB6
  source. Check it against the `Test_Hand` column: each rule should bid
  its own test hand. That needs no Windows at all. Keep the
  `.accdb` as the source of truth so Nelson can keep editing in his
  Editor, or offer him an export to text.
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
- Ask Nelson before publishing his current code anywhere (GitHub etc.); the
  2018 release is already public.

## Sources

- http://www.aeyec.com/bidbase/ (license: "All the databases, text files
  and programs making up BidBase are free public domain"; BidBaseDB.DLL
  free for any bridge software)
- http://www.aeyec.com/bidbase/CardShark_BidBase_Programming_Notes.htm
- http://www.aeyec.com/bidbase/Notes/Index.htm, incl.
  CardShark_BidBase_Editor_Overview.htm (how the database is structured)
- https://sourceforge.net/projects/bidbase/ (Sept 2018 release)
- Nelson's emails to Rick, 2026-10-09
