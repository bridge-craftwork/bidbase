# The 2026 engine against the 2018 engine

`BidBaseDB_Bidder.cls` as Nelson sent it on 2026-10-11 (10,955 lines,
last changed 2026-09-28; in `aeyec-2026/source/`) against the September 2018 SourceForge release (6,701
lines). Only what decides which rule matches is listed; analysis-form
and editor plumbing is left out. Line numbers are in the 2026 file.

## Lookup (`LookUpBid`, 768)

- Same loop: rules for the prior bids in `Bid_Order_Number` order, first
  rule passing every test wins, "ZZZZZZ" ends the section.
- Tests, in order: `Activation_Ok`, `TotalHCPs_Ok`, `CustPoints_Ok`,
  `DistribPts_Ok`, `Winners_Ok`, `Losers_Ok`, `QuickTricks_Ok`,
  `Stoppers_Ok`, `Shape_Ok`, `Intermed_Ok`, `Vul_Ok`, `Position_Ok`,
  `Suits_Ok`. **`Eval_Pts_Ok` is gone** ("Val systems no longer used");
  the `ValSystem` columns were dropped.
- No rule found and all four hands known (`int_OptAvg > 0`, the practice
  program): after a 1NT overcall ("1C-1N-P", "1D-1N-P", "1H-1N-P",
  "1S-1N-P") the lookup is repeated under "1N-P" (systems on); failing
  that, **`ComputeOptBid`** picks a bid from double-dummy analysis of the
  deal (`GetDDA`, `Optimum*`, DDS via `dds.dll`). Not ported.
- Matched rule whose bid has "x" but no x suit was found: if the hand
  passed `Winners_Ok` with a winners spec, x is the longest suit.

## Switching rules on and off (`Activation_Ok`, 2518; was `PctUsed_Ok`)

Same `Pct_Used` and `BidNames.Selected` logic. The entry under test now
passes before anything else. A new "blocked bid" rule (entries numbered
`...9999` block their bid for the rest of the section) cannot fire: it
reads `sel`/`sel2` from a lookup of convention 0, so they are always "".

## Hand analysis

- **Losing-trick count** (`Calculate_LosingTrickCount`): a new table with
  half losers (AQ doubleton 0.5, AJT 1.5, KQT 1.5, Qx(x) 2.5, Qxx no J/T
  2.5), but the counts are stored in an `Integer`, so each suit rounds
  half to even: 0.5 to 0, 1.5 to 2, 2.5 to 2.
  Void 0; singleton 1 unless the ace; doubleton AK 0, AQ 0.5, Ax 1, Kx 1,
  else 2; three or more: AKQ 0, AK 1, AQJ/AQT 1, AJT 1.5, A 2, KQJ 1,
  KQT 1.5, K 2, QJ/QT 2, Q 2.5, else 3.
- **Shape** (`Calculate_Shape`): a 5-4 is now a two-suiter ("2"), 5-4-4-0
  included (was "3").
- **Stoppers** (`Calculate_Stoppers`): QJ with the 9 or 8 is 2; QT8x is 2
  like QT9x; Qxx is 0.5 (was 1); J with five or more cards is 1.
- **Solid suits** (`Calculate_Winners`): five+ cards with all 10 top HCP,
  six+ headed AKQT, nine+ with 9 HCP, eight+ with 9+ HCP. Used by the suit
  spec "SS".
- **Ace/king counts** (`Calculate_CustPoints`, new): aces, kings, trump K
  and Q for ace-asking replies; the trump suit comes from the client or
  from the "@" code in the prior bids (below).
- 1NT, 2NT, 3NT ranges (`GetOpeningNTpts`, replaces `GetMinNTpts`): from
  the first active opening rule for each bid (1NT: the one named "Opening
  1N HCP Specification"). This is what the port already does.

## Tests

- **HCP** (`TotalHCPs_Ok`): "P" partnership specs now work (the 2018
  case bug is fixed). New "1N+a-b": 1NT minimum plus a to b. "Min" means
  the 2NT range only when the bid is 2NT. **No intermediates credit** for
  extra HCP (the line is commented out); ">n" still reads as 0-40.
- **Distribution points**: new "Rule20" (HCP plus the two longest suits,
  20 or more; the length tracking is buggy). No intermediates credit.
- **Winners**: "<n" is fewer than n; anything else a minimum.
- **Losers**: **a number is now a maximum** ("# in DB means <= #"), "n+"
  a minimum; "<n" never matches (`Val("<5")` is 0).
- **Intermediates**: "0" means exactly none; otherwise a minimum.
- **Vulnerability**: one-letter codes compared with what the client
  passes: B both, N neither, U unfavourable, F favourable, V we are
  vulnerable, X we are not (matches any client code but U and V). The
  practice program passes U, F, B, or "" when neither side is
  vulnerable, so a rule with "N" never matches there and "V" only in the
  Editor.
- **Position** (`Position_Ok`, `Position_calc`): seat counted from the
  dealer over the full prior calls. "1"/"F" first seat, "2", "3"/"T" third
  seat, "4"; "Y"/"N" the bidder's first call was / wasn't a pass; "P"/"U"
  partner's first call was / wasn't a pass (for seats 1-2 "U" looks at the
  bidder's own first call).

## Suits (`Suits_Ok` 3803, `Suit_Ok`, `SuitQty_Ok`, `SuitHCP_Ok`, `Combo_Ok`)

Rewritten. New:

- Suits are tested in order and **the first failing suit fails the rule**
  (unless its length spec is exactly "x").
- **x suits**: "x" in either the length or the strength field takes
  `xNum`/`xHCP`; "x,..." alternatives split differently; "<" in `xNum`
  removes x from the suits not yet tested; the matched x suit is the
  longest candidate.
- `Suit_Ok`: a length spec is required (a suit with only a strength spec
  fails: 43 suits in the 2026 data); "or..." in the strength field means
  length **or** strength.
- `SuitQty_Ok`: "<n" and ">n" read `Val` of the text after the sign.
- `SuitHCP_Ok`, new specs: "SS" solid suit, "<A" / "<K" (no ace / no ace
  or king), "hh" (short or headed by A/K), "ha", "n/m" (n of the top m
  honours, ace counted), "Hn"/"Hnm" honour counts, "A/K", "A" anywhere
  with a points prefix, the "ab"/"bc"/"a"/"b"/"c"/"d" lists matched on the
  suit's top cards.
- "Trmp" in a suit's strength field makes it the trump suit.
- `Combo_Ok`: ace-asking specs read the counts from
  `Calculate_CustPoints` ("2A+TK-TQ"); suit sums as before, with "x"
  allowed in the second letter. "M..." specs (the major) depend on a
  variable that is never set.

## Ace asking: "@" keys

A slam-asking rule's bid carries a code: `4N@30` is 4NT with hearts (3)
agreed, `4D@CS` a control bid. The practice program keeps the whole
token in the auction; for the following calls it passes only the calls
from the ask on ("4N@30-P", the rules' key) and sets the trump suit from
the digits (`TrumpSuit` 30 means suit 3). The 2026 data has 92 such keys.

## Not ported

`ComputeOptBid` and the double-dummy optimum (needs all four hands and a
solver), the analysis form, the editor's save and test plumbing
(`SaveBid`, `SaveBidEntry`, `Activated`, `Combo`, `Job`).

## In the port

`RuleBase::load` reads an export without `Strength_Discl` as 2026
(`rules::Dialect`); the 2026 tests are in `checks2026.rs`, the hand
analysis branches in `hand.rs`, and `bid-pbn` carries "@" tokens as the
practice program does. Self-test on Nelson's 2026 rules (16,673 test
hands), before and after:

| | 2018 reading | 2026 dialect |
|---|---|---|
| first match is the rule itself | 12,161 | 13,133 |
| an earlier rule takes the hand | 1,413 | 1,291 |
| rule rejects its own test hand | 1,681 | 840 |
| of which Losers | 791 | 19 |
| of which Suits | 666 | 377 |

What is left matches the 2026 code as written, as far as checked: test
hands saved under older code (the intermediates credit, a doubleton rule
now commented out), position specs the Editor's own test fails the same
way (it tests with the entry's prior bids, no opening passes), and 55
rules filed under "P" (one opening pass), a key the engine never looks
up because it strips opening passes.

500 random deals on BBA's 21GF-DEFAULT card (runs/random-500), IMPs from
double-dummy par per board: BBA 4.32, Rusty 4.80, BidBase 2018 rules
6.61, BidBase 2026 rules 6.39 with the card (6.44 with Nelson's own
switches); 66 games and 1 slam against BBA's 203 and 27. Along BBA's
auctions the 2026 rules fire on 60% of BBA's bids and match it 67% of
the time they do.
