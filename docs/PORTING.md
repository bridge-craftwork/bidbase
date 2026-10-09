# Porting notes: BidBaseDB_Bidder.cls to Rust

What `crates/engine` takes from the 2018 VB6 engine
(`sourceforge-2018/CardShark BidBase Bidding Module/BidBaseDB_Bidder.cls`),
what it leaves out, and where it does what the code meant rather than what
it did. Line numbers are in that file.

## Ported

| VB | Rust |
|---|---|
| `LookUpBid` (570) | `lookup.rs` `Bidder::lookup` |
| `Calculate_BidNum`, `Calculate_Position` (868, 1842) | `auction.rs` |
| `ConvertHandStringToArrays`, `Calculate_Winners`, `_LosingTrickCount`, `_QuickTricks`, `_Stoppers`, `_Shape` (1003-1839) | `hand.rs` |
| `GetMinNTpts` (1917) | `rules.rs` `find_nt_ranges` |
| `PctUsed_Ok` (2023) | `lookup.rs` `pct_used` |
| `TotalHCPs_Ok` ... `SuitHCP_Ok` (2184-4091) | `checks.rs` |
| `GetDataFromMatchingEntry`, `GetDiscl`, `AddToDisclosure` (4171-4480) | `lookup.rs` |
| The Editor's entry check (`TestID`, `TestEntry`) | `selftest.rs`, `bidbase selftest` |

## Not ported

- **Hand valuation** (`Eval_*`, 4824-6701, and the BB-HEP database). It
  runs only when a rule names a `ValSystem` and the points to reach
  (`SuitVal`/`NT_Val`): 3 of 6,314 rules. Those 3 fail with "valuation
  system not ported". `ValSystem` alone (45 more rules) is ignored, as in
  the VB.
- **The analysis form** (`ShowAnalysis`, `frmBidBaseAnalysis`). Its "NO
  MATCH" lines come back as `Tried` records instead (`bidbase bid --why`).
- **`GetOtherPlayersData`** (4509): the VB skips it on the bidding path
  (`GoTo SkipPard`, 1498), so LHO's suit, partner's suits and fits are
  never set. See "LHO's suit" below.
- **`DupePBs`** and the auction shortcuts in Nelson's notes (`1M-P`, `@`):
  neither the 2018 engine nor the 2018 rules use them. The rules key
  slam replies by pseudo-auctions instead ("Blackwood: 4N-P", "RKCB:
  4N-P", "Gerber: 4C-P"); a client has to ask for those by name.
- The trump-suit `InputBox` (3324): the trump suit comes from the rule's
  `Source` ("<H..."), else from the request.

## Where the port departs from the 2018 code

Each is a case where the code contradicts its own comments, the rule
data, or both. The self-test (below) was the check.

1. **Notrump ranges** (`GetMinNTpts`, 1945). The VB looks up 1NT at
   `Bid_Order_Number` "130000" and 2NT at "110000"; by 2018 those are the
   25-27 2C/3NT rule and the 15-17 1NT rule (Nelson left a "Don't
   understand" `Stop` there). The port takes the first enabled 1NT and 2NT
   openings: 15-17 and 20-21.
2. **"P" HCP specs** (2243). Responses to 1NT say "P24-25" (partnership
   total). The VB lower-cases the spec and then tests for "P", so the
   conversion never ran and "P24-25" read as 0-25. Ported as meant (9-10
   opposite 15-17). 59 rules.
3. **">n" HCP** (2280) read as 0-40 (`Val(">12")` is 0); now 13+.
   **"hcp2n"** was tested against the bid, not the spec (2217).
4. **Ace, king and key-card replies.** HCP specs "A2", "K1", "KC2+Q" and
   `cmHCP` specs like "2A+TQ,1A+TK+TQ" are new to the port. The VB read
   the first as numbers (so they matched only 0-HCP hands) and the second
   only behind an "SB:" prefix that no row has (so every hand passed).
   Semantics in `checks.rs` `slam_spec`.
5. **Intermediates credit** (2319, 2579). Extra HCP (or total points)
   over a rule's minimum count as intermediates. The VB added the credit
   to a module variable, so it grew with every rule tested. Now per rule.
6. **Seats** (`Calculate_Position`). The VB took the bidder's seat for
   F/T and Y/N from the seat counted from the opener, so every opener was
   "F" (third-seat "Rule of 20" openings passed the `F` test) and a third-
   seat overcaller was never "T". The port counts seats from the dealer,
   which is what the rules' notes describe. Also `str_Call(3)` was never
   set after "P-P-P" (919).
7. **Vulnerability** (`Vul`, 303, and `Vul_Ok`, 2846). The VB mapped rule
   codes to letters ("Eq" to "B", "Y" to "U", ...) that the client's
   vulnerability strings mostly don't contain: "Eq" never matched and
   "Y" failed when both sides were vulnerable. The port matches the codes
   themselves: we vul = {Y, Unfav}; they = {N, Fav, Eq+}; neither = {Eq, N,
   Eq+}; both = {Eq, Y, Eq+}.
8. **Count fields** (Winrs, Losers, QTs, Stopped_Suits). Kept the VB
   reading, "any value not preceded by '<' is considered a minimum" (2427),
   so "2-3" losers is 2 or more; a test hand (#3285) depends on it. "<n"
   now means fewer than n in all four (the VB did so only for Losers).
9. **Suit specs.** `pts2` is reset for each suit (it leaked from the
   previous suit, 3011); every "<x"/">x" comparison is checked, not just
   the last (3122); ">=n" reads the number after ">=" (3679); ">H" and
   ">n" in a suit's HCP field are new (the data has ">H").
10. **Combined suits.** An "x" anywhere in the suit letters ("Sx>11")
    stands for the x suit (only a leading x was read, 3459); an "=" range
    in `cmHCP` reads its upper bound from `cmHCP`, not `cmNum` (3491).
11. **"2x" with no x suit** (4230). Twelve rules bid "Nx" without any spec
    that picks the suit; the VB returned the bare level ("5"). The port
    fails such a rule with "names no suit for x".

## Kept as in 2018, on purpose

- **LHO's suit** stays unknown, so `Calculate_Stoppers` never takes the
  LHO discount (Kx, KJx, Qxx...). Measured on the test hands: unknown
  (as 2018) 54 rules reject their own hand; LHO's first natural suit 59;
  RHO's 64. The test hands were made with the 2018 behaviour.
- A singleton king counts no loser (the comment says 1; the code adds 0).
- `Pct_Used` percentages: one draw (1-100) over the run of numbered rules
  sharing a `Bid_Order_Number`; the last of the run takes what is left.
  The generator is seeded (`--seed`), not the clock.

## Self-test (2018 data)

`bidbase selftest` runs each rule's `Test_Hand` the way the Editor does
when saving an entry: the rule alone, then the whole auction.

| | |
|---|---|
| rules (excluding 849 "ZZZZZZ" section ends) | 5,465 |
| with a test hand | 5,258 |
| first match is the rule itself | 3,440 |
| switched off, but accepts its test hand | 1,454 |
| an earlier rule takes the hand | 289 |
| rule rejects its own test hand | 66 |
| test hand unreadable (typos like "K33") | 9 |

`--list own|first|all` names them. The rejects are mostly 2004 entries
whose test hands don't fit their own specs (#478, a 2D rebid after
1D-2C, has a test hand with 9 HCP and diamonds 432), the 12 x-bids above, 3 valuation-system rules and 5 "0stop"
cases that want Qxx in the overcaller's suit not to count.

**Switched-off conventions.** `BidNames.Selected`/`SubCat_Selected` are
Nelson's settings, stored in the database: in the 2018 copy 1,424 rules
belong to conventions switched off, including every natural notrump
response and every four-card 1-over-1 response to 1C, so 1C-P with
K32-KJ32-A32-Q32 finds no call. `--all-conventions` ignores the switches
(then 1,082 test hands are taken by an earlier rule, as competing
conventions overlap). A sensible profile for comparisons with BBA and
Rusty is still to be chosen.
