# Random 500: BidBase vs BBA vs Rusty (2026-10-09)

500 random deals (dealer rotating, vulnerability rotating), each engine
bidding all four seats with the 2/1 card BBA ships as `21GF-DEFAULT`:
BBA from `21GF-DEFAULT.bbsa`, Rusty from its stock card of that name,
BidBase from `21GF-DEFAULT.card.json` (convention-card's import of the same
`.bbsa`) through `card/bidbase-map.toml`, on the 2018 rules.

Contracts are scored double-dummy against par (`rbb dd`). Results in
`report.txt`; to reproduce (from this directory, siblings checked out):

```
echo "" | dealer -p 500 -s 20261009 -f pbn -T "Random 500" > deals.pbn
"/Applications/Bridge Utilities/bba-cli" -i deals.pbn -o bba.pbn \
    --ns-conventions 21GF-DEFAULT.bbsa --ew-conventions 21GF-DEFAULT.bbsa
../../../rusty-bidding-bot/target/release/rbb bid-pbn -i deals.pbn -o rusty.pbn \
    --ns-card 21GF-DEFAULT --ew-card 21GF-DEFAULT
../../target/release/bidbase --card 21GF-DEFAULT.card.json bid-pbn -i deals.pbn -o bidbase.pbn
../../../rusty-bidding-bot/target/release/rbb dd deals.pbn > dd.jsonl
python3 -I ../../tools/compare_runs.py dd.jsonl BBA=bba.pbn Rusty=rusty.pbn BidBase=bidbase.pbn
```

## What it shows

BidBase is well behind: 6.6 IMPs from par per board against BBA's 4.3
and Rusty's 4.8, with 48 games and no slams where BBA bids 203 and 27.
The card mapping is not the cause: the database's own switches (6.65)
and every convention on (6.75) score the same.

The cause is coverage. The 2018 rules have no rule for many ordinary
hands, and the client (as the 2018 practice program) passes then: 2,580
of 3,819 calls found no rule. Examples: flat 12-HCP openers (the plain
1C/1D openings start at 13 HCP or need suit features; 62 hands of 12+
HCP passed in front of the opening), and continuations after forcing
calls. Rules were written around Nelson's test hands one auction at a
time since 2004; his current database (8 more years) may cover more.
