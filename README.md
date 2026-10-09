# bidbase

A Rust port of the bidding engine of Nelson Ford's
[CardShark BidBase](http://www.aeyec.com/bidbase/), a rule-based bridge
bidder whose knowledge lives in an Access database of about 6,300 rules,
together with the September 2018
[SourceForge release](https://sourceforge.net/projects/bidbase/) it is
ported from. BidBase is public domain.

```
./dev-build.sh run -q -p bidbase-cli -- bid AK2-KQ2-A432-J32
1N  (rule #13)
./dev-build.sh run -q -p bidbase-cli -- bid Q8432-32-K32-J32 -a "1NT Pass" --why
./dev-build.sh run -q -p bidbase-cli -- selftest
```

- `crates/engine`: the engine (`docs/PORTING.md` for what was ported and
  where it departs from the VB6 code)
- `crates/cli`: the `bidbase` command
- `data/2018`: the rules, exported to CSV with `tools/export-db.sh`
- `sourceforge-2018`: the original release, untouched

Build with `./dev-build.sh` (see CLAUDE.md).
