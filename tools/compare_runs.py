#!/usr/bin/env python3
"""Compare bidding engines on the same deals, scored double-dummy.

    python3 -I tools/compare_runs.py DD.jsonl NAME=FILE.pbn NAME=FILE.pbn ...

DD.jsonl is `rbb dd` output for the deals (tricks per declarer N E S W and
strain C D H S NT, `par_ns`). Each FILE.pbn is one engine's auctions, in
bba-cli's layout (`[Deal]`, `[Vulnerable]`, `[Contract]`, `[Declarer]`,
`[Auction]`). The first engine is the reference for agreement figures.

Every contract is scored as if played double-dummy. "From par" is the IMP
distance between that score and par (always >= 0, lower is better); on a
board where two engines reach different contracts, the one nearer par
wins the difference, as rusty-bidding-bot's "vs BBA" measures it.
"""

import json
import re
import sys
from collections import Counter

IMP_STEPS = [20, 50, 90, 130, 170, 220, 270, 320, 370, 430, 500, 600, 750, 900,
             1100, 1300, 1500, 1750, 2000, 2250, 2500, 3000, 3500, 4000]
STRAINS = "CDHSN"


def imps(diff):
    n = sum(1 for step in IMP_STEPS if abs(diff) >= step)
    return n if diff >= 0 else -n


def score(contract, declarer, tricks, vul):
    """Duplicate score for the declaring side, made or set by `tricks`."""
    level = int(contract[0])
    strain = contract[1]
    doubled = contract.count("X")
    vulnerable = vul in ("All", "Both") or (vul == "NS" and declarer in "NS") or (vul == "EW" and declarer in "EW")
    need = level + 6
    if tricks >= need:
        per = 20 if strain in "CD" else 30
        base = per * level + (10 if strain == "N" else 0)
        base *= (1, 2, 4)[doubled]
        s = base
        s += (500 if vulnerable else 300) if base >= 100 else 50
        if level == 6:
            s += 750 if vulnerable else 500
        if level == 7:
            s += 1500 if vulnerable else 1000
        over = tricks - need
        if doubled:
            s += 50 * doubled
            s += over * (200 if vulnerable else 100) * doubled
        else:
            s += over * per
        return s
    down = need - tricks
    if not doubled:
        return -down * (100 if vulnerable else 50)
    if vulnerable:
        pen = 200 + 300 * (down - 1)
    else:
        pen = 100 + 200 * min(down - 1, 2) + 300 * max(down - 3, 0)
    return -pen * doubled


def read_pbn(path):
    """Boards keyed by their [Deal]: tags, plus the auction's calls."""
    boards = {}
    cur = {"calls": []}
    in_auction = False

    def close():
        if "Deal" in cur:
            boards[cur["Deal"]] = dict(cur)

    for line in open(path, encoding="utf-8", errors="replace"):
        line = line.rstrip("\n")
        if not line.strip():
            close()
            cur = {"calls": []}
            in_auction = False
            continue
        m = re.match(r'^\[(\w+) "(.*)"\]', line)
        if m:
            tag, val = m.groups()
            in_auction = tag == "Auction"
            cur[tag] = val
        elif in_auction and not line.startswith(("%", "{")):
            cur["calls"] += [t for t in line.split() if not t.startswith("=")]
    close()
    return boards


def ns_score(board, tricks_by_decl):
    contract = board.get("Contract", "Pass")
    if contract in ("Pass", ""):
        return 0, "Pass"
    declarer = board["Declarer"]
    t = tricks_by_decl[declarer][STRAINS.index(contract[1])]
    s = score(contract, declarer, t, board["Vulnerable"])
    return (s if declarer in "NS" else -s), f"{contract}{declarer}"


def main():
    dd_path, *runs = sys.argv[1:]
    dd = {}
    for line in open(dd_path):
        rec = json.loads(line)
        dd[rec["deal"]] = rec
    names = [r.split("=", 1)[0] for r in runs]
    data = {r.split("=", 1)[0]: read_pbn(r.split("=", 1)[1]) for r in runs}
    deals = [d for d in data[names[0]] if d in dd and all(d in data[n] for n in names)]
    ref = names[0]
    dist = {n: [] for n in names}
    contracts = {n: [] for n in names}
    level_counts = {n: Counter() for n in names}
    same_auction = Counter()
    same_contract = Counter()
    for deal in deals:
        rec = dd[deal]
        for n in names:
            b = data[n][deal]
            s, c = ns_score(b, rec["tricks"])
            dist[n].append(abs(imps(s - rec["par_ns"])))
            contracts[n].append(c)
            if c == "Pass":
                level_counts[n]["passed out"] += 1
            else:
                lvl = int(c[0])
                strain = c[1]
                game = (strain == "N" and lvl >= 3) or (strain in "HS" and lvl >= 4) or (strain in "CD" and lvl >= 5)
                level_counts[n]["slam" if lvl >= 6 else "game" if game else "partscore"] += 1
            if n != ref:
                same_auction[n] += b["calls"] == data[ref][deal]["calls"]
                same_contract[n] += c == contracts[ref][-1]
    total = len(deals)
    print(f"{total} deals, scored double-dummy against par.\n")
    print(f"{'engine':<10} {'IMPs from par':>14} {'at par':>7} {'passed out':>11} {'partscore':>10} {'game':>6} {'slam':>6}")
    for n in names:
        d = dist[n]
        c = level_counts[n]
        print(f"{n:<10} {sum(d) / total:>9.2f}/brd {sum(1 for x in d if x == 0) / total:>7.0%} "
              f"{c['passed out']:>11} {c['partscore']:>10} {c['game']:>6} {c['slam']:>6}")
    print(f"\nAgreement with {ref}:")
    for n in names[1:]:
        print(f"  {n:<10} same auction {same_auction[n] / total:>4.0%}   same contract and declarer {same_contract[n] / total:>4.0%}")
    print("\nHead to head (IMPs nearer par, on boards where the contracts differ):")
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            diff = [dist[b][k] - dist[a][k] for k in range(total) if contracts[a][k] != contracts[b][k]]
            boards = len(diff)
            print(f"  {a} vs {b}: {sum(diff):+d} IMPs to {a} over {boards} boards "
                  f"({sum(diff) / max(boards, 1):+.2f} per differing board)")


if __name__ == "__main__":
    main()
