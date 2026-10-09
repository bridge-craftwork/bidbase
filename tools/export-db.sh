#!/bin/bash
#
# export-db.sh - dump a BidBase .accdb to CSV, one file per table.
#
# The engine reads the CSVs, not the database: they diff in git and need
# no mdbtools at run time. The .accdb stays the source of truth (Nelson
# edits it in his BidBase Editor); re-run this after taking a new copy.
#
#   tools/export-db.sh <BidBase.accdb> <out-dir>
#   tools/export-db.sh sourceforge-2018/extracted/database/BidBase.accdb data/2018

set -euo pipefail

if [[ $# -ne 2 ]]; then
    echo "usage: export-db.sh <BidBase.accdb> <out-dir>" >&2
    exit 2
fi
db=$1
out=$2
mkdir -p "$out"

# Row order matters: rules with the same Bid_Order_Number keep the order
# the database stores them in, so export rows as stored (no sorting here).
for table in Bids BidNames DupePBs Garbage; do
    mdb-export "$db" "$table" > "$out/$table.csv"
    echo "$table: $(($(wc -l < "$out/$table.csv") - 1)) lines" >&2
done
mdb-schema "$db" > "$out/schema.sql"
mdb-queries -1 "$db" | while read -r q; do
    printf -- '-- %s\n%s\n\n' "$q" "$(mdb-queries "$db" "$q")"
done > "$out/queries.sql"
mdb-ver "$db" > "$out/VERSION"
