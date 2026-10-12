#!/usr/bin/env bash
#
# Build the BidBase page into web/dist/: the static files, the engine
# (crates/wasm/build.sh into dist/pkg, with the 2018 and 2026 rules compiled
# in) and the example convention card. dist/ is what the Cloudflare Pages
# project serves (wrangler.jsonc); bridge-craftwork.com mounts it at
# /bidbase/.
#
#   web/build.sh          # release build
#   web/build.sh --dev    # faster engine build
#
# Then: python3 -m http.server -d web/dist 8000, and open
# http://localhost:8000/ (WASM does not load from file://).
# Deploy: npx --yes wrangler@4 pages deploy (from the repo root).
set -euo pipefail
here=$(cd "$(dirname "$0")" && pwd)
root=$(cd "$here/.." && pwd)
dist="$here/dist"

rm -rf "$dist"
mkdir -p "$dist/cards"
BIDBASE_WASM_OUT="$dist/pkg" "$root/crates/wasm/build.sh" "$@"
rm -f "$dist/pkg/.gitignore" "$dist/pkg/package.json" "$dist/pkg/README.md"
cp "$here"/{index.html,app.js,app.css,styles.css,favicon.svg,_headers} "$dist/"
cp "$root/card/examples/21GF-DEFAULT.card.json" "$dist/cards/"
ls -l "$dist" "$dist/pkg"
