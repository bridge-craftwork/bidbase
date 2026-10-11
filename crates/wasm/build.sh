#!/usr/bin/env bash
#
# Build the browser package: crates/wasm/pkg/ (bidbase_wasm.js,
# bidbase_wasm_bg.wasm, .d.ts, package.json), an ES module for
# `<script type="module">`. The 2026 and 2018 rules are compiled in.
#
#   crates/wasm/build.sh                         # release build
#   crates/wasm/build.sh --dev                   # faster build, bigger wasm
#   BIDBASE_WASM_OUT=/path crates/wasm/build.sh  # write the package elsewhere
#
# Needs: rustup target add wasm32-unknown-unknown; cargo install wasm-pack.
# Copied from rusty-bidding-bot's crates/wasm/build.sh.
set -euo pipefail
cd "$(dirname "$0")"

profile=--release
if [[ ${1:-} == --dev ]]; then
    profile=--dev
fi
out=${BIDBASE_WASM_OUT:-pkg}

command -v wasm-pack >/dev/null || {
    echo "build.sh: wasm-pack not found: cargo install wasm-pack" >&2
    exit 1
}
rustup target list --installed 2>/dev/null | grep -q wasm32-unknown-unknown || {
    echo "build.sh: installing the wasm32-unknown-unknown target" >&2
    rustup target add wasm32-unknown-unknown
}

# wasm-pack runs cargo itself, so the gitignored [patch] overrides in
# .cargo/config.toml (local sibling checkouts; dev-build.sh) would make
# cargo rewrite the committed Cargo.lock with local-path entries. Keep the
# lock as it was before the build.
lock=../../Cargo.lock
saved=$(mktemp)
cp "$lock" "$saved"
trap 'cp "$saved" "$lock"; rm -f "$saved"' EXIT

wasm-pack build --target web "$profile" --out-dir "$out" --out-name bidbase_wasm
ls -l "$out"/bidbase_wasm_bg.wasm
