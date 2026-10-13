#!/usr/bin/env bash
#
# Build the practice program for the browser: pkg/ (bidbase_practice.js,
# bidbase_practice_bg.wasm), an ES module that web/practice/index.html
# loads. Same pattern as crates/wasm/build.sh.
#
#   crates/practice-web/build.sh                          # release build
#   crates/practice-web/build.sh --dev                    # faster build
#   BIDBASE_PRACTICE_OUT=/path crates/practice-web/build.sh
set -euo pipefail
cd "$(dirname "$0")"

profile=--release
if [[ ${1:-} == --dev ]]; then
    profile=--dev
fi
out=${BIDBASE_PRACTICE_OUT:-pkg}

command -v wasm-pack >/dev/null || {
    echo "build.sh: wasm-pack not found: cargo install wasm-pack" >&2
    exit 1
}
rustup target list --installed 2>/dev/null | grep -q wasm32-unknown-unknown || rustup target add wasm32-unknown-unknown

# Keep the committed Cargo.lock as it was (see crates/wasm/build.sh).
lock=../../Cargo.lock
saved=$(mktemp)
cp "$lock" "$saved"
trap 'cp "$saved" "$lock"; rm -f "$saved"' EXIT

wasm-pack build --target web "$profile" --no-typescript --out-dir "$out" --out-name bidbase_practice
rm -f "$out/.gitignore" "$out/package.json" "$out/README.md"
ls -l "$out"/bidbase_practice_bg.wasm
