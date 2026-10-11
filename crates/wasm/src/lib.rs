//! Browser (WASM) boundary for the BidBase port. Each export takes a JSON
//! request string and returns a JSON response string; the logic is in
//! `api` (plain Rust, tested natively) and this file only binds it.

pub mod api;

use wasm_bindgen::prelude::*;

#[wasm_bindgen(start)]
pub fn start() {
    #[cfg(feature = "console_error_panic_hook")]
    console_error_panic_hook::set_once();
}

/// `{ok, api, version, rule_sets}`
#[wasm_bindgen]
pub fn info() -> String {
    api::info()
}

/// `{rules?, card?, all_conventions?, seed?}` → `{ok, engine, source, dialect, rules, conventions, nt, card}`
#[wasm_bindgen(js_name = createEngine)]
pub fn create_engine(request: &str) -> String {
    api::create_engine(request)
}

/// `{engine}` → `{ok}`
#[wasm_bindgen(js_name = freeEngine)]
pub fn free_engine(request: &str) -> String {
    api::free_engine(request)
}

/// `{engine, hand, auction?, vul?, trump?, scoring?}` → `{ok, call, rule, meaning, meanings, key, tried, error?}`
#[wasm_bindgen]
pub fn bid(request: &str) -> String {
    api::bid(request)
}

/// `{engine, id}` → `{ok, rule, convention, meaning}`
#[wasm_bindgen]
pub fn rule(request: &str) -> String {
    api::rule(request)
}

/// `{engine, auction}` → `{ok, key, rules}`
#[wasm_bindgen]
pub fn section(request: &str) -> String {
    api::section(request)
}

/// `{engine}` → `{ok, source, conventions}`
#[wasm_bindgen]
pub fn conventions(request: &str) -> String {
    api::conventions(request)
}

/// `{engine, id, on}` → `{ok}`
#[wasm_bindgen(js_name = setConvention)]
pub fn set_convention(request: &str) -> String {
    api::set_convention(request)
}

/// `{engine, hands: {N, E, S, W}, dealer?, vul?}` → `{ok, calls, contract, declarer}`
#[wasm_bindgen(js_name = bidDeal)]
pub fn bid_deal(request: &str) -> String {
    api::bid_deal(request)
}

/// `{seed?}` → `{ok, hands}`
#[wasm_bindgen]
pub fn deal(request: &str) -> String {
    api::deal(request)
}
