// The BidBase page: the engine (pkg/, built from crates/wasm) answers
// JSON requests; this file turns them into the four tabs.

import init, * as bb from './pkg/bidbase_wasm.js';

const $ = (id) => document.getElementById(id);
const call = (fn, req) => JSON.parse(fn(JSON.stringify(req)));
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

let engine = null;
let card = null; // the 2/1 card JSON, fetched once

// ----- Suits ---------------------------------------------------------------

const SUIT = { S: '♠', H: '♥', D: '♦', C: '♣' };
const RED = new Set(['H', 'D']);

/** "2H" -> "2♥" with colour; "1N"/"1NT" -> "1NT"; P/X/XX as words. */
function callHtml(c) {
  if (!c) return '';
  const t = String(c).toUpperCase();
  if (t === 'P' || t === 'PASS') return '<span class="pass">Pass</span>';
  if (t === 'D' || t === 'X') return '<span class="dbl">Dbl</span>';
  if (t === 'R' || t === 'XX') return '<span class="dbl">Rdbl</span>';
  const m = t.match(/^([1-7])(NT|N|[SHDC])/);
  if (!m) return esc(c);
  if (m[2] === 'N' || m[2] === 'NT') return `${m[1]}NT`;
  return `${m[1]}<span class="${RED.has(m[2]) ? 'red' : 'black'}">${SUIT[m[2]]}</span>`;
}

/** Suit letters inside text ("5+♠") are already symbols; colour them. */
function textHtml(s) {
  return esc(s).replace(/[♥♦]/g, (x) => `<span class="red">${x}</span>`);
}

/** A hand "AK2.KQ2.A432.J32" as four suit lines. */
function handHtml(pbn) {
  const parts = String(pbn).split(/[.-]/);
  return ['S', 'H', 'D', 'C']
    .map((s, i) => `<div><span class="${RED.has(s) ? 'red' : 'black'}">${SUIT[s]}</span> ${esc(parts[i] || '—').replace(/T/g, '10')}</div>`)
    .join('');
}

function hcp(pbn) {
  return [...String(pbn)].reduce((n, c) => n + ({ A: 4, K: 3, Q: 2, J: 1 }[c] || 0), 0);
}

// ----- Engine ---------------------------------------------------------------

async function makeEngine() {
  const rules = $('rules').value;
  const convs = $('convs').value;
  const req = { rules };
  if (convs === 'all') req.all_conventions = true;
  if (convs === 'card') {
    if (!card) card = await (await fetch('cards/21GF-DEFAULT.card.json')).json();
    req.card = card;
  }
  if (engine) call(bb.freeEngine, { engine });
  const r = call(bb.createEngine, req);
  if (!r.ok) {
    $('status').textContent = r.error;
    return;
  }
  engine = r.engine;
  const nt = r.nt.one[0] ? `, 1NT ${r.nt.one[0]}-${r.nt.one[1]}` : '';
  $('status').textContent = `${r.rules.toLocaleString()} rules, ${r.dialect} engine${nt}`;
}

// ----- Bid a hand -------------------------------------------------------------

function bidHand() {
  const req = { engine, hand: $('hand').value, auction: $('auction').value, vul: $('vul').value };
  const r = call(bb.bid, req);
  saveHash();
  const out = $('hand-out');
  if (!r.ok) {
    out.innerHTML = `<p class="error">${esc(r.error)}</p>`;
    return;
  }
  let html = '';
  if (r.call) {
    html += `<div class="answer"><div class="big">${callHtml(r.call)}</div>
      <div><div class="meaning">${textHtml(r.meaning)}</div>
      <div class="hint">by <a href="#" data-rule="${r.rule}">rule #${r.rule}</a>, in the section for “${esc(r.key || 'opening')}”</div></div></div>`;
    if (r.meanings.length > 1) {
      html += `<h3>${callHtml(r.call)} can show</h3><p class="hint">Every enabled rule here that makes this call: what the table may know.</p><ul class="meanings">`;
      for (const m of r.meanings) html += `<li><a href="#" data-rule="${m.rule}">#${m.rule}</a> ${textHtml(m.meaning)}</li>`;
      html += '</ul>';
    }
  } else {
    html += `<div class="answer"><div class="big"><span class="pass">Pass</span></div><div><div class="meaning">No rule fits: ${esc(r.error)}</div>
      <div class="hint">BidBase passes when no rule fits the hand.</div></div></div>`;
  }
  if (r.tried.length) {
    html += `<h3>Rules tried, in order</h3><table class="trace"><thead><tr><th>Rule</th><th>Order</th><th>Call</th><th>Result</th></tr></thead><tbody>`;
    for (const t of r.tried) {
      const res = t.test ? `<span class="fail">${esc(t.test)}</span> ${textHtml(t.why)}` : '<span class="ok">matches</span>';
      html += `<tr class="${t.test ? '' : 'hit'}"><td><a href="#" data-rule="${t.rule}">#${t.rule}</a></td><td>${esc(t.bon)}</td><td>${callHtml(t.bid)}</td><td>${res}</td></tr>`;
    }
    html += '</tbody></table>';
  }
  out.innerHTML = html;
}

// ----- Bid a deal ----------------------------------------------------------------

let deal = null;

function newDeal() {
  const seed = Math.floor(Math.random() * 2 ** 31);
  deal = call(bb.deal, { seed }).hands;
  bidDeal();
}

function bidDeal() {
  if (!deal) return;
  const dealer = $('dealer').value;
  const r = call(bb.bidDeal, { engine, hands: deal, dealer, vul: $('deal-vul').value });
  const out = $('deal-out');
  if (!r.ok) {
    out.innerHTML = `<p class="error">${esc(r.error)}</p>`;
    return;
  }
  const seat = (s) => `<div class="seat seat-${s}"><div class="seat-name">${{ N: 'North', E: 'East', S: 'South', W: 'West' }[s]} <span class="hint">${hcp(deal[s])} HCP</span></div>${handHtml(deal[s])}</div>`;
  let html = `<div class="table-wrap"><div class="compass">${seat('N')}${seat('W')}<div class="seat center"></div>${seat('E')}${seat('S')}</div><div class="auction-wrap">`;
  const order = ['W', 'N', 'E', 'S'];
  html += '<table class="auction"><thead><tr>' + order.map((s) => `<th>${s}</th>`).join('') + '</tr></thead><tbody><tr>';
  let col = order.indexOf(dealer);
  for (let i = 0; i < col; i++) html += '<td></td>';
  r.calls.forEach((c, i) => {
    if (col === 4) {
      html += '</tr><tr>';
      col = 0;
    }
    const cls = c.missing ? 'nocall' : '';
    html += `<td class="${cls}"><button class="callbtn" data-i="${i}">${callHtml(c.call)}</button></td>`;
    col++;
  });
  html += '</tr></tbody></table>';
  html += `<p class="contract">${r.contract === 'Pass' ? 'Passed out' : `Contract: ${callHtml(r.contract.replace(/X+$/, ''))}${esc(r.contract.match(/X+$/)?.[0] || '')} by ${esc(r.declarer)}`}</p>`;
  html += '<div id="call-meaning" class="hint">Click a call for its meaning. A call in grey is a pass for want of a rule.</div></div></div>';
  out.innerHTML = html;
  out.querySelectorAll('.callbtn').forEach((b) =>
    b.addEventListener('click', () => {
      const c = r.calls[+b.dataset.i];
      $('call-meaning').innerHTML = c.rule
        ? `${callHtml(c.call)} by ${esc(c.seat)}: ${textHtml(c.meaning)} <a href="#" data-rule="${c.rule}">rule #${c.rule}</a>`
        : `${callHtml(c.call)} by ${esc(c.seat)}: ${esc(c.missing || '')}`;
    }));
}

// ----- Rules -----------------------------------------------------------------------

function showRules() {
  const r = call(bb.section, { engine, auction: $('rules-auction').value });
  const out = $('rules-out');
  if (!r.ok) {
    out.innerHTML = `<p class="error">${esc(r.error)}</p>`;
    return;
  }
  if (!r.rules.length) {
    out.innerHTML = `<p>No rules for “${esc(r.key || 'opening')}”.</p>`;
    return;
  }
  let html = `<p class="hint">${r.rules.length} rules for “${esc(r.key || 'opening')}”.</p><table class="trace"><thead><tr><th>Rule</th><th>Order</th><th>Call</th><th>Meaning</th><th>Test hand</th></tr></thead><tbody>`;
  for (const x of r.rules) {
    html += `<tr class="${x.enabled ? '' : 'off'}"><td><a href="#" data-rule="${x.id}">#${x.id}</a></td><td>${esc(x.bon)}</td><td>${callHtml(x.bid)}</td><td>${textHtml(x.meaning)}</td><td class="mono">${esc(x.test_hand)}</td></tr>`;
  }
  out.innerHTML = html + '</tbody></table>';
}

// ----- Conventions -----------------------------------------------------------------

function showConventions() {
  const r = call(bb.conventions, { engine });
  if (!r.ok) return;
  const f = $('conv-filter').value.trim().toLowerCase();
  const list = r.conventions.filter((c) => !f || c.name.toLowerCase().includes(f));
  let html = `<p class="hint">${list.length} of ${r.conventions.length} conventions that rules use.</p><table class="trace"><thead><tr><th>On</th><th>Convention</th><th>Rules</th></tr></thead><tbody>`;
  for (const c of list) {
    html += `<tr><td><input type="checkbox" data-conv="${c.id}" ${c.on ? 'checked' : ''}></td><td>${esc(c.name || '(unnamed)')}</td><td>${c.rules}</td></tr>`;
  }
  $('conv-out').innerHTML = html + '</tbody></table>';
}

// ----- Rule detail -------------------------------------------------------------------

function showRule(id) {
  const r = call(bb.rule, { engine, id });
  const box = $('rule-detail');
  if (!r.ok) {
    box.innerHTML = `<p class="error">${esc(r.error)}</p>`;
  } else {
    const rows = Object.entries(r.rule).map(([k, v]) => `<tr><th>${esc(k)}</th><td>${esc(v)}</td></tr>`).join('');
    const conv = r.convention ? `<p>Convention: ${esc(r.convention.name)} ${r.convention.selected || r.convention.sub_selected ? '(switched off in Nelson\'s settings)' : ''}</p>` : '';
    box.innerHTML = `<div class="row spread"><h3>Rule #${id}</h3><button id="close-rule">Close</button></div>
      <p class="meaning">${callHtml(r.rule.Bid)} ${textHtml(r.meaning)}</p>${conv}
      <table class="fields">${rows}</table>`;
    $('close-rule').addEventListener('click', () => (box.hidden = true));
  }
  box.hidden = false;
  box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// ----- Tabs, links -------------------------------------------------------------------

function selectTab(name) {
  document.querySelectorAll('[data-tab]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === name)));
  document.querySelectorAll('.tab').forEach((s) => (s.hidden = s.dataset.for !== name));
  if (name === 'conventions') showConventions();
  if (name === 'rules') showRules();
  if (name === 'deal' && !deal) newDeal();
}

/** The hand tab's inputs live in the fragment, so a link shows the same bid. */
function saveHash() {
  const p = new URLSearchParams({ hand: $('hand').value, auction: $('auction').value, vul: $('vul').value, rules: $('rules').value, convs: $('convs').value });
  history.replaceState(null, '', '#' + p.toString());
}

function loadHash() {
  const p = new URLSearchParams(location.hash.slice(1));
  for (const k of ['hand', 'auction', 'vul', 'rules', 'convs']) if (p.has(k)) $(k).value = p.get(k);
  return p.has('hand');
}

async function main() {
  await init();
  const fromLink = loadHash();
  await makeEngine();
  document.querySelectorAll('[data-tab]').forEach((b) => b.addEventListener('click', () => selectTab(b.dataset.tab)));
  document.querySelectorAll('[data-tab-link]').forEach((a) =>
    a.addEventListener('click', (e) => {
      e.preventDefault();
      selectTab(a.dataset.tabLink);
    }));
  $('hand-form').addEventListener('submit', (e) => {
    e.preventDefault();
    bidHand();
  });
  $('rules-form').addEventListener('submit', (e) => {
    e.preventDefault();
    showRules();
  });
  $('new-deal').addEventListener('click', newDeal);
  $('dealer').addEventListener('change', bidDeal);
  $('deal-vul').addEventListener('change', bidDeal);
  $('conv-filter').addEventListener('input', showConventions);
  for (const id of ['rules', 'convs']) {
    $(id).addEventListener('change', async () => {
      await makeEngine();
      bidHand();
      if (deal) bidDeal();
      if (!$('rules-out').closest('.tab').hidden) showRules();
      if (!$('conv-out').closest('.tab').hidden) showConventions();
    });
  }
  document.addEventListener('change', (e) => {
    const id = e.target.dataset?.conv;
    if (id) {
      call(bb.setConvention, { engine, id: +id, on: e.target.checked });
      bidHand();
    }
  });
  document.addEventListener('click', (e) => {
    const a = e.target.closest('[data-rule]');
    if (a) {
      e.preventDefault();
      showRule(+a.dataset.rule);
    }
  });
  bidHand();
  if (!fromLink) saveHash();
}

main().catch((e) => {
  $('status').textContent = `Could not start: ${e}`;
});
