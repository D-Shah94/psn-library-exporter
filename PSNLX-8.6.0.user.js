// ==UserScript==
// @name         PSN Library Exporter (PSNLX)
// @namespace    https://github.com/D-Shah94/psn-library-exporter
// @version      8.6.0
// @description  Reads the PlayStation Library (Purchased tab) and Transaction History pages that YOU have open and saves CSV files. No API calls, no tokens, no credentials.
// @match        https://library.playstation.com/*
// @match        https://checkout.playstation.com/*
// @run-at       document-idle
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @grant        GM_addValueChangeListener
// ==/UserScript==

/*
  PRIVACY: this script never reads, stores or sends cookies, tokens, NPSSO codes, e-mail, card numbers or device IDs.
  The account is identified automatically from the page's own session data, and ONLY two fields are kept from it:
  your PSN online ID (a public handle) and your country. Everything else in that data (e-mail, account IDs, real name...) is ignored.
  Expanded order text is cleaned before it is stored: e-mail addresses and card digits are removed.
  Nothing is sent anywhere. CSVs contain your game list, purchase history and your online ID - don't publish them.
  It adds no click handlers to the page and does nothing until you press one of its buttons.
  Stored data lives in Tampermonkey's own per-script storage, shared by library.playstation.com and checkout.playstation.com (falls back to page storage if GM_* is unavailable).
*/
(function () {
  'use strict';
  const CFG = {
    minDelayMs: 2200, maxDelayMs: 4000, pageTimeoutMs: 25000, controlWaitMs: 8000, maxPages: 300,
    txPatienceMs: 6000, txPatienceStepMs: 4000, txStagnantLimit: 5, txMaxRounds: 3000, txFastMinMs: 60, txFastMaxMs: 160,
    txDetailMinMs: 500, txDetailMaxMs: 1000,
    // storage keys are unchanged since 8.4.0, so existing stored data keeps working
    keyLib: 'psnlx_v84_lib', keyTx: 'psnlx_v84_tx', keyUi: 'psnlx_v84_ui', keyScans: 'psnlx_v84_scans', keyIdent: 'psnlx_v84_ident', keyBusy: 'psnlx_v84_busy', keyCheckpoint: 'psnlx_v84_checkpoint'
  };
  const TAB = 'recently-purchased';
  const SYNC_KEYS = [CFG.keyTx, CFG.keyScans, CFG.keyIdent];
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const rnd = (a, b) => a + Math.random() * (b - a);
  const hhmmss = () => new Date().toLocaleTimeString('en-GB');
  const today = () => new Date().toISOString().slice(0, 10);
  const regionFromLocale = loc => { const m = /([a-z]{2})[-_]([a-z]{2})/i.exec(loc || ''); if (!m) return ''; const c = m[2].toUpperCase(); return c === 'GB' ? 'UK' : c; };
  const csvCell = v => { v = (v === undefined || v === null) ? '' : String(v); return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  const toCSV = (cols, rows) => [cols.map(c => c[0]).join(',')].concat(rows.map(r => cols.map(c => csvCell(r[c[1]])).join(','))).join('\r\n');
  const slug = s => (s || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 24);
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  function download(name, text) {
    const blob = new Blob(['\ufeff', text], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }
  // Text of the PAGE only: our own panel is excluded (its text quotes the site's messages) and so are script/style blocks.
  const pageText = () => [...document.body.children].filter(c => !c.hasAttribute('data-psnlx') && !/^(SCRIPT|STYLE|NOSCRIPT|TEMPLATE)$/.test(c.tagName)).map(c => c.innerText || '').join('\n');
  async function waitFor(fn, ms, step = 250) { const t0 = Date.now(); while (Date.now() - t0 < ms) { if (fn()) return true; await sleep(step); } return false; }

  // ---------- privacy helpers ----------
  const EMAIL_RE = /[\w.+-]+@[\w-]+(?:\.[\w-]+)+/g, CARD_RE = /\(\s*\*{2,}\s*\d{2,4}\s*\)/g;
  function cleanDetails(text) {            // keep payment + item lines, drop the ACCOUNT / e-mail line and card digits
    let lines = String(text || '').split('\n').map(x => x.replace(/\u00a0/g, ' ').trim()).filter(Boolean);
    const i = lines.findIndex(l => /^payment method$/i.test(l)); if (i >= 0) lines = lines.slice(i);
    return lines.join(' | ').replace(EMAIL_RE, '[email]').replace(CARD_RE, '(****)').slice(0, 3000);
  }
  // Identity: ONLY the online ID and country are ever kept. (Session data also holds e-mail, account IDs, real name... all ignored.)
  const HDR_ID = '[data-qa$="user-profile-details#playerName2"] .psw-player-name-text';
  const IDENT_MAX_AGE_MS = 120 * 60 * 1000;
  function ownIdentity() {
    let account = '', region = '';
    try {
      const el = document.getElementById('__NEXT_DATA__'); if (el) {
        const u = (((JSON.parse(el.textContent).props || {}).appProps || {}).session || {}).userData || {};
        account = String(u.onlineId || '').trim(); const c = String(u.country || '').toUpperCase(); region = c === 'GB' ? 'UK' : c;
      }
    } catch (e) {}
    if (!account) { const h = document.querySelector(HDR_ID); if (h) account = h.textContent.trim(); }
    return { account, region };
  }
  let inFrame = true; try { inFrame = window.self !== window.top; } catch (e) { inFrame = true; }
  const isDrawer = location.hostname === 'checkout.playstation.com';

  // ---------- storage: Tampermonkey GM storage (shared by both sites); page storage + postMessage only as a fallback ----------
  const HAS_GM = typeof GM_getValue === 'function' && typeof GM_setValue === 'function';
  const store = {
    mem: {},
    load(k) {
      if (HAS_GM) { try { return JSON.parse(GM_getValue(k, '{}') || '{}'); } catch (e) { return {}; } }
      let a = {}, b = {};
      try { a = JSON.parse(localStorage.getItem(k) || '{}'); } catch (e) {}
      try { b = JSON.parse(this.mem[k] || '{}'); } catch (e) {}
      return Object.assign({}, a, b);
    },
    save(k, v, silent) {
      const s = JSON.stringify(v);
      if (HAS_GM) { try { GM_setValue(k, s); } catch (e) {} }
      else {
        try { localStorage.setItem(k, s); } catch (e) {}
        this.mem[k] = s;
        if (!silent && isDrawer && inFrame && SYNC_KEYS.includes(k)) { try { window.parent.postMessage({ psnlx: 1, type: 'sync', key: k, value: v }, '*'); } catch (e) {} }
      }
    },
    onChange(fn) {                                  // another tab / the other panel changed the data
      if (typeof GM_addValueChangeListener !== 'function') return;
      let t = null; [CFG.keyLib, CFG.keyTx, CFG.keyScans, CFG.keyIdent].forEach(k => { try { GM_addValueChangeListener(k, (n, o, nv, remote) => { if (remote === false) return; clearTimeout(t); t = setTimeout(fn, 150); }); } catch (e) {} });
    }
  };
  function rememberIdentity(id) {
    if (!id.account) return; const s = store.load(CFG.keyIdent);
    if (s.account !== id.account || s.region !== id.region || !s.at || Date.now() - Date.parse(s.at) > 60000) store.save(CFG.keyIdent, { account: id.account, region: id.region, at: new Date().toISOString() });
  }
  // 1) this page's own session data / header  2) the identity another page stored (recent, same region only)
  function readIdentity(hintRegion) {
    const o = ownIdentity();
    if (o.account) { const out = { account: o.account, region: o.region || hintRegion || '', source: 'page' }; if (!inFrame) rememberIdentity({ account: o.account, region: o.region || '' }); return out; }
    const s = store.load(CFG.keyIdent), age = s.at ? Date.now() - Date.parse(s.at) : Infinity;
    if (s.account && age < IDENT_MAX_AGE_MS && (!hintRegion || !s.region || s.region === hintRegion)) return { account: s.account, region: s.region || hintRegion || '', source: 'shared', at: s.at };
    return { account: '', region: hintRegion || '', source: '' };
  }
  const ago = at => { const m = Math.max(0, Math.round((Date.now() - Date.parse(at)) / 60000)); return m < 1 ? 'just now' : m < 60 ? `${m} min ago` : `${Math.round(m / 60)} h ago`; };

  // ---------- accounts: region + online ID. The list is built ONLY from data that exists (nothing is pre-listed) ----------
  const acctId = (region, account) => `${(region || '').toUpperCase()}|${account || ''}`;
  const acctLabel = id => { const [r, a] = id.split('|'); return a ? `${r} · ${a}` : r; };
  function logScan(kind, region, account, tab, info) {
    region = (region || '').toUpperCase(); if (!region) return;
    const s = store.load(CFG.keyScans); s[`${kind}|${region}|${account || ''}|${tab || ''}`] = Object.assign({ at: new Date().toISOString() }, info); store.save(CFG.keyScans, s);
  }
  const getScan = (kind, region, account, tab) => store.load(CFG.keyScans)[`${kind}|${(region || '').toUpperCase()}|${account || ''}|${tab || ''}`];
  function accountUniverse(kind) {
    const set = new Set();
    Object.keys(store.load(CFG.keyScans)).forEach(k => { const p = k.split('|'); if (p[0] === kind && p[1]) set.add(acctId(p[1], p[2])); });
    Object.values(store.load(kind === 'lib' ? CFG.keyLib : CFG.keyTx)).forEach(r => r.region && set.add(acctId(r.region, r.account)));
    return [...set].sort();
  }
  function summary(kind) {
    const ids = accountUniverse(kind), lines = [];
    ids.forEach(id => {
      const [r, a] = id.split('|'); const e = getScan(kind, r, a, kind === 'lib' ? TAB : '');
      const state = !e ? 'data stored' : e.complete ? '✔ complete' : '⚠ partial';
      const short = !e ? `${esc(acctLabel(id))}: ${state}`
        : kind === 'lib' ? `${esc(acctLabel(id))}: ${state} · ${e.rows}${e.total ? '/' + e.total : ''} games`
        : `${esc(acctLabel(id))}: ${state} · ${e.rows} transactions`;
      const tip = kind === 'tx' && e ? `${acctLabel(id)}: ${state} · ${e.rows} transactions · back to ${e.oldest || '?'}${e.details ? ' · ' + e.details + ' paid orders read' : ''}` : short.replace(/<[^>]*>/g, '');
      lines.push({ html: short, tip, id, hasData: true });
    });
    return { have: ids.length, total: ids.length, lines };
  }
  function controller() {
    const c = { running: false, paused: false, stop: false,
      async tick() { while (c.paused && !c.stop) await sleep(300); },
      begin() { c.running = true; c.stop = false; c.paused = false; }, end() { c.running = false; c.paused = false; } };
    return c;
  }

  // The checkout site ships global button/link styles; these !important rules keep our panels identical (and compact) on both sites.
  function injectStyle() {
    if (document.getElementById('psnlx-style')) return;
    const s = document.createElement('style'); s.id = 'psnlx-style';
    s.textContent = '[data-psnlx]{font:12px/1.3 system-ui,sans-serif!important;color:#fff!important}'
      + '[data-psnlx] *{box-sizing:border-box!important;font-family:inherit!important;letter-spacing:normal!important;text-transform:none!important}'
      + '[data-psnlx] button{display:inline-block!important;flex:0 0 auto!important;width:auto!important;height:auto!important;min-width:0!important;min-height:0!important;margin:0!important;padding:2px 8px!important;border:0!important;border-radius:5px!important;box-shadow:none!important;color:#fff!important;font-size:12px!important;font-weight:400!important;line-height:1.5!important;cursor:pointer!important}'
      + '[data-psnlx] button.psx-s{padding:1px 7px!important;margin:2px 4px 2px 0!important}'
      + '[data-psnlx] a{color:#7fc1ff!important;font-size:11px!important;font-weight:400!important;text-decoration:none!important;cursor:pointer!important}'
      + '[data-psnlx] a.psx-on{text-decoration:underline!important}'
      + '[data-psnlx] .psx-row{display:flex!important;align-items:center!important;gap:6px!important}'
      + '[data-psnlx] .psx-txt{flex:1 1 auto!important;min-width:0!important;overflow:hidden!important;text-overflow:ellipsis!important;white-space:nowrap!important}'
      + '[data-psnlx] .psx-dl{flex:0 0 auto!important}';
    (document.head || document.documentElement).appendChild(s);
  }
  // ---------- compact panel ----------
  function makePanel(id, title, width, defaultPos, shield) {
    injectStyle();
    const saved = store.load(CFG.keyUi)[id] || {};
    const box = document.createElement('div');
    box.style.cssText = `position:fixed;z-index:2147483647;background:#0b1e3a;color:#fff;font:12px/1.3 system-ui,sans-serif;border-radius:9px;box-shadow:0 4px 16px rgba(0,0,0,.45);width:${width}px;user-select:none;max-height:90vh;overflow:auto`;
    const pos = saved.pos || defaultPos;
    box.style.left = Math.max(0, Math.min(pos.x, window.innerWidth - 80)) + 'px'; box.style.top = Math.max(0, Math.min(pos.y, window.innerHeight - 40)) + 'px';
    box.innerHTML = `<div class="hd" style="cursor:move;padding:3px 8px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #27456f;position:sticky;top:0;background:#0b1e3a"><b>${title}</b><span class="mn" style="cursor:pointer;padding:0 6px;font-size:15px" title="minimise / restore">&#8211;</span></div><div class="bd" style="padding:4px 8px 5px"></div>`;
    box.setAttribute('data-psnlx', '1'); document.body.appendChild(box);
    if (shield) ['mousedown', 'mouseup', 'click', 'pointerdown', 'pointerup', 'touchstart', 'focusin'].forEach(ev => box.addEventListener(ev, e => e.stopPropagation()));
    const bd = box.querySelector('.bd');
    const persist = () => { const u = store.load(CFG.keyUi); u[id] = Object.assign(u[id] || {}, { pos: { x: box.offsetLeft, y: box.offsetTop }, min: bd.style.display === 'none' }); store.save(CFG.keyUi, u); };
    const setMin = (m, keep = true) => { bd.style.display = m ? 'none' : 'block'; box.querySelector('.mn').innerHTML = m ? '+' : '&#8211;'; if (keep) persist(); };
    box.querySelector('.mn').onclick = () => setMin(bd.style.display !== 'none');
    if (saved.min) setMin(true);
    let drag = null;
    box.querySelector('.hd').onmousedown = e => { if (e.target.className === 'mn') return; drag = { dx: e.clientX - box.offsetLeft, dy: e.clientY - box.offsetTop }; e.preventDefault(); };
    window.addEventListener('mousemove', e => { if (drag) { box.style.left = Math.max(0, e.clientX - drag.dx) + 'px'; box.style.top = Math.max(0, e.clientY - drag.dy) + 'px'; } }, true);
    window.addEventListener('mouseup', () => { if (drag) { drag = null; persist(); } }, true);
    return { box, bd, setMin };
  }
  const BTN = '';
  // One compact block: identity line, status, optional sub-line, two button rows, and a footer of toggle links (Accounts / Delete / Help).
  function panelUI(parent) {
    const el = document.createElement('div'); let setAt = 0;
    el.innerHTML = '<div class="id" style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis"></div>'
      + '<div class="bar" style="display:none;height:3px;background:#27456f;border-radius:2px;margin:3px 0 1px"><div style="height:100%;width:0;background:#2185f4;border-radius:2px;transition:width .3s"></div></div>'
      + '<div class="st" style="margin:2px 0;min-height:15px;word-break:break-word"></div>'
      + '<div class="sub" style="font-size:11px;opacity:.8;margin:0 0 2px;word-break:break-word"></div>'
      + '<div class="r1" style="display:flex;flex-wrap:wrap;gap:3px;margin:3px 0 0"></div><div class="r2" style="display:flex;flex-wrap:wrap;gap:3px;margin:3px 0 0"></div>'
      + '<div class="ft" style="display:flex;gap:12px;border-top:1px solid #27456f;padding-top:3px;margin-top:5px"></div>'
      + '<div class="pn" style="display:none;font-size:11px;margin-top:4px;max-height:230px;overflow:auto"></div>';
    parent.appendChild(el);
    const q = s => el.querySelector(s), panes = [];
    const showPane = which => { panes.forEach(p => { const on = p === which && p.body.style.display === 'none'; p.body.style.display = on ? 'block' : 'none'; p.link.classList.toggle('psx-on', on); }); q('.pn').style.display = panes.some(p => p.body.style.display === 'block') ? 'block' : 'none'; };
    const api = {
      el,
      set(t) { const s = q('.st'); s.textContent = t; setAt = Date.now(); s.title = 'updated ' + hhmmss(); },
      age() { return Date.now() - setAt; },
      text() { return q('.st').textContent; },
      sub(t) { q('.sub').textContent = t || ''; q('.sub').style.display = t ? 'block' : 'none'; },
      identity(html) { q('.id').innerHTML = html; },
      progress(f) { const b = q('.bar'); if (f === null || f === undefined) { b.style.display = 'none'; return; } b.style.display = 'block'; b.firstChild.style.width = Math.max(0, Math.min(100, Math.round(f * 100))) + '%'; },
      button(label, fn, color, title, row) { const b = document.createElement('button'); b.textContent = label; if (title) b.title = title; b.style.setProperty('background', color || '#2185f4', 'important'); b.onclick = fn; q(row === 2 ? '.r2' : '.r1').appendChild(b); return b; },
      pane(label) {
        const link = document.createElement('a'); link.href = '#'; link.textContent = label;
        const body = document.createElement('div'); body.style.display = 'none'; q('.pn').appendChild(body); q('.ft').appendChild(link);
        const p = { link, body, setLabel(t) { link.textContent = t; } }; panes.push(p);
        link.onclick = e => { e.preventDefault(); showPane(p); }; return p;
      },
      help(steps) { const p = api.pane('Help'); p.body.innerHTML = `<ol style="margin:0 0 0 16px;padding:0;opacity:.9">${steps.map(s => `<li style="margin:1px 0">${s}</li>`).join('')}</ol>`; return p; }
    };
    api.sub(''); return api;
  }
  const smallBtn = (parent, label, fn, title) => { const b = document.createElement('button'); b.className = 'psx-s'; b.textContent = label; if (title) b.title = title; b.style.setProperty('background', '#7f8c8d', 'important'); b.onclick = fn; parent.appendChild(b); return b; };
  const pauseToggle = (ui, ctl) => { const b = ui.button('⏸', () => { ctl.paused = !ctl.paused; b.textContent = ctl.paused ? '▶' : '⏸'; b.title = ctl.paused ? 'Resume' : 'Pause'; b.style.setProperty('background', ctl.paused ? '#d68910' : '#2185f4', 'important'); if (ctl.paused) ui.set('⏸ Paused - press ▶ to continue.'); }, '#2185f4', 'Pause'); return b; };
  // Accounts pane: one block per kind; lists exist only for data that exists (nothing pre-listed). Each row ends with its own ⬇ download.
  function accountsView(pane, kinds) {
    pane.body.addEventListener('click', e => { const a = e.target.closest('[data-k]'); if (!a) return; e.preventDefault(); const k = kinds.find(x => x.kind === a.getAttribute('data-k')); if (k) k.dl(a.getAttribute('data-dl') || ''); });
    return { refresh() {
      const ids = new Set(); let html = '';
      kinds.forEach(k => {
        const s = summary(k.kind); s.lines.forEach(l => ids.add(l.id));
        html += `<div style="margin:1px 0 4px"><div class="psx-row"><b>${k.title} (${s.have})</b>${s.have ? `<a href="#" data-k="${k.kind}" data-dl="" title="download every ${k.title.toLowerCase()} account in one CSV">⬇ all</a>` : ''}</div>`;
        html += s.lines.length ? s.lines.map(l => `<div class="psx-row"><span class="psx-txt" title="${esc(l.tip)}">${l.html}</span><a class="psx-dl" href="#" data-k="${k.kind}" data-dl="${esc(l.id)}" title="download this account">⬇</a></div>`).join('') : '<div style="opacity:.7">Nothing scanned yet.</div>';
        html += '</div>';
      });
      pane.body.innerHTML = html; pane.setLabel(`Accounts (${ids.size})`); return ids.size;
    } };
  }

  // =====================================================================================
  //  LIBRARY ENGINE
  // =====================================================================================
  const LIB_COLS = [['Name', 'name'], ['Platform', 'platform'], ['TitleID', 'titleId'], ['ProductID', 'productId'], ['ConceptID', 'conceptId'], ['Subscription', 'subscription'],
    ['StoreURL', 'storeUrl'], ['Image', 'image'], ['Region', 'region'], ['Account', 'account'], ['Locale', 'locale'], ['Tab', 'tab'], ['ReportedTotal', 'reportedTotal'], ['LibraryPage', 'page'], ['ScannedAt', 'scannedAt']];
  function downloadLib(id, say) {
    const [r, a] = id ? id.split('|') : ['', '']; const items = store.load(CFG.keyLib);
    const rows = Object.values(items).filter(x => !id || (x.region === r && (x.account || '') === a)).sort((x, y) => x.region.localeCompare(y.region) || (x.account || '').localeCompare(y.account || '') || x.name.localeCompare(y.name));
    if (!rows.length) { say(id ? `No stored library data for ${acctLabel(id)}.` : 'No stored library data.'); return; }
    const name = id ? `${r}${a ? '-' + slug(a) : ''}` : `ALL-${new Set(rows.map(x => acctId(x.region, x.account))).size}-accounts`;
    download(`psn-library_${name}_${today()}.csv`, toCSV(LIB_COLS, rows)); say(`Downloaded ${rows.length} library rows (${id ? acctLabel(id) : 'all accounts'}).`);
  }
  function libraryEngine(ui, ctl, onChange) {
    const SEL = {
      tile: '[data-qa="collection-game-list-product"]', link: 'a[data-qa="collection-game-list-product#store-link"]',
      title: '[data-qa="collection-game-list-product#title"]', platform: '[data-qa="collection-game-list-product#platform-tags"] span',
      service: '[data-qa="collection-game-list-product#service-upsell"]', image: '[data-qa="collection-game-list-product#game-art#image#image"]',
      next: 'button[data-qa="pagination#next"]', prev: 'button[data-qa="pagination#previous"]', page1: 'button[data-qa="pagination#page1"]',
      pageBtns: 'nav[data-qa="pagination"] ol button[value]', current: 'nav[data-qa="pagination"] button[aria-current="page"]'
    };
    let items = store.load(CFG.keyLib);
    const tabName = () => (location.pathname.split('/')[1] || 'library');
    const tiles = () => document.querySelectorAll(SEL.tile);
    const reportedTotal = () => { for (const h of document.querySelectorAll('h1')) { const m = h.textContent.match(/of\s+([\d,]+)\)/i); if (m) return parseInt(m[1].replace(/,/g, ''), 10); } return null; };
    const totalPages = () => { const n = [...document.querySelectorAll(SEL.pageBtns)].map(b => parseInt(b.value, 10)).filter(x => !isNaN(x)); return n.length ? Math.max(...n) : 1; };
    const currentPageNo = () => { const c = document.querySelector(SEL.current); return c ? (parseInt(c.textContent, 10) || 1) : 1; };
    const signature = () => [...tiles()].map(t => { const a = t.querySelector(SEL.link); return a ? a.getAttribute('data-telemetry-meta') : ''; }).join('|');
    function parseTile(tile, pageNo, account) {
      const a = tile.querySelector(SEL.link); if (!a) return null;
      let meta = {}; try { meta = JSON.parse(a.getAttribute('data-telemetry-meta') || '{}'); } catch (e) {}
      const name = (meta.titleName || (tile.querySelector(SEL.title) || {}).textContent || '').trim();
      const href = (a.getAttribute('href') || '').split('?')[0];
      const loc = (/\/([a-z]{2}-[a-z]{2})\//i.exec(href) || [])[1] || '';
      const img = tile.querySelector(SEL.image);
      let image = img ? (img.currentSrc || img.src || '') : ''; image = /^https?:/i.test(image) ? image.split('?')[0] : '';
      return { name, platform: [...tile.querySelectorAll(SEL.platform)].map(s => s.textContent.trim()).filter(Boolean).join('/'),
        titleId: meta.titleId || '', productId: meta.productId || '', conceptId: meta.conceptId || '',
        subscription: /ps\s*plus/i.test((tile.querySelector(SEL.service) || {}).textContent || '') ? 'PS_PLUS' : 'NONE',
        storeUrl: href, image, locale: loc.toLowerCase(), region: regionFromLocale(loc), account: account || '', tab: TAB, page: pageNo, scannedAt: new Date().toISOString() };
    }
    function scrapeVisible(reported, account, region) {
      const pageNo = currentPageNo(); let added = 0;
      tiles().forEach(t => {
        const rec = parseTile(t, pageNo, account); if (!rec || !rec.name) return;
        rec.region = region; rec.reportedTotal = reported || '';
        const key = [region, account || '', rec.titleId || rec.productId, rec.productId].join('|');
        if (!items[key]) added++;
        items[key] = Object.assign(items[key] || {}, rec);
      });
      store.save(CFG.keyLib, items); return added;
    }
    const countFor = (region, account) => Object.values(items).filter(r => r.region === region && (r.account || '') === (account || '')).length;
    const pageRegion = () => { const t = tiles()[0]; return (t && parseTile(t, 1, '').region) || readIdentity().region || ''; };
    const needsScroll = () => [...tiles()].some(t => { const i = t.querySelector(SEL.image); return i && !/^https?:/i.test(i.currentSrc || i.src || ''); });   // artwork still a lazy placeholder
    async function gentleScroll() { if (!needsScroll()) return; const h = document.documentElement.scrollHeight; for (let y = 0; y < h; y += 900) { window.scrollTo(0, y); await ctl.tick(); await sleep(70); } await sleep(350); window.scrollTo(0, 0); }
    // Require a stable, nonempty tile set; a changed first tile alone is not a finished render.
    async function settle() {
      let last = '', stable = 0; const t0 = Date.now();
      while (Date.now() - t0 < CFG.pageTimeoutMs && !ctl.stop) {
        await ctl.tick(); await sleep(200); const sig = signature();
        if (sig && sig === last) { if (++stable >= 4) return true; }
        else { last = sig; stable = 0; }
      }
      return false;
    }
    const usable = b => !!b && !b.disabled && b.getAttribute('aria-disabled') !== 'true';
    let scanRegion = '', scanAccount = '';
    const CP_MAX_AGE_MS = 24 * 60 * 60 * 1000;
    const saveCheckpoint = (region, account, page, rows, total) => { const c = store.load(CFG.keyCheckpoint); c[acctId(region, account)] = { page, rows, total, at: new Date().toISOString() }; store.save(CFG.keyCheckpoint, c); };
    const loadCheckpoint = (region, account) => { const c = store.load(CFG.keyCheckpoint)[acctId(region, account)]; const age = c && c.at ? Date.now() - Date.parse(c.at) : Infinity; return c && age < CP_MAX_AGE_MS ? c : null; };
    const clearCheckpoint = (region, account) => { const c = store.load(CFG.keyCheckpoint); delete c[acctId(region, account)]; store.save(CFG.keyCheckpoint, c); };
    // End of the list: only when the pager is present, we are on its highest page number AND Next is unusable. A missing pager is never "the end" unless the whole list fits one page.
    function atLastPage() {
      const nav = document.querySelector('nav[data-qa="pagination"]');
      const nums = [...document.querySelectorAll(SEL.pageBtns)].map(b => Number(b.value)).filter(Number.isFinite);
      if (!nav || !nums.length) { const total = reportedTotal(); return !!(total && countFor(scanRegion, scanAccount) >= total); }
      return currentPageNo() >= Math.max(...nums) && !usable(document.querySelector(SEL.next));
    }
    // One page forward (dir 1) or back (dir -1). Controls are re-queried on every poll (the site re-renders them); a delayed first click is never followed by a second one.
    async function stepPage(dir) {
      const from = currentPageNo(), want = from + dir, sigBefore = signature();
      const moved = () => currentPageNo() === want && !!signature() && signature() !== sigBefore;
      const edge = () => document.querySelector(dir > 0 ? SEL.next : SEL.prev);
      const button = () => { const num = document.querySelector(`nav[data-qa="pagination"] ol button[value="${want}"]`); return usable(num) ? num : usable(edge()) ? edge() : null; };
      for (let attempt = 1; attempt <= 4 && !ctl.stop; attempt++) {
        await ctl.tick();
        if (moved()) return (await settle()) ? 'moved' : 'stuck';
        if (dir > 0 && atLastPage()) return 'end';
        if (dir < 0 && from <= 1) return 'end';
        ui.set(`Waiting for page ${want} controls (try ${attempt}/4)...`);
        await waitFor(() => ctl.stop || moved() || button(), CFG.controlWaitMs, 150);
        if (ctl.stop) return 'stopped';
        if (moved()) return (await settle()) ? 'moved' : 'stuck';
        const btn = button(); if (!btn) continue;
        btn.scrollIntoView({ block: 'center' }); await sleep(200); await ctl.tick(); if (ctl.stop) return 'stopped';
        if (moved()) return (await settle()) ? 'moved' : 'stuck';
        const again = button(); if (again) again.click();
        if (await waitFor(() => ctl.stop || moved(), CFG.pageTimeoutMs, 150)) return ctl.stop ? 'stopped' : ((await settle()) ? 'moved' : 'stuck');
      }
      return ctl.stop ? 'stopped' : 'stuck';
    }
    // Go to any page from wherever the user left the pager (the pager shows only a few numbers, and page 1 is not always among them): jump to 1 when possible, otherwise step.
    async function gotoPage(target) {
      let guard = 0;
      while (!ctl.stop && currentPageNo() !== target && guard++ < 400) {
        const cur = currentPageNo();
        if (target === 1 && cur > 2) { const f = document.querySelector(SEL.page1); if (usable(f)) { const sig = signature(); f.click(); if (await waitFor(() => currentPageNo() === 1 && signature() && signature() !== sig, CFG.pageTimeoutMs)) { await settle(); continue; } } }
        ui.set(`Going to page ${target} (now on ${cur})...`); ui.progress(Math.min(cur, target) / Math.max(cur, target, 1));
        const r = await stepPage(target > cur ? 1 : -1); if (r !== 'moved') return false; await sleep(rnd(350, 800));
      }
      return currentPageNo() === target;
    }
    async function run() {
      if (ctl.running) return; ctl.begin();
      try {
        if (tabName() !== TAB) { ui.set('Open the Purchased tab first (library.playstation.com/recently-purchased).'); return; }
        ui.set('Starting - looking for game tiles...');
        if (!(await waitFor(() => tiles().length, CFG.pageTimeoutMs))) { ui.set('No game tiles found on this tab.'); return; }
        const region = pageRegion(), account = readIdentity(region).account;
        if (!region) { ui.set('Region not detected yet - wait for the tiles to load and try again.'); return; }
        if (!account) { ui.set('Account not detected. Make sure you are signed in, reload this page, and try again.'); return; }
        scanRegion = region; scanAccount = account;
        const total = reportedTotal();
        // Resume only when a recent checkpoint AND its stored games exist; otherwise start fresh (old rows for this account are replaced).
        let cp = loadCheckpoint(region, account);
        if (cp && (cp.page <= 1 || countFor(region, account) < 1 || (cp.total && total && cp.total !== total))) cp = null;
        if (!cp) { for (const k of Object.keys(items)) if (items[k].region === region && (items[k].account || '') === account) delete items[k]; store.save(CFG.keyLib, items); clearCheckpoint(region, account); }
        const target = cp ? cp.page : 1;
        if (cp) ui.set(`Resuming from page ${cp.page} (${countFor(region, account)} games already stored)...`);
        if (currentPageNo() !== target && !(await gotoPage(target))) {
          ui.set(ctl.stop ? 'Stopped while moving to the start page.' : `Could not reach page ${target}. ${cp ? 'The checkpoint is kept - ' : ''}press Start to retry.`); return;
        }
        let pages = 0;
        while (!ctl.stop && pages < CFG.maxPages) {
          await ctl.tick();
          if (document.hidden) { ui.set('Paused: this browser tab is in the background. Bring it to the front to continue.'); await waitFor(() => !document.hidden || ctl.stop, 3600000, 500); if (ctl.stop) break; }
          const who = readIdentity(region).account; if (who && who !== account) { ui.set('The signed-in account changed. Stopped - reload this page and press Start.'); logScan('lib', region, account, TAB, { rows: countFor(region, account), total, complete: false }); onChange(); return; }
          const n = currentPageNo(), pagesTotal = totalPages();
          ui.set(`Scanning page ${n} of ${pagesTotal}... ${countFor(region, account)}${total ? ' / ' + total : ''} games.`); ui.progress(n / pagesTotal);
          const stable = await settle(); if (ctl.stop) break;
          if (!stable) { ui.set(`Page ${n} did not finish loading. Saved as PARTIAL - press Start to resume from page ${n}.`); logScan('lib', region, account, TAB, { rows: countFor(region, account), total, complete: false }); onChange(); return; }
          await gentleScroll(); const added = scrapeVisible(total, account, region); pages++;
          saveCheckpoint(region, account, n, countFor(region, account), total);
          logScan('lib', region, account, TAB, { rows: countFor(region, account), total, complete: false });
          ui.set(`Page ${n} of ${pagesTotal} done: +${added}. ${countFor(region, account)}${total ? ' of ' + total : ''} games.`);
          if (atLastPage()) break;
          await sleep(rnd(CFG.minDelayMs, CFG.maxDelayMs)); await ctl.tick(); if (ctl.stop) break;
          const r = await stepPage(1);
          if (r === 'end' || r === 'stopped') break;
          if (r === 'stuck') { ui.set(`Could not move past page ${n}. Saved as PARTIAL (${countFor(region, account)} games) - press Start to resume from page ${n}.`); logScan('lib', region, account, TAB, { rows: countFor(region, account), total, complete: false }); onChange(); return; }
        }
        const got = countFor(region, account), complete = !ctl.stop && !!total && got >= total;
        if (complete) clearCheckpoint(region, account);
        ui.set(`${ctl.stop ? 'Stopped' : 'Done'} - ${acctLabel(acctId(region, account))}: ${got} games${total ? ' of ' + total : ''}. ${complete ? 'COMPLETE - you can download.' : ctl.stop ? 'PARTIAL - press Start to resume.' : 'PARTIAL - the list ended below the reported total; press Start to redo it.'}`);
        logScan('lib', region, account, TAB, { rows: got, total, complete }); onChange();
      } catch (e) { ui.set('Error: ' + (e && e.message || e)); } finally { ui.progress(null); ctl.end(); }
    }
    return {
      run, pageRegion, reload() { items = store.load(CFG.keyLib); }, resetCheckpoint(region, account) { clearCheckpoint(region, account); ui.set('Next Start rescans from page 1.'); },
      clearAccount(region, account) { if (!region) { ui.set('Open the Purchased tab so the region can be read.'); return; }
        if (!confirm(`Delete stored GAME LIBRARY data for ${acctLabel(acctId(region, account))}?`)) return;
        let n = 0; for (const k of Object.keys(items)) if (items[k].region === region && (items[k].account || '') === (account || '')) { delete items[k]; n++; }
        store.save(CFG.keyLib, items); const sc = store.load(CFG.keyScans); delete sc[`lib|${region}|${account || ''}|${TAB}`]; store.save(CFG.keyScans, sc); clearCheckpoint(region, account);
        ui.set(`Cleared ${n} stored games for ${acctLabel(acctId(region, account))}.`); onChange(); },
      clearAll() { if (confirm('Delete stored GAME LIBRARY data for ALL accounts? (Transactions are not touched.)')) { items = {}; store.save(CFG.keyLib, items);
        const sc = store.load(CFG.keyScans); Object.keys(sc).forEach(k => { if (k.startsWith('lib|')) delete sc[k]; }); store.save(CFG.keyScans, sc); store.save(CFG.keyCheckpoint, {}); ui.set('All stored game library data cleared.'); onChange(); } }
    };
  }

  // =====================================================================================
  //  TRANSACTION ENGINE (inside the Transaction History drawer) - loading loop is the proven v4.0.0 one, only faster
  // =====================================================================================
  const TX_COLS = [['Date', 'date'], ['Name', 'name'], ['Amount', 'amount'], ['OrderID', 'orderId'], ['Locale', 'locale'], ['Region', 'region'], ['Account', 'account'], ['Details', 'details'], ['ScannedAt', 'scannedAt']];
  function downloadTx(id, say) {
    const [r, a] = id ? id.split('|') : ['', '']; const tx = store.load(CFG.keyTx);
    const rows = Object.values(tx).filter(t => !id || (t.region === r && (t.account || '') === a));
    if (!rows.length) { say(id ? `No stored transactions for ${acctLabel(id)}.` : 'No stored transactions.'); return; }
    const name = id ? `${r}${a ? '-' + slug(a) : ''}` : `ALL-${new Set(rows.map(x => acctId(x.region, x.account))).size}-accounts`;
    download(`psn-transactions_${name}_${today()}.csv`, toCSV(TX_COLS, rows)); say(`Downloaded ${rows.length} transactions (${id ? acctLabel(id) : 'all accounts'}).`);
  }
  function clearTxAccount(region, account, say, onChange) {
    if (!region) { say('Region not known yet.'); return; }
    if (!confirm(`Delete stored TRANSACTION data for ${acctLabel(acctId(region, account))}?`)) return;
    const tx = store.load(CFG.keyTx); let n = 0; for (const k of Object.keys(tx)) if (tx[k].region === region && (tx[k].account || '') === (account || '')) { delete tx[k]; n++; } store.save(CFG.keyTx, tx);
    const sc = store.load(CFG.keyScans); delete sc[`tx|${region}|${account || ''}|`]; store.save(CFG.keyScans, sc); say(`Cleared ${n} stored transactions for ${acctLabel(acctId(region, account))}.`); onChange();
  }
  function clearTxAll(say, onChange) {
    if (!confirm('Delete stored TRANSACTION data for ALL accounts? (Game library data is not touched.)')) return; store.save(CFG.keyTx, {});
    const sc = store.load(CFG.keyScans); Object.keys(sc).forEach(k => { if (k.startsWith('tx|')) delete sc[k]; }); store.save(CFG.keyScans, sc); say('All stored transaction data cleared.'); onChange();
  }
  function parseAmt(txt) {                                              // '£0.79' -> 0.79, '1.099,00 TL' -> 1099, '+£50.00' / '(£3.00)' -> null (credit / refund)
    const t = String(txt || '').replace(/\u00a0/g, ' ').trim(); if (/^[+(\-\u2212]/.test(t)) return null;
    let n = t.replace(/[^\d.,]/g, ''); if (!n) return null;
    if (n.includes(',') && n.includes('.')) { const dec = n.lastIndexOf(',') > n.lastIndexOf('.') ? ',' : '.'; n = n.replace(dec === ',' ? /\./g : /,/g, '').replace(dec, '.'); }
    else if (n.includes(',')) { const tail = n.split(',').pop(); n = tail.length <= 2 ? n.replace(',', '.') : n.replace(/,/g, ''); }
    const v = parseFloat(n); return isNaN(v) ? null : v;
  }

  function txEngine(report, ctl, getCtx, onChange) {
    const SEL = { card: '.transaction-history-card', fields: '.transaction-history-card-details-field', idEl: '.transaction-history-card-content-details',
      wrap: '.transaction-history-workstream-wrapper', btn: 'button.transaction-history-card-content-button' };
    let reachedEnd = false;
    const cards = () => document.querySelectorAll(SEL.card), cardCount = () => cards().length;
    const fieldsOf = card => [...card.querySelectorAll(SEL.fields)].map(x => x.textContent.trim());
    const hasError = () => /something went wrong/i.test(pageText());
    const atEnd = () => /no more transactions present/i.test(pageText());
    const oldestDate = () => { const c = cards(); return c.length ? (fieldsOf(c[c.length - 1])[0] || '?') : '-'; };
    const wrapEl = () => document.querySelector(SEL.wrap) || document.scrollingElement;       // exactly as in v4.0.0
    const guessRegion = () => { const c = cards()[0]; const amt = c ? (fieldsOf(c)[1] || '') : '';
      if (/£/.test(amt)) return 'UK'; if (/TL|₺/.test(amt)) return 'TR'; if (/€/.test(amt)) return 'EU'; if (/\$/.test(amt)) return 'US'; return regionFromLocale(document.documentElement.lang) || ''; };
    const ctx = () => { const c = getCtx(); return { region: (c.region || guessRegion() || '').toUpperCase(), account: (c.account || '').trim() }; };
    const storedFor = (r, a) => Object.values(store.load(CFG.keyTx)).filter(t => t.region === r && (t.account || '') === a).length;
    const keyOf = (c, card) => { const f = fieldsOf(card), idEl = card.querySelector(SEL.idEl), oid = idEl && idEl.id ? idEl.id.split('-')[0] : ''; return { key: [c.region, c.account, oid, f[2], f[1], f[0]].join('|'), f, oid }; };

    async function loadAll() {
      reachedEnd = false;
      let last = cardCount(), stagnant = 0, round = 0, errors = 0; const t0 = Date.now();
      if (last === 0) { report('No transactions are visible yet - wait for the list to appear, then press Start again.'); return false; }
      report(`Loading: ${last} transactions visible. Fast-scrolling to load the rest...`);
      while (!ctl.stop && round < CFG.txMaxRounds) {
        round++; await ctl.tick();
        if (atEnd()) { reachedEnd = true; break; }
        if (stagnant >= CFG.txStagnantLimit) break;
        if (document.hidden) { report('Paused: this browser tab is in the background. Bring it to the front to continue.'); await waitFor(() => !document.hidden || ctl.stop, 3600000, 500); }
        if (hasError()) {
          errors++; report(`The site showed an error. Waiting 30 s (attempt ${errors}/3)...`);
          if (errors > 3) { report(`Stopped after repeated site errors at ${cardCount()} transactions. Reopen Transaction History and press Start again.`); return false; }
          await sleep(30000); const w2 = wrapEl(); w2.scrollTo({ top: Math.max(0, w2.scrollHeight - 1500) }); await sleep(2000); continue;
        }
        const w = wrapEl(); if (stagnant === 0) w.scrollTop = w.scrollHeight; else w.scrollTo({ top: w.scrollHeight, behavior: 'smooth' });   // instant jump first; smooth only as a nudge
        const more = [...document.querySelectorAll('button')].find(b => /^(show|load|view)\s+more$/i.test(b.textContent.trim())); if (more) more.click();
        const patience = CFG.txPatienceMs + stagnant * CFG.txPatienceStepMs; const secs = Math.max(1, Math.round((Date.now() - t0) / 1000));
        report(`Loading history... ${cardCount()} loaded, back to ${oldestDate()} (${secs}s).${stagnant ? ` Waiting for the site (${stagnant}/${CFG.txStagnantLimit})...` : ''}`);
        const grew = await waitFor(() => cardCount() > last || atEnd() || ctl.stop, patience, 50);
        if (atEnd()) { reachedEnd = true; break; }
        if (grew && cardCount() > last) { last = cardCount(); stagnant = 0; await sleep(rnd(CFG.txFastMinMs, CFG.txFastMaxMs)); }
        else { stagnant++; w.scrollTo({ top: Math.max(0, w.scrollHeight - 700) }); if (stagnant === 2) { const c = cards(); if (c.length) c[c.length - 1].scrollIntoView({ block: 'end' }); } await sleep(800); }
      }
      return reachedEnd;
    }
    function scanRows(c) {
      const locale = (document.documentElement.lang || '').trim(); const tx = store.load(CFG.keyTx); let n = 0;
      cards().forEach(card => { const k = keyOf(c, card); if (k.f.length < 3) return;
        tx[k.key] = Object.assign(tx[k.key] || {}, { date: k.f[0], amount: k.f[1], name: k.f[2], orderId: k.oid, locale, region: c.region, account: c.account, scannedAt: new Date().toISOString() }); n++; });
      store.save(CFG.keyTx, tx);
      const prev = getScan('tx', c.region, c.account, '') || {};
      logScan('tx', c.region, c.account, '', { rows: storedFor(c.region, c.account), oldest: oldestDate(), complete: reachedEnd || atEnd(), details: prev.details || 0 });
      return n;
    }
    async function openPaid(c, all) {
      const stored = store.load(CFG.keyTx);
      const todo = [...cards()].filter(card => {
        const k = keyOf(c, card); if (k.f.length < 3) return false;
        const amt = parseAmt(k.f[1]); if (amt === null || amt <= 0) return false;                      // free, credit or refund
        const single = !k.f[2].includes(', ');
        if (single && !all) return false;
        if (single && /^wallet\s*funding$/i.test(k.f[2])) return false;
        return !(stored[k.key] || {}).details;
      });
      let i = 0, dirty = 0;
      const flush = () => { if (dirty) { const cur = store.load(CFG.keyTx); Object.keys(stored).forEach(k => { if (stored[k].details && cur[k]) cur[k].details = stored[k].details; }); store.save(CFG.keyTx, cur); dirty = 0; } };
      try {
        for (const card of todo) {
          await ctl.tick(); if (ctl.stop) break; i++;
          report(`Reading paid order ${i} of ${todo.length}...`);
          const btn = card.querySelector(SEL.btn) || card.querySelector('button'); if (!btn) continue;
          card.scrollIntoView({ block: 'center' }); await sleep(150);
          if (btn.getAttribute('aria-expanded') !== 'true') btn.click();
          if (!(await waitFor(() => btn.getAttribute('aria-expanded') === 'true', 4000))) continue;
          await sleep(rnd(250, 500));
          const k = keyOf(c, card); if (stored[k.key]) { stored[k.key].details = cleanDetails(card.innerText); dirty++; }
          if (dirty >= 10) flush();
          if (btn.getAttribute('aria-expanded') === 'true') btn.click();
          await sleep(rnd(CFG.txDetailMinMs, CFG.txDetailMaxMs));
        }
      } finally { flush(); }
      const withDetails = Object.values(store.load(CFG.keyTx)).filter(t => t.region === c.region && (t.account || '') === c.account && t.details).length;
      const cur = getScan('tx', c.region, c.account, '') || {}; logScan('tx', c.region, c.account, '', Object.assign(cur, { details: withDetails }));
      return i;
    }
    // One button: load everything -> save the transactions -> (optional) open paid orders
    async function start(opts) {
      if (ctl.running) return; ctl.begin();
      try {
        const c = ctx(); if (!c.region) { report('Region not known yet - open the library page once, or wait for the list to show prices.'); return; }
        if (!c.account) { report('Account not detected. Open library.playstation.com once while signed in with this account (or reload this page), then press Start.'); return; }
        const sawEnd = await loadAll();
        const n = scanRows(c); onChange();
        if (!sawEnd) { report(ctl.stop ? `Stopped at ${cardCount()} (back to ${oldestDate()}). Saved ${n} as PARTIAL.` : `The list stopped growing at ${cardCount()} (back to ${oldestDate()}) without the end message. Saved ${n} as PARTIAL - press Start to retry.`); return; }
        let msg = `Loaded and saved ${n} transactions back to ${oldestDate()}.`;
        report(msg);
        if (opts.paid) { const m = await openPaid(c, opts.all); msg += ` Opened ${m} paid orders.`; onChange(); }
        report(msg + (ctl.stop ? ' STOPPED early.' : ' COMPLETE - you can download now.'));
      } catch (e) { report('Error: ' + (e && e.message || e)); } finally { ctl.end(); }
    }
    function readyMessage() {
      const c = ctx(), e = c.region ? getScan('tx', c.region, c.account, '') : null;
      if (!c.region) return `Ready. ${cardCount()} transactions visible.`;
      if (!c.account) return 'Account not detected yet - open the profile menu once (or visit library.playstation.com while signed in), then press Start.';
      if (!e) return 'Nothing stored yet. Press ▶ Start.';
      if (e.complete) return `✔ Stored: ${e.rows} transactions${e.details ? ', ' + e.details + ' orders read' : ''}. Press Start to refresh.`;
      return `⚠ Partial: ${e.rows} transactions. Press Start to finish.`;
    }
    return { start, cardCount, guessRegion, readyMessage };
  }

  // =====================================================================================
  //  WIRING
  // =====================================================================================
  if (location.hostname === 'library.playstation.com' && !inFrame) {
    // =========================== library page: games scan + read-only view of the shared transaction data ===========================
    const panel = makePanel('lib', 'PSN Library Exporter <span style="opacity:.45;font-weight:400;font-size:10px">8.6.0</span>', 300, { x: 12, y: 60 }, true);
    const ctl = controller(), ui = panelUI(panel.bd);
    let acc, lastReady = '';
    const lib = libraryEngine(ui, ctl, () => refreshAll());
    const who = () => { const r0 = lib.pageRegion(), id = readIdentity(r0); return { region: r0 || id.region, account: id.account }; };
    const label = w => acctLabel(acctId(w.region, w.account));
    const libReady = () => { const w = who(); if (!w.region || !w.account) return 'Idle.'; const e = getScan('lib', w.region, w.account, TAB);
      return !e ? 'Not scanned yet. Press ▶ Start.' : e.complete ? `✔ Scanned: ${e.rows}${e.total ? '/' + e.total : ''} games. Press Start to refresh.` : `⚠ Partial: ${e.rows}${e.total ? '/' + e.total : ''} games. Press Start to resume.`; };
    const txLine = () => { const w = who(); if (!w.region || !w.account) return ''; const e = getScan('tx', w.region, w.account, '');
      return e ? `Transactions: ${e.complete ? '✔ complete' : '⚠ partial'} · ${e.rows}` : 'Transactions: not scanned (use the Order History drawer).'; };
    ui.button('▶ Start', () => lib.run(), '#1d9d57', 'Scan the Purchased tab for this account (resumes a stopped scan)');
    pauseToggle(ui, ctl);
    ui.button('⏹', () => { ctl.stop = true; ctl.paused = false; }, '#c0392b', 'Stop (what was captured is kept; Start resumes)');
    ui.button('⬇ Games', () => { const w = who(); downloadLib(acctId(w.region, w.account), t => ui.set(t)); }, '#2185f4', "Download this account's game library CSV", 2);
    ui.button('⬇ Transactions', () => { const w = who(); downloadTx(acctId(w.region, w.account), t => ui.set(t)); }, '#2185f4', "Download this account's transaction history CSV", 2);
    acc = accountsView(ui.pane('Accounts'), [
      { kind: 'lib', title: 'Games', dl: id => downloadLib(id, t => ui.set(t)) },
      { kind: 'tx', title: 'Transactions', dl: id => downloadTx(id, t => ui.set(t)) }]);
    const del = ui.pane('Delete');
    del.body.innerHTML = '<div style="margin:1px 0"><b>Games</b></div>'; const d1 = document.createElement('div'); del.body.appendChild(d1);
    smallBtn(d1, 'This account', () => { const w = who(); lib.clearAccount(w.region, w.account); }, 'Delete stored games for this account'); smallBtn(d1, 'All accounts', () => lib.clearAll(), 'Delete stored games for every account');
    smallBtn(d1, 'Rescan from page 1', () => { const w = who(); lib.resetCheckpoint(w.region, w.account); }, 'Forget the resume point: the next Start scans everything from page 1');
    const t2 = document.createElement('div'); t2.innerHTML = '<b>Transactions</b>'; t2.style.margin = '3px 0 1px'; del.body.appendChild(t2); const d2 = document.createElement('div'); del.body.appendChild(d2);
    smallBtn(d2, 'This account', () => { const w = who(); clearTxAccount(w.region, w.account, t => ui.set(t), refreshAll); }, 'Delete stored transactions for this account'); smallBtn(d2, 'All accounts', () => clearTxAll(t => ui.set(t), refreshAll), 'Delete stored transactions for every account');
    ui.help(['Open <b>Purchased</b> (library.playstation.com › Purchased), press <b>▶ Start</b> and keep this tab in front.', 'Your account is detected automatically. A stopped or failed scan resumes from its last page (Delete › Rescan from page 1 to start over).', 'Transactions: profile menu › <b>Order History</b>, then press Start in the Transactions Exporter inside the drawer. The result shows up here.', 'Switched PSN account? Sign in, reload, repeat.']);
    ui.set('Idle.');
    const expired = cur => /^(Downloaded|Cleared|No stored|All stored|Next Start)/.test(cur) && ui.age() > 6000;
    function refreshAll() {
      lib.reload(); acc.refresh();
      const w = who(); ui.identity(w.region ? `<b>${esc(label(w))}</b>${w.account ? '' : ' <span style="opacity:.7">(account not detected)</span>'}` : '<span style="opacity:.7">detecting account…</span>');
      if (!ctl.running) { const cur = ui.text(), m = libReady(); if ((m !== lastReady && (cur === '' || cur === 'Idle.' || cur === lastReady)) || expired(cur)) { lastReady = m; ui.set(m); } }
      ui.sub(txLine());
    }
    window.addEventListener('beforeunload', e => { if (ctl.running) { e.preventDefault(); e.returnValue = ''; } });
    refreshAll(); setInterval(refreshAll, 2000); store.onChange(refreshAll);
  } else if (isDrawer) {
    // =========================== checkout site ===========================
    if (!inFrame) {
      // top-level checkout page: share the account shown in its own header, answer the drawer, and show a harmless hint while the drawer is closed
      let tip = null;
      const drawerShown = () => { const f = document.querySelector('iframe[name="embeddedcart"]'); return !!(f && f.getBoundingClientRect().width > 100); };
      const busy = () => { const b = store.load(CFG.keyBusy); return !!(b.at && Date.now() - b.at < 12000); };
      const dismissed = () => { try { return sessionStorage.getItem('psnlx_tip_off') === '1'; } catch (e) { return false; } };
      setInterval(() => {
        const o = ownIdentity(); if (o.account) rememberIdentity({ account: o.account, region: o.region || '' });
        const want = /order-history/.test(location.pathname) && !drawerShown() && !document.querySelector('.transaction-history-card') && !busy() && !dismissed();
        if (want && !tip && document.body) {
          tip = document.createElement('div'); tip.setAttribute('data-psnlx', '1');
          tip.style.cssText = 'position:fixed;right:10px;bottom:10px;z-index:2147483647;background:#0b1e3a;color:#fff;font:12px/1.3 system-ui,sans-serif;padding:6px 9px;border-radius:7px;max-width:250px;box-shadow:0 3px 12px rgba(0,0,0,.4)';
          tip.innerHTML = 'PSNLX: the Transaction History drawer is not open. Open the profile menu → Order History once more. <span class="x" style="cursor:pointer;margin-left:6px;opacity:.8">✕</span>';
          ['mousedown', 'mouseup', 'click', 'pointerdown', 'pointerup', 'touchstart'].forEach(ev => tip.addEventListener(ev, e => e.stopPropagation()));
          tip.querySelector('.x').addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); try { sessionStorage.setItem('psnlx_tip_off', '1'); } catch (er) {} if (tip) { tip.remove(); tip = null; } });
          document.body.appendChild(tip);
        }
        if (!want && tip) { tip.remove(); tip = null; }
      }, 2000);
      window.addEventListener('message', e => {                           // fallback path (no GM storage): answer the drawer's hello
        const d = e.data; if (!d || d.psnlx !== 1 || d.type !== 'hello' || !e.source) return;
        const o = ownIdentity(); e.source.postMessage({ psnlx: 1, type: 'hydrate', ctx: { region: o.region || '', account: o.account || '' }, data: { [CFG.keyTx]: store.load(CFG.keyTx), [CFG.keyScans]: store.load(CFG.keyScans) } }, '*');
      });
    }

    const boot = setInterval(() => {
      if (!(document.querySelector('.transaction-history-screen-cards') || document.querySelector('.transaction-history-card'))) return;
      clearInterval(boot);
      const ctl = controller(); let lastBeat = 0;
      const panel = makePanel('tx', 'Transactions Exporter <span style="opacity:.45;font-weight:400;font-size:10px">8.6.0</span>', 290, { x: 8, y: Math.max(10, window.innerHeight - 250) }, false);
      const ui = panelUI(panel.bd);
      let engine, acc;
      const say = t => { ui.set(t); if (ctl.running && Date.now() - lastBeat > 3000) { lastBeat = Date.now(); store.save(CFG.keyBusy, { at: lastBeat }); } };
      const onChange = () => acc.refresh();
      const me = () => { const g = engine.guessRegion(), id = readIdentity(g); return { account: id.account, region: (id.region || g || '').toUpperCase(), source: id.source, at: id.at || '' }; };
      engine = txEngine(say, ctl, me, onChange);
      const showId = () => { const m = me();
        ui.identity(m.region ? `<b>${esc(acctLabel(acctId(m.region, m.account)))}</b>${m.account ? (m.source === 'shared' ? ` <span style="opacity:.7">(detected ${ago(m.at)})</span>` : '') : ' <span style="opacity:.7">(account not detected - open the profile menu once, or visit the library page)</span>'}` : '<span style="opacity:.7">detecting…</span>'); };
      let lastReady = '';
      const expired = cur => /^(Downloaded|Cleared|No stored|All stored)/.test(cur) && ui.age() > 6000;
      const ready = () => { showId(); if (!ctl.running) { const cur = ui.text(), m = engine.readyMessage(); if (cur === '' || cur === lastReady || cur === 'Idle.' || expired(cur)) { lastReady = m; ui.set(m); } } };
      ui.button('▶ Start', () => engine.start({ paid: true, all: true }), '#1d9d57', "Load everything, save it, and read every paid order's item prices");
      pauseToggle(ui, ctl);
      ui.button('⏹', () => { ctl.stop = true; ctl.paused = false; }, '#c0392b', 'Stop (what was loaded is saved as partial)');
      ui.button('⬇ Transactions', () => { const m = me(); downloadTx(acctId(m.region, m.account), t => ui.set(t)); }, '#2185f4', "Download this account's transactions", 2);
      acc = accountsView(ui.pane('Accounts'), [{ kind: 'tx', title: 'Transactions', dl: id => downloadTx(id, t => ui.set(t)) }]);
      const del = ui.pane('Delete'); const dd = document.createElement('div'); del.body.appendChild(dd);
      smallBtn(dd, 'This account', () => { const m = me(); clearTxAccount(m.region, m.account, t => ui.set(t), () => { onChange(); ready(); }); }, 'Delete stored transactions for this account');
      smallBtn(dd, 'All accounts', () => clearTxAll(t => ui.set(t), () => { onChange(); ready(); }), 'Delete stored transactions for every account');
      ui.help(['Your account is detected automatically (page header, or the library page).', "Press <b>▶ Start</b>: it fast-scrolls until the site says there is nothing more, saves everything, then reads each paid order's item prices. Keep this tab in front.", 'Open <b>Accounts</b> to download one account or all of them.']);
      window.addEventListener('beforeunload', e => { if (ctl.running) { e.preventDefault(); e.returnValue = ''; } });
      ready(); onChange(); setInterval(() => { showId(); acc.refresh(); if (!ctl.running) ready(); }, 3000); store.onChange(() => { onChange(); if (!ctl.running) ready(); });
      if (inFrame) {                                                       // non-GM fallback: identity + saved data come from the parent page
        window.addEventListener('message', e => {
          const d = e.data; if (!d || d.psnlx !== 1 || e.source !== window.parent || d.type !== 'hydrate') return;
          if (!HAS_GM) Object.keys(d.data || {}).forEach(k => { const cur = store.load(k), inc = d.data[k] || {}; store.save(k, Object.assign({}, inc, cur), true); });
          if (d.ctx && d.ctx.account && !HAS_GM) store.save(CFG.keyIdent, { account: d.ctx.account, region: d.ctx.region || '', at: new Date().toISOString() }, true);
          onChange(); ready();
        });
        try { window.parent.postMessage({ psnlx: 1, type: 'hello' }, '*'); } catch (e) {}
      }
    }, 1000);
  }
})();
