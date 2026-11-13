/* Freight Hub for Outlook (v4.16.1) · the panel.
   v4.16.1: previews, 3 tries before an error, a light load index (card before the Hub answers), memory per thread,
   repeated images = signature, instant Today, and now → email when the driver changes.
   Holds no data, no keys and no Hub address in code: those are typed in once and kept in Outlook roamingSettings (your mailbox).
   Reads only the email you have open. Nothing is saved, ticked or opened without a tap (unless you turn on auto-save).
   v4.16: one clear next step per email (driver info, status/ETA, detention/lumper, POD/RC/invoice), Today, Find, reply templates,
   load alerts → your carriers, save to a second load, a status line on the email, and a memory of recent emails. */
(function () {
  'use strict';
  const CLIENT = 'outlook-1.2';
  const TYPES = ['POD', 'Rate Con (Signed)', 'Invoice', 'BOL', 'Lumper', 'Scale', 'Other'];
  const PAPER = { 'POD': 'pod', 'Rate Con (Signed)': 'rc', 'Invoice': 'inv' };   // saving one of these also ticks its step
  const MAX_ONE_CALL = 25 * 1024 * 1024;
  const $ = s => document.querySelector(s);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const rid = () => { const a = new Uint8Array(12); crypto.getRandomValues(a); return 'r' + Array.from(a, b => b.toString(16).padStart(2, '0')).join(''); };
  const P = { url: '', key: '', name: '', switches: null, view: 'main', tab: 'email', seq: 0, mail: null, res: null, files: {}, showSig: false, conflicts: [], doneTexts: {}, busy: false, checks: null,
    ready: null, today: null, idx: null, quick: '', find: { q: '', res: null, mode: '' }, loadView: null, upd: null, details: false };

  /* ---------- memory: this session (full) + this computer (12 h, no file contents, no money, no key) ---------- */
  const MEMO = new Map(); const MEMO_MAX = 50; const MEM_KEY = 'fh.mem.v2'; const MEM_TTL = 12 * 3600e3;
  const memoKey = m => m && (m.internetMessageId || m.itemId);
  const ls = { get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }, set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* full or blocked */ } }, del(k) { try { localStorage.removeItem(k); } catch (e) { /* blocked */ } } };
  function memRead() { try { const o = JSON.parse(ls.get(MEM_KEY) || '{}'); const now = Date.now(); Object.keys(o).forEach(k => { if (!o[k] || now - o[k].t > MEM_TTL) delete o[k]; }); return o; } catch (e) { return {}; } }
  function memWrite(o) { const ks = Object.keys(o).sort((a, b) => o[a].t - o[b].t); while (ks.length > MEMO_MAX) delete o[ks.shift()]; ls.set(MEM_KEY, JSON.stringify(o)); }
  const strip = m => Object.assign({}, m, { atts: m.atts.map(a => ({ id: a.id, name: a.name, size: a.size, contentType: a.contentType, isInline: a.isInline, sha256: a.sha256 })) });
  function remember() { const k = memoKey(P.mail); if (!k || !P.res) return; const R = P.res; if (R.loadId && (R.already || (R.match && (R.match.certain || R.match.picked)))) thrSet(P.mail.conversationId, R.loadId); MEMO.delete(k); MEMO.set(k, { mail: P.mail, res: P.res, files: JSON.parse(JSON.stringify(P.files)) }); while (MEMO.size > MEMO_MAX) MEMO.delete(MEMO.keys().next().value);
    const o = memRead(); o[k] = { t: Date.now(), mail: strip(P.mail), res: P.res, files: P.files }; memWrite(o); }
  function forget() { const k = memoKey(P.mail); if (!k) return; MEMO.delete(k); const o = memRead(); if (o[k]) { delete o[k]; memWrite(o); } }
  function recall(k) { const a = MEMO.get(k); if (a) return a; const o = memRead()[k]; if (!o) return null; o.mail.atts.forEach(f => { f.read = false; }); return o; }

  /* ---------- settings in the mailbox (the key never goes in localStorage or a URL) ---------- */
  const rs = () => Office.context.roamingSettings;
  function loadSettings() { P.url = rs().get('hubUrl') || ''; P.key = rs().get('hubKey') || ''; P.name = rs().get('hubName') || ''; P.details = ls.get('fh.details') === '1'; }
  function saveSettings(o) { return new Promise((ok, bad) => { Object.keys(o).forEach(k => { if (o[k] == null) rs().remove(k); else rs().set(k, o[k]); }); rs().saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? ok() : bad(r.error)); }); }
  const sleep = ms => new Promise(z => setTimeout(z, ms)); const RETRY_WAIT = [1000, 3000];
  /** v4.16.1: up to 3 tries when the Hub can't be reached, with the same request ID (the Hub keeps each answer for 6 h, so a save is never done twice) */
  async function api(action, data, ms, requestId) { const id = requestId || rid(); let last;
    for (let i = 0; i < 3; i++) { try { const r = await apiOnce(action, data, ms, id); if (i) retryNote(''); return r; }
      catch (e) { last = e; if (!e.net || i === 2) break; retryNote('The Hub didn\'t answer · trying again (' + (i + 2) + ' of 3)…'); await sleep((window.__fhRetryWait || RETRY_WAIT)[i]); } }
    retryNote(''); throw last; }
  function retryNote(t) { let el = document.getElementById('retry-note'); if (!t) { if (el) el.remove(); return; } if (!el) { el = document.createElement('div'); el.id = 'retry-note'; el.className = 'note'; el.setAttribute('role', 'status'); const app = $('#app'); app.parentNode.insertBefore(el, app); } el.textContent = t; }
  async function apiOnce(action, data, ms, requestId) {
    if (!P.url || !P.key) throw Object.assign(new Error('Not connected'), { auth: true });
    const ctl = new AbortController(); const t = setTimeout(() => ctl.abort(), ms || 20000);
    try {
      const r = await fetch(P.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ key: P.key, action, requestId: requestId || rid(), client: CLIENT }, data || {})), redirect: 'follow', signal: ctl.signal });
      if (r.status >= 500) throw Object.assign(new Error('The Hub had a problem (' + r.status + ').'), { net: true }); let j; try { j = await r.json(); } catch (x) { throw Object.assign(new Error('The Hub sent back a page instead of an answer.'), { net: true }); } if (j && j.ok === false) { const e = new Error(j.message || 'The Hub said no.'); e.auth = !!j.auth; throw e; } return j;
    } catch (e) { if (e.name === 'AbortError') throw Object.assign(new Error('Can\'t reach the Hub (it took too long).'), { net: true }); if (e instanceof TypeError) throw Object.assign(new Error('Can\'t reach the Hub.'), { net: true }); throw e; }
    finally { clearTimeout(t); }
  }

  /* ---------- reading the open email ---------- */
  const cb = (fn, ...a) => new Promise((ok, bad) => fn(...a, r => r.status === Office.AsyncResultStatus.Succeeded ? ok(r.value) : bad(r.error || new Error('Outlook said no'))));
  const sup = v => { try { return Office.context.requirements.isSetSupported('Mailbox', v); } catch (e) { return false; } };
  const b64bytes = b64 => { const s = atob(b64); const a = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) a[i] = s.charCodeAt(i); return a; };
  async function sha256(bytes) { const h = await crypto.subtle.digest('SHA-256', bytes); return Array.from(new Uint8Array(h), b => b.toString(16).padStart(2, '0')).join(''); }
  async function readItem() {
    const it = Office.context.mailbox.item; if (!it) return null;
    const addr = x => x ? (x.emailAddress || '') : ''; const list = a => (a || []).map(addr).join(', ');
    let body = ''; try { body = await cb(it.body.getAsync.bind(it.body), Office.CoercionType.Text); } catch (e) { body = ''; }
    const m = { internetMessageId: it.internetMessageId || '', conversationId: it.conversationId || '', itemId: it.itemId || '', subject: it.subject || '', from: addr(it.from), fromName: it.from ? it.from.displayName || '' : '', to: list(it.to), cc: list(it.cc),
      received: it.dateTimeCreated ? new Date(it.dateTimeCreated).toISOString().slice(0, 16) : '', body: String(body || '').slice(0, 20000), atts: [] };
    for (const a of (it.attachments || [])) { if (a.attachmentType && a.attachmentType !== 'file' && a.attachmentType !== Office.MailboxEnums.AttachmentType.File) continue;
      m.atts.push({ id: a.id, name: a.name, size: a.size, contentType: a.contentType, isInline: !!a.isInline, sha256: '', base64: '' }); }
    return m;
  }
  /** attachment contents and fingerprints, read in parallel behind the load card; stops if you move to another email */
  async function readContents(m) { const it = Office.context.mailbox.item; if (!it || memoKey({ internetMessageId: it.internetMessageId, itemId: it.itemId }) !== memoKey(m)) return;
    await Promise.all(m.atts.filter(f => !f.read).map(async f => {
      if (sup('1.8') && f.size <= MAX_ONE_CALL) { try { const c = await cb(it.getAttachmentContentAsync.bind(it), f.id); if (c.format === 'base64' || c.format === Office.MailboxEnums.AttachmentContentFormat.Base64) { f.base64 = c.content; f.sha256 = await sha256(b64bytes(c.content)); } } catch (e) { /* unreadable: can't be saved */ } }
      f.read = true; })); try { sigLearn(m); } catch (e) { /* storage blocked */ } }
  async function emailCopy() { const it = Office.context.mailbox.item;
    if (sup('1.14') && it.getAsFileAsync) { try { return { eml: await cb(it.getAsFileAsync.bind(it)) }; } catch (e) { /* fall back */ } }
    let headers = ''; if (sup('1.8') && it.getAllInternetHeadersAsync) { try { headers = await cb(it.getAllInternetHeadersAsync.bind(it)); } catch (e) { headers = ''; } }
    let html = ''; try { html = await cb(it.body.getAsync.bind(it.body), Office.CoercionType.Html); } catch (e) { html = ''; }
    return { html, headers }; }

  /* ---------- v4.16.1: a light index of active loads (load cards only, no money) → the card shows before the Hub answers ---------- */
  const IDX_KEY = 'fh.idx.v1', IDX_TTL = 10 * 60e3, THR_KEY = 'fh.thr.v1', THR_TTL = 30 * 864e5, SIG_KEY = 'fh.sig.v1';
  function idxLoad() { if (P.idx) return P.idx; try { const o = JSON.parse(ls.get(IDX_KEY) || 'null'); if (o && Date.now() - o.t < IDX_TTL) { P.idx = o; if (o.today && !P.today) P.today = o.today; return o; } } catch (e) { /* none */ } return null; }
  async function idxRefresh() { try { const r = await api('index', { withToday: true }, 25000); P.idx = { t: Date.now(), loads: r.loads || [], today: r.today || null }; ls.set(IDX_KEY, JSON.stringify(P.idx));
      if (r.today) { P.today = r.today; tabs(); if (P.tab === 'today' && P.view === 'main') renderToday(false); } return true; }
    catch (e) { if (/Unknown request/i.test(e.message || '')) { try { P.today = await api('today', {}, 20000); tabs(); } catch (x) { /* Today off */ } } return false; } }
  function thrRead() { try { return JSON.parse(ls.get(THR_KEY) || '{}'); } catch (e) { return {}; } }
  function thrGet(c) { if (!c) return ''; const v = thrRead()[c]; return v && Date.now() - v.t < THR_TTL ? v.id : ''; }
  function thrSet(c, id) { if (!c || !id) return; const o = thrRead(); if (o[c] && o[c].id === id && Date.now() - o[c].t < 864e5) return; o[c] = { id, t: Date.now() }; const ks = Object.keys(o).sort((a, b) => o[a].t - o[b].t); while (ks.length > 300) delete o[ks.shift()]; ls.set(THR_KEY, JSON.stringify(o)); }
  /** the card from the index while the Hub checks: a load # in the subject or body, or this thread's load */
  function quick(m) { const I = idxLoad(); if (!I || !(I.loads || []).length || P.res) return; const by = id => I.loads.find(L => L.id === id);
    const ids = s => (String(s || '').match(/\b\d{5}-\d{5}\b/g) || []).filter((x, i, a) => a.indexOf(x) === i);
    const s = ids(m.subject).filter(by), bd = ids(String(m.body || '').slice(0, 4000)).filter(by); let id = '', why = '';
    if (s.length === 1) { id = s[0]; why = 'Load # in subject'; } else if (!s.length && bd.length === 1) { id = bd[0]; why = 'Load # in the email'; }
    else if (!s.length && !bd.length) { const t = thrGet(m.conversationId); if (t && by(t)) { id = t; why = 'Same thread as an email saved to ' + t; } }
    if (!id) return; P.quick = id; $('#app').innerHTML = loadCard(by(id), `<span class="chip ok">${esc(why)}</span>`) + '<div class="card muted"><span class="spin" aria-hidden="true"></span>Checking the Hub…</div>'; }
  /** a picture that came in an earlier email from the same sender (logo, "Bonded carrier" banner) is a signature, whatever its size */
  const PUBLIC_MAIL = /@(gmail|yahoo|outlook|hotmail|icloud|aol|live|msn|comcast|att|ymail|protonmail)\./i;
  const sigOwner = m => { const f = String(m && m.from || '').toLowerCase(); return PUBLIC_MAIL.test(f) ? f : (f.split('@')[1] || f); };
  const isImg = a => /^image\//i.test(a.contentType || '') || /\.(png|jpe?g|gif|bmp|webp)$/i.test(a.name || '');
  const isPdf = a => /pdf/i.test(a.contentType || '') || /\.pdf$/i.test(a.name || '');
  function sigStore() { try { return JSON.parse(ls.get(SIG_KEY) || '{}'); } catch (e) { return {}; } }
  function markRepeat(files) { const m = P.mail; if (!m || !files) return; const S = sigStore()[sigOwner(m)] || {}; const me = memoKey(m);
    files.forEach(f => { const a = (m.atts || []).find(x => x.id === f.id); if (a && a.sha256 && isImg(a) && (S[a.sha256] || []).some(id => id !== me)) { f.signature = true; f.repeat = true; } }); }
  function sigLearn(m) { const own = sigOwner(m); if (!own) return; const o = sigStore(); const S = o[own] = o[own] || {}; const me = memoKey(m); let ch = false;
    (m.atts || []).forEach(a => { if (a.sha256 && isImg(a)) { const l = S[a.sha256] = S[a.sha256] || []; if (l.indexOf(me) < 0) { l.push(me); if (l.length > 5) l.shift(); ch = true; } } });
    if (ch) { const ks = Object.keys(o); while (ks.length > 200) delete o[ks.shift()]; ls.set(SIG_KEY, JSON.stringify(o)); } }
  /* previews: pictures straight from the email; PDFs page 1 through pdf.js (loaded only when a PDF is open) */
  const PDFV = '3.11.174'; let pdfjs = null; const PDFC = new Map();
  function pdfLib() { if (pdfjs) return pdfjs; pdfjs = new Promise((ok, bad) => { const s = document.createElement('script'); s.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/' + PDFV + '/pdf.min.js';
    s.onload = () => { try { window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/' + PDFV + '/pdf.worker.min.js'; ok(window.pdfjsLib); } catch (e) { bad(e); } }; s.onerror = () => bad(new Error('no pdf.js')); document.head.appendChild(s); });
    pdfjs.catch(() => {}); return pdfjs; }
  async function pdfPage1(a) { const k = a.sha256 || a.id; if (PDFC.has(k)) return PDFC.get(k); const p = (async () => { const lib = await pdfLib(); const doc = await lib.getDocument({ data: b64bytes(a.base64) }).promise; const pg = await doc.getPage(1);
      const v = pg.getViewport({ scale: 1 }); const sc = Math.min(2, 900 / v.width); const vp = pg.getViewport({ scale: sc }); const cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height;
      await pg.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise; return { url: cv.toDataURL('image/jpeg', 0.85), pages: doc.numPages }; })(); PDFC.set(k, p); p.catch(() => PDFC.set(k, Promise.resolve(null))); return p; }
  function thumb(f) { const a = P.mail && (P.mail.atts || []).find(x => x.id === f.id); if (!a) return '';
    if (!a.base64) return (isImg(a) || isPdf(a)) && !a.read ? '<span class="thumb ph" aria-hidden="true"></span>' : '';
    if (isImg(a)) return `<button class="thumb" type="button" data-view="${esc(a.id)}" aria-label="Preview ${esc(a.name)}"><img alt="" src="data:${esc(a.contentType || 'image/png')};base64,${a.base64}" onerror="this.parentNode.innerHTML='<span class=&quot;ic&quot;>IMG</span>'"></button>`;
    if (isPdf(a)) { const k = a.sha256 || a.id; const hit = PDFC.get(k); if (hit && hit.done) return `<button class="thumb" type="button" data-view="${esc(a.id)}" aria-label="Preview ${esc(a.name)}">${hit.done.url ? `<img alt="" src="${hit.done.url}">` : '<span class="ic">PDF</span>'}</button>`;
      if (!hit || !hit.wait) { const my = P.seq; pdfPage1(a).then(v => { PDFC.get(k).done = v || { url: '' }; if (my === P.seq) render(); }, () => { const x = PDFC.get(k) || {}; x.done = { url: '' }; PDFC.set(k, x); if (my === P.seq) render(); }); PDFC.get(k).wait = true; }
      return `<button class="thumb" type="button" data-view="${esc(a.id)}" aria-label="Preview ${esc(a.name)}"><span class="ic">PDF</span></button>`; }
    return ''; }
  async function view(id) { const a = P.mail && P.mail.atts.find(x => x.id === id); if (!a || !a.base64) return; const v = $('#viewer'); v.hidden = false;
    const bar = (extra) => `<div class="vbar"><span>${esc(a.name)}${extra ? ' · ' + esc(extra) : ''}</span><button type="button" data-act="vclose">Close</button></div>`;
    if (isImg(a)) { v.innerHTML = bar('') + `<div class="vbody"><img alt="${esc(a.name)}" src="data:${esc(a.contentType || 'image/png')};base64,${a.base64}"></div>`; }
    else { v.innerHTML = bar('') + '<div class="vbody"><div class="vmsg"><span class="spin" aria-hidden="true"></span>Opening…</div></div>'; let r = null; try { r = await pdfPage1(a); } catch (e) { r = null; } if (v.hidden) return;
      v.innerHTML = r && r.url ? bar('page 1 of ' + r.pages) + `<div class="vbody"><img alt="${esc(a.name)} page 1" src="${r.url}"></div>` : bar('') + '<div class="vbody"><div class="vmsg">No preview for this PDF here. It still saves normally.</div></div>'; }
    const b = v.querySelector('[data-act="vclose"]'); if (b) b.focus(); }
  function viewClose() { const v = $('#viewer'); if (v) { v.hidden = true; v.innerHTML = ''; } }
  document.addEventListener('keydown', e => { if (e.key === 'Escape') viewClose(); });

  /* ---------- the status line on the email itself (S4) ---------- */
  function statusLine() { const it = Office.context.mailbox.item; const r = P.res; if (!it || !it.notificationMessages || !r) return; let msg = '';
    if (r.load) { const L = r.load; msg = 'Freight Hub · ' + L.id + ' · ' + L.status + (r.driver ? ' · driver info found' : r.update ? ' · update: ' + r.update.status + (r.update.eta ? ' ETA ' + r.update.eta : '') : L.next ? ' · ' + L.next : (L.missing || []).length ? ' · missing ' + L.missing.join(', ') : ''); }
    else if (r.alert) msg = 'Freight Hub · load alert ' + r.alert.lane + ' · ' + r.alert.carriers.length + ' of your carriers ran it';
    else if ((r.candidates || []).length) msg = 'Freight Hub · ' + r.candidates.length + ' possible loads · pick one in the panel';
    try { if (!msg) return it.notificationMessages.removeAsync('fh', () => {}); const type = Office.MailboxEnums && Office.MailboxEnums.ItemNotificationMessageType ? Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage : 'informationalMessage';
      it.notificationMessages.replaceAsync('fh', { type, message: msg.slice(0, 150), icon: 'icon16', persistent: false }, () => {}); } catch (e) { /* older Outlook */ } }

  /* ---------- each email ---------- */
  async function onItem() {
    const my = ++P.seq; P.res = null; P.files = {}; P.conflicts = []; P.doneTexts = {}; P.showSig = false; P.mail = null; P.ready = null; P.upd = null; if (P.view !== 'main') return;
    if (P.tab !== 'email') { P.tab = 'email'; tabs(); }
    const it = Office.context.mailbox.item; const hit = it ? recall(memoKey({ internetMessageId: it.internetMessageId, itemId: it.itemId })) : null;
    if (hit) { P.mail = hit.mail; take(hit.res); P.files = JSON.parse(JSON.stringify(hit.files || P.files)); render(); }   // seen before → show it now, refresh behind it
    else $('#app').innerHTML = '<div class="card muted"><span class="spin" aria-hidden="true"></span>Finding the load…</div>';
    let m; try { m = hit ? hit.mail : await readItem(); } catch (e) { m = null; }
    if (my !== P.seq) return; if (!m) { $('#app').innerHTML = '<div class="card muted">Open an email to see its load.</div>'; return; }
    P.mail = m; P.ready = readContents(m); P.quick = ''; if (!hit) try { quick(m); } catch (e) { /* no index */ }
    await lookup(my, '', true);
  }
  /** one call for the card now; a second (only with attachments) once their fingerprints are read, for the duplicate check */
  async function lookup(my, loadId, auto) { const m = P.mail;
    try { const need = () => m.atts.some(f => !f.read);
      if (!need() || !P.res || loadId) { const r = await api('match', match(m, loadId), 20000); if (my !== P.seq) return; take(r); render(); }
      if (need()) { await P.ready; if (my !== P.seq) return; const r2 = await api('match', match(m, loadId || (P.res && P.res.match && P.res.match.picked ? P.res.loadId : '')), 20000); if (my !== P.seq) return; take(r2, true); render(); }
      remember(); statusLine();
      const r = P.res; if (auto && P.switches && P.switches.me && P.switches.me.auto && r.match && r.match.certain && !r.already && r.loadId) save(true);
    } catch (e) { if (my !== P.seq) return; if (P.res && !e.auth) return toast(e.message); if (e.auth) return setup(e.message); $('#app').innerHTML = `<div class="card"><div class="err">${esc(e.message)}</div><button class="btn" id="retry" type="button" style="margin-top:10px">Try again</button></div>`; $('#retry').onclick = onItem; }
  }
  function take(r, again) { const keepUpd = again && P.upd; try { markRepeat(r.files); } catch (e) { /* storage blocked */ } P.res = r; P.switches = r.switches || P.switches; if (!again) P.files = {};
    (r.files || []).forEach(f => { const st = P.files[f.id]; if (st && (st.mine || st.saved)) return; P.files[f.id] = { type: f.type, tick: f.state === 'new' && !f.signature }; });
    if (!keepUpd) P.upd = r.update ? { status: r.update.status, eta: r.update.eta, note: r.update.note } : null; }
  function match(m, loadId) { return { internetMessageId: m.internetMessageId, conversationId: m.conversationId, subject: m.subject, from: m.from, body: m.body, loadId: loadId || '', attachments: m.atts.map(a => ({ id: a.id, name: a.name, size: a.size, contentType: a.contentType, isInline: a.isInline, sha256: a.sha256 })) }; }
  async function pickLoad(id) { const my = ++P.seq; P.res = null; P.tab = 'email'; tabs(); $('#app').innerHTML = '<div class="card muted"><span class="spin" aria-hidden="true"></span>Opening ' + esc(id) + '…</div>'; await lookup(my, id, false); }

  /* ---------- tabs ---------- */
  function tabs() { const t = $('#tabs'); if (!t) return; t.hidden = P.view !== 'main'; const n = P.today ? P.today.total : 0;
    t.innerHTML = [['email', 'This email', ''], ['today', 'Today', n ? `<span class="badge">${n}</span>` : ''], ['find', 'Find', '']].map(x => `<button class="tab" role="tab" type="button" data-tab="${x[0]}" aria-selected="${P.tab === x[0]}">${x[1]}${x[2]}</button>`).join(''); }
  async function showTab(k) { P.tab = k; tabs(); if (k === 'email') { if (P.res) render(); else onItem(); return; } if (k === 'today') return renderToday(true); if (k === 'find') { P.find.mode = P.find.mode || ''; return renderFind(); } }

  /* ---------- render: This email ---------- */
  const plural = (n, w) => n + ' ' + w + (n === 1 ? '' : 's');
  const localISO = () => { const d = new Date(), z = n => String(n).padStart(2, '0'); return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate()) + 'T' + z(d.getHours()) + ':' + z(d.getMinutes()); };
  function when(s) { const m = /^(\d{4})-(\d{2})-(\d{2})T?(\d{2})?:?(\d{2})?/.exec(String(s || '')); if (!m) return String(s || ''); const h = +(m[4] || 0); return (+m[2]) + '/' + (+m[3]) + (m[4] ? ' ' + ((h % 12) || 12) + ':' + m[5] + ' ' + (h < 12 ? 'AM' : 'PM') : ''); }
  function loadCard(L, chip, opt) { opt = opt || {}; const open = P.details || opt.open;
    const rows = [['Pickup', L.pickup], ['Delivery', L.delivery], ['Carrier', L.carrier], ['Dispatch', L.dispatch], ['Driver', L.driver], ['Truck · trailer', [L.truck, L.trailer].filter(Boolean).join(' · ')], ['ETA', L.eta]].filter(x => x[1]);
    return `<section class="card load" aria-label="Load ${esc(L.id)}"><div class="row">${chip || ''}${opt.change ? '<button class="lk" type="button" data-act="find" style="margin-left:auto">Not this load?</button>' : ''}</div>
      <div class="lid">${esc(L.id)} <span class="status">${esc(L.status)}</span></div><div class="lane">${esc(L.lane)}</div>
      <div class="sub">${esc([L.carrier, L.driver ? 'Driver ' + L.driver.split(' · ')[0] : ''].filter(Boolean).join(' · '))}</div>
      ${L.next ? `<div class="next">Next: ${esc(L.next)}</div>` : L.missing && L.missing.length ? `<div class="next">Missing: ${esc(L.missing.join(', '))}</div>` : ''}${L.issue ? `<div class="line">Open issue: ${esc(L.issue)}</div>` : ''}
      <details ${open ? 'open' : ''} data-details="1"><summary>Details</summary><div class="kv">${rows.map(x => `<span>${x[0]}</span><span>${esc(x[1])}</span>`).join('')}</div></details></section>`; }
  function render() { if (P.tab !== 'email' || P.view !== 'main') return; const r = P.res, m = P.mail; if (!r || !m) return; const sw = P.switches || {}; const me = sw.me || {}; const h = [];
    if (me.auto) h.push('<div class="note">Auto-save is on: certain matches save when you open them.</div>');
    if (r.load) { const ok = r.match.certain || r.match.picked; h.push(loadCard(r.load, `<span class="chip ${ok ? 'ok' : 'warn'}">${esc(r.match.reason)}</span>`, { change: true })); }
    else if ((r.candidates || []).length) h.push(`<section class="card pick" aria-label="Pick the load"><b>Which load is this for?</b><div class="line muted">${esc(r.match.reason)}</div><div class="stack" style="margin-top:10px">${r.candidates.map(c => `<button class="btn act" type="button" data-pick="${esc(c.id)}"><b>${esc(c.id)}</b><div class="line muted">${esc(c.lane)} · ${esc(c.status)}</div></button>`).join('')}</div><button class="lk" type="button" data-act="find">None of these · find another</button></section>`);
    else if (r.alert) h.push(alertHTML(r.alert));
    else h.push(`<section class="card"><b>No load found</b><div class="line muted">${r.party ? esc(r.match.reason) : 'No load # here and the sender isn\'t a carrier or customer in the Hub.'}</div><div class="find-inline"><input class="inp" id="qi" type="search" placeholder="Load #, carrier or city" aria-label="Find a load"><button class="btn" type="button" data-act="findgo">Find</button></div></section>`);
    if (r.already) { const ids = r.already.loadIds && r.already.loadIds.length ? r.already.loadIds : [r.already.loadId]; h.push(`<div class="saved" role="status"><span aria-hidden="true">✓</span><span>Saved to ${esc(ids.join(' and '))} · ${esc(when(r.already.at))} · ${esc(r.already.by)}</span></div>`); }
    if (r.load) h.push(nextSteps(r, me));
    if (r.load && (r.drafts || []).length) h.push(`<section class="card" aria-label="Reply with a template"><span class="lab">Reply with</span><div class="chips">${r.drafts.map(t => `<button class="pill" type="button" data-tpl="${esc(t.id)}">${esc(t.name)}</button>`).join('')}</div></section>`);
    if (r.load && r.already && sw.also) h.push(`<button class="lk" type="button" data-act="also">Also save this email to another load</button>`);
    $('#app').innerHTML = h.join('');
  }
  /** one primary button per email, picked in this order: driver info › detention/lumper › save paperwork › status update › save the email. The rest are smaller. */
  function leadOf(r) { const tickedNew = (r.files || []).some(f => (P.files[f.id] || {}).tick && !(P.files[f.id] || {}).saved);
    if (r.driver && !P.doneTexts.driver) return 'driver'; const i = r.issueHint; if (i && !P.doneTexts[i.kind] && !(i.open && i.kind === 'det')) return 'issue';
    if (tickedNew) return 'files'; if (r.update && P.upd && !P.doneTexts.update) return 'update'; return r.already ? '' : 'files'; }
  function nextSteps(r, me) { const h = []; const lead = leadOf(r); const B = on => on ? 'btn pri big' : 'btn big sec'; const blocks = {};
    if (r.driver) { const d = r.driver; const f = d.fields; const now = d.now || {}; const rows = [['Driver', 'Driver'], ['Phone', 'Driver Phone'], ['Truck', 'Truck'], ['Trailer', 'Trailer']].map(x => [x[0], f[x[1]], now[x[1]] || '']).concat([['Tracking', d.link ? d.link.provider : '', '']]).filter(x => x[1]); const rep = rows.some(x => x[2]);
      blocks.driver = P.doneTexts.driver ? done(P.doneTexts.driver) : `<section class="card ${lead === 'driver' ? 'hi' : ''}" aria-label="Driver info found"><span class="lab">${rep ? 'New driver info in this email' : 'Driver info in this email'}</span><div class="found">${rows.map(x => `<span>${x[0]}</span><b>${esc(x[1])}${x[2] ? `<span class="was">now ${esc(x[2])}</span>` : ''}</b>`).join('')}</div><button class="${B(lead === 'driver')}" type="button" data-do="driver">${rep ? 'Replace driver info' : 'Fill the load'} (${plural(rows.length, 'field')})</button></section>`; }
    if (r.issueHint) { const i = r.issueHint; const k = i.kind; const on = lead === 'issue';
      blocks.issue = P.doneTexts[k] ? done(P.doneTexts[k]) : i.open && k === 'det' ? done('Detention issue ' + i.open + ' is already open') : `<section class="card ${on ? 'hi' : ''}" aria-label="${esc(i.type)}"><span class="lab">${k === 'det' ? 'Detention' : 'Lumper'}</span>
        ${k === 'det' ? `<div class="found"><span>Stop</span><b>${esc(i.stop)}</b>${i.arrived ? `<span>Arrived</span><b>${esc(i.arrived)}</b>` : ''}<span>Free time</span><b>${esc(i.free)} h (customer rule)</b>${i.starts ? `<span>Starts</span><b>${esc(i.starts)}</b>` : ''}</div>` : `<div class="line">${i.open ? 'Adds this email to ' + esc(i.open) + '. Save the receipt below as Lumper.' : 'Starts a Lumper issue. Save the receipt below as Lumper.'}</div>`}
        <button class="${B(on)}" type="button" data-do="${k}">${k === 'det' ? 'Start detention issue' : i.open ? 'Add to Lumper issue ' + esc(i.open) : 'Start Lumper issue'}</button>${k === 'det' && (r.drafts || []).some(t => t.id === 'det' && t.to.length) ? '<button class="lk" type="button" data-tpl="det">Email the customer</button>' : ''}</section>`; }
    if (r.update && P.upd) { const u = P.upd; const on = lead === 'update'; const sts = ['Dispatched', 'At Pickup', 'Picked Up', 'En Route', 'Delayed', 'At Delivery', 'Delivered'];
      const form = `<div class="grid2"><label class="f" for="ust">Status</label><select class="inp" id="ust">${sts.map(s => `<option${s === u.status ? ' selected' : ''}>${s}</option>`).join('')}</select>
        <label class="f" for="ueta">ETA</label><input class="inp" id="ueta" value="${esc(u.eta || '')}" placeholder="e.g. 11/13 1500 PST"></div>
        <label class="f" for="unote">Note</label><input class="inp" id="unote" value="${esc(u.note || '')}">
        <button class="${B(on)}" type="button" data-do="update">Save update · log check call</button>${(r.drafts || []).some(t => t.id === 'eta' && t.to.length) ? '<button class="lk" type="button" data-tpl="eta">Then email the customer</button>' : ''}`;
      blocks.update = P.doneTexts.update ? done(P.doneTexts.update) : on ? `<section class="card hi" aria-label="Quick update"><span class="lab">Update from this email</span>${form}</section>`
        : `<details class="card" aria-label="Quick update"><summary>Log as a check call: ${esc(u.status)}${u.eta ? ' · ETA ' + esc(u.eta) : ''}</summary>${form}</details>`; }
    blocks.files = filesHTML(r, me, lead === 'files');
    ['driver', 'issue', 'files', 'update'].sort((a, b) => (b === lead) - (a === lead)).forEach(k => { if (blocks[k]) h.push(blocks[k]); });
    if ((r.actions || []).length) { const types = Object.keys(P.files).filter(id => P.files[id].tick && !P.files[id].saved).map(id => P.files[id].type); const auto = types.map(t => PAPER[t]).filter(Boolean);
      const rest = r.actions.filter(a => !(a.kind === 'cc' && r.update) && !(auto.indexOf(a.kind) >= 0 && !a.done && !P.doneTexts[a.kind]));
      if (rest.length) h.push(`<section class="stack" aria-label="One-tap actions">${rest.map(a => { const t = P.doneTexts[a.kind] || (a.done ? a.text : ''); if (t) return done(t);
        if (a.kind === 'cc') return `<details class="card"><summary>Log a check call</summary><label class="f" for="ccnote">Note</label><input class="inp" id="ccnote" value="${esc(a.note || '')}"><button class="btn act" type="button" data-do="cc" style="margin-top:8px">${esc(a.label)}</button></details>`;
        return `<button class="btn act" type="button" data-do="${esc(a.kind)}">${esc(a.label)}</button>`; }).join('')}</section>`); }
    return h.join(''); }
  const done = t => `<div class="done"><span aria-hidden="true">✓</span><span>${esc(t)}</span></div>`;
  function filesHTML(r, me, lead) { const files = r.files || []; const hide = me.hideSig !== false && !P.showSig; const shown = files.filter(f => !(hide && f.signature)); const hidden = files.length - shown.length; const h = [];
    if (files.length) h.push(`<section class="card" aria-label="Attachments"><span class="lab">Attachments</span>${shown.map((f, i) => { const st = P.files[f.id] || {}; const dis = f.state === 'dup' || f.state === 'noload' || !!st.saved; const txt = st.saved ? st.saved : f.repeat ? 'In earlier emails from this sender too · treated as a signature' : f.state === 'new' && st.type !== f.type ? 'New · will be saved as ' + st.type : f.text;
        return `<div class="att"><input type="checkbox" id="att${i}" data-tick="${esc(f.id)}" ${st.tick ? 'checked' : ''} ${dis ? 'disabled' : ''}>${thumb(f)}<div style="flex:1;min-width:0"><label class="nm" for="att${i}">${esc(f.name)}</label><div class="st ${esc(st.saved ? 'new' : f.state)}">${esc(txt)}</div></div><select class="type" data-type="${esc(f.id)}" aria-label="File type for ${esc(f.name)}" ${dis ? 'disabled' : ''}>${TYPES.map(t => `<option${t === (st.type || f.type) ? ' selected' : ''}>${esc(t)}</option>`).join('')}</select></div>`; }).join('')}
      ${hidden || P.showSig ? `<button class="lk" type="button" data-act="sig">${P.showSig ? 'Hide signature images' : plural(hidden, 'signature image') + ' hidden · show'}</button>` : ''}</section>`);
    P.conflicts.forEach(c => h.push(`<section class="conflict" aria-label="File conflict"><b>${esc(c.name)}</b><p>${esc(c.text)}</p><div class="row"><button class="btn pri" type="button" data-res="new" data-doc="${esc(c.docId)}">Save as ${esc(c.next)}</button><button class="btn" type="button" data-res="replace" data-doc="${esc(c.docId)}">Replace</button><button class="btn" type="button" data-res="skip" data-doc="${esc(c.docId)}">Skip</button></div></section>`));
    const ticked = shown.filter(f => (P.files[f.id] || {}).tick && !(P.files[f.id] || {}).saved); if (!r.load || (r.already && !ticked.length)) return h.join('');
    const steps = paperSteps(ticked); const what = ticked.length === 1 ? (P.files[ticked[0].id].type || 'file') : plural(ticked.length, 'file');
    const label = P.busy ? 'Saving…' : (r.already ? 'Save ' + what : ticked.length ? 'Save email + ' + what : 'Save email to load') + (steps.length ? ' · ' + steps.map(s => s.short).join(' · ') : '');
    h.push(`<button class="${lead || P.busy ? 'btn pri big' : 'btn big sec'}" type="button" data-act="save" ${P.busy ? 'disabled' : ''}>${esc(label)}</button>`); return h.join(''); }
  /** saving a POD, signed rate con or invoice also ticks its step (when that step isn't done yet) */
  function paperSteps(ticked) { const r = P.res; const acts = r.actions || []; const out = [];
    ticked.forEach(f => { const k = PAPER[(P.files[f.id] || {}).type]; const a = k && acts.find(x => x.kind === k); if (a && !a.done && !P.doneTexts[k] && !out.some(o => o.kind === k)) out.push({ kind: k, short: { pod: 'mark POD received', rc: 'mark RC signed', inv: 'mark invoice received' }[k] }); }); return out; }
  function alertHTML(a) { return `<section class="card" aria-label="Load alert"><span class="lab">Load alert</span><div class="lid">${esc(a.lane)}</div><div class="sub">${esc([a.pick && 'Pickup ' + a.pick, a.equip, a.ref && 'Ref ' + a.ref].filter(Boolean).join(' · '))}</div></section>
    <section class="card" aria-label="Your carriers on this lane"><span class="lab">Your carriers on this lane</span>${a.carriers.length ? a.carriers.map((c, i) => `<div class="carr"><div style="min-width:0"><b>${esc(c.name)}</b><div class="line muted">${c.exact ? plural(c.exact, 'load') + ' on this lane' : ''}${c.exact && c.near ? ' · ' : ''}${c.near ? c.near + ' state-to-state' : ''}${c.last ? ' · last ' + esc(when(c.last)) : ''}${c.active ? ' · ' + c.active + ' active now' : ''}</div>${c.blocked ? `<div class="line err">Safety: ${esc(c.safety)}</div>` : ''}</div><button class="btn sm" type="button" data-offer="${i}" ${c.blocked ? 'disabled' : ''}>${P.doneTexts['offer' + i] ? 'Drafted ✓' : 'Offer'}</button></div>`).join('') : '<div class="line muted">None of your carriers has run this lane yet.</div>'}
    <div class="line muted" style="margin-top:6px">Offer opens an email draft. No rate is in it unless you type one.</div></section>`; }

  /* ---------- Today (F3) ---------- */
  async function renderToday(fetch) { if (P.tab !== 'today') return; const t = P.today;
    if (t && fetch) { renderToday(false); api('today', {}, 20000).then(x => { P.today = x; tabs(); if (P.tab === 'today' && P.view === 'main') renderToday(false); }).catch(() => {}); return; }
    if (!t || fetch) { if (!t) $('#app').innerHTML = '<div class="card muted"><span class="spin" aria-hidden="true"></span>Checking your loads…</div>';
      try { P.today = await api('today', {}, 20000); tabs(); } catch (e) { if (!t) { $('#app').innerHTML = `<div class="card"><div class="err">${esc(e.message)}</div></div>`; return; } } if (P.tab !== 'today') return; }
    const T = P.today; if (!T.groups.length) { $('#app').innerHTML = '<section class="card"><b>All clear</b><div class="line muted">No missing PODs, check calls, driver info or rate cons right now.</div></section>'; return; }
    $('#app').innerHTML = T.groups.map(g => `<details class="card grp" ${g.key === (P.todayOpen || T.groups[0].key) ? 'open' : ''} data-grp="${esc(g.key)}"><summary><b>${esc(g.label)}</b> <span class="badge">${g.count}</span></summary>${g.loads.map(L => `<div class="li"><button class="lirow" type="button" data-open="${esc(L.id)}"><b>${esc(L.id)}</b> <span class="muted">${esc(L.carrier)}</span><div class="line muted">${esc(L.lane)} · ${esc(L.sub)}</div></button>${g.tpl ? `<button class="btn sm" type="button" data-ask="${esc(g.tpl)}" data-load="${esc(L.id)}">${esc(g.act)}</button>` : ''}</div>`).join('')}${g.count > g.loads.length ? `<div class="line muted">+ ${g.count - g.loads.length} more in the Hub</div>` : ''}</details>`).join('') + `<div class="line muted">From the Hub · ${esc(when(T.at))}${T.cached ? ' · updates every 2 min' : ''}</div>`; }

  /* ---------- Find (F8) and the load view ---------- */
  function renderFind() { if (P.tab !== 'find') return; const F = P.find; const also = F.mode === 'also';
    $('#app').innerHTML = `<section class="card">${also ? '<b>Also save this email to…</b>' : '<b>Find a load</b>'}<div class="find-inline"><input class="inp" id="q" type="search" placeholder="Load #, carrier or city" aria-label="Find a load" value="${esc(F.q)}"><button class="btn" type="button" data-act="findgo">Find</button></div>${also ? '<button class="lk" type="button" data-act="alsooff">Cancel</button>' : ''}</section><div id="fres">${findList()}</div>`;
    const q = $('#q'); q.focus(); q.onkeydown = e => { if (e.key === 'Enter') doFind(q.value); }; }
  function findList() { const R = P.find.res; if (!R) return ''; if (!R.loads.length) return '<div class="card muted">No loads match.</div>';
    return `<section class="card">${R.loads.map(L => `<div class="li"><button class="lirow" type="button" data-${P.find.mode === 'also' ? 'alsoto' : 'open'}="${esc(L.id)}"><b>${esc(L.id)}</b> <span class="muted">${esc(L.status)}</span><div class="line muted">${esc(L.lane)} · ${esc(L.carrier)} · ${esc(L.ship)}</div></button></div>`).join('')}${R.more ? `<div class="line muted">+ ${R.more} more · type more to narrow it</div>` : ''}</section>`; }
  let findT = 0; async function doFind(q) { P.find.q = q; const my = ++findT; try { const r = await api('find', { q }, 15000); if (my !== findT) return; P.find.res = r; const f = $('#fres'); if (f) f.innerHTML = findList(); else if (P.tab !== 'find') { P.tab = 'find'; tabs(); renderFind(); } } catch (e) { toast(e.message); } }
  async function openLoad(id) { P.tab = 'find'; tabs(); $('#app').innerHTML = '<div class="card muted"><span class="spin" aria-hidden="true"></span>Opening ' + esc(id) + '…</div>';
    try { const r = await api('load', { loadId: id }, 15000); P.loadView = r; const canUse = P.mail && P.res && !P.res.load && !P.res.already;
      $('#app').innerHTML = loadCard(r.load, '<span class="chip">From the Hub</span>', { open: true }) + (canUse ? `<button class="btn pri big" type="button" data-pick="${esc(id)}">Use this load for the open email</button>` : '') +
        ((r.drafts || []).length ? `<section class="card"><span class="lab">Email</span><div class="chips">${r.drafts.map(t => `<button class="pill" type="button" data-ltpl="${esc(t.id)}">${esc(t.name)}</button>`).join('')}</div></section>` : '') + '<button class="lk" type="button" data-act="back">‹ Back</button>'; }
    catch (e) { $('#app').innerHTML = `<div class="card"><div class="err">${esc(e.message)}</div></div>`; } }

  /* ---------- drafts (F5, offers, Today asks) ---------- */
  const html = t => esc(t).replace(/\n/g, '<br>');
  function draft(t, preferReply) { const it = Office.context.mailbox.item; const fromSender = it && P.mail && t.to.some(x => String(x).toLowerCase() === String(P.mail.from).toLowerCase());
    try { if (preferReply !== false && fromSender && it.displayReplyFormAsync) { it.displayReplyFormAsync({ htmlBody: html(t.body) }, () => {}); return 'reply'; }
      const p = { toRecipients: t.to, subject: t.subject, htmlBody: html(t.body) }; const mb = Office.context.mailbox;
      if (mb.displayNewMessageFormAsync) mb.displayNewMessageFormAsync(p, () => {}); else mb.displayNewMessageForm(p); return 'new'; }
    catch (e) { copy(t.subject + '\n\n' + t.body); toast('Outlook couldn\'t open a draft. The text is copied.'); return 'copy'; } }
  function copy(txt) { (navigator.clipboard ? navigator.clipboard.writeText(txt) : Promise.reject()).catch(() => { const ta = document.createElement('textarea'); ta.value = txt; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove(); }); }

  /* ---------- taps ---------- */
  document.addEventListener('click', e => { const b = e.target.closest('button'); if (!b || b.disabled) return; const d = b.dataset;
    if (b.id === 'gear') return P.view === 'settings' ? back() : settings();
    if (d.tab) return showTab(d.tab);
    if (d.view) return view(d.view);
    if (d.act === 'vclose') return viewClose();
    if (d.pick) return pickLoad(d.pick);
    if (d.open) return openLoad(d.open);
    if (d.alsoto) return saveAlso(d.alsoto);
    if (d.res) return resolve(d.doc, d.res);
    if (d.do) return act(d.do);
    if (d.tpl) { const t = (P.res.drafts || []).find(x => x.id === d.tpl); if (t) { if (!t.to.length && t.party === 'customer') return toast('No customer email in the Hub for this load. Add it under Customers.'); draft(t); } return; }
    if (d.ltpl) { const t = ((P.loadView || {}).drafts || []).find(x => x.id === d.ltpl); if (t) draft(t, false); return; }
    if (d.ask) return ask(d.load, d.ask);
    if (d.offer) { const c = P.res.alert.carriers[+d.offer]; if (!c.email) { copy(c.offer.subject + '\n\n' + c.offer.body); return toast('No email for ' + c.name + ' in the Hub. The offer text is copied.'); } draft({ to: [c.email], subject: c.offer.subject, body: c.offer.body }, false); P.doneTexts['offer' + d.offer] = 1; return render(); }
    const a = d.act; if (a === 'save') return save(false); if (a === 'sig') { P.showSig = !P.showSig; return render(); }
    if (a === 'find') { P.find.mode = ''; return showTab('find'); } if (a === 'findgo') { const q = ($('#q') || $('#qi') || {}).value || ''; return doFind(q.trim()); }
    if (a === 'also') { P.find.mode = 'also'; P.find.res = null; P.find.q = ''; return showTab('find'); } if (a === 'alsooff') { P.find.mode = ''; return showTab('email'); }
    if (a === 'back') return showTab(P.today ? 'today' : 'email'); });
  document.addEventListener('change', e => { const t = e.target; if (t.dataset && t.dataset.tick) { const f = P.files[t.dataset.tick]; if (f) { f.tick = t.checked; f.mine = true; } render(); }
    if (t.dataset && t.dataset.type) { const f = P.files[t.dataset.type]; if (f) { f.type = t.value; f.mine = true; } render(); } });
  document.addEventListener('toggle', e => { const t = e.target; if (t.dataset && t.dataset.details) { P.details = t.open; ls.set('fh.details', t.open ? '1' : '0'); } if (t.dataset && t.dataset.grp && t.open) P.todayOpen = t.dataset.grp; }, true);
  document.addEventListener('keydown', e => { if (e.key === 'Enter' && e.target.id === 'qi') doFind(e.target.value.trim()); });

  /* ---------- writes ---------- */
  async function save(auto, alsoTo) { if (P.busy) return; const r = P.res, m = P.mail; const lid = alsoTo || (r && r.loadId); if (!r || !lid) return; const sw = P.switches || {}; const me = sw.me || {};
    const files = alsoTo ? (r.files || []).filter(f => (P.files[f.id] || {}).saved && !f.signature) : (r.files || []).filter(f => { const st = P.files[f.id] || {}; if (st.saved || f.state === 'dup' || f.state === 'noload') return false; if (auto) return me.autoFiles && f.state === 'new' && !f.signature; return st.tick; });
    const steps = alsoTo ? [] : paperSteps(files);
    P.busy = true; render(); const my = P.seq; forget();
    try { if (P.ready) await P.ready; const copyE = r.already && !alsoTo ? null : await emailCopy(); const base = { internetMessageId: m.internetMessageId, conversationId: m.conversationId, itemId: m.itemId, loadId: lid, subject: m.subject, from: m.from, to: m.to, cc: m.cc, received: m.received, body: m.body, party: r.party ? { type: r.party.type, name: r.party.name } : null, auto: !!auto, also: !!alsoTo };
      const part = f => { const a = m.atts.find(x => x.id === f.id) || {}; const st = P.files[f.id] || {}; return { id: f.id, name: f.name, type: st.type || f.type, contentType: a.contentType, base64: a.base64, sha256: a.sha256, signature: f.signature, ticked: true }; };
      const total = files.reduce((n, f) => n + (+f.size || 0), 0); const calls = total > MAX_ONE_CALL ? [[]].concat(files.map(f => [f])) : [files];
      const tokens = []; let out = { saved: [], skipped: [], conflicts: [], refused: [] };
      for (let i = 0; i < calls.length; i++) { const res = await api('save', Object.assign({}, base, { email: i === 0 ? copyE : null, files: calls[i].map(part) }), 60000); if (res.undoToken) tokens.push(res.undoToken); ['saved', 'skipped', 'conflicts', 'refused'].forEach(k => { out[k] = out[k].concat(res[k] || []); }); }
      if (my !== P.seq) return;
      const did = [];
      if (alsoTo) { const ids = (r.already && r.already.loadIds) || (r.already ? [r.already.loadId] : []); r.already = Object.assign(r.already || { at: localISO(), by: P.name || 'you' }, { loadIds: ids.concat([alsoTo]) }); P.find.mode = ''; P.tab = 'email'; tabs(); }
      else { out.saved.forEach(s => { P.files[s.id] = Object.assign(P.files[s.id] || {}, { saved: 'Saved as ' + s.as, tick: false }); }); out.skipped.forEach(s => { P.files[s.id] = Object.assign(P.files[s.id] || {}, { saved: s.text, tick: false }); });
        P.conflicts = out.conflicts; P.res.already = P.res.already || { loadId: lid, at: localISO(), by: P.name || 'you' };
        for (const s of steps) { if (!out.saved.length && !out.skipped.length) break; try { const x = await api('action', { kind: s.kind, loadId: lid, internetMessageId: m.internetMessageId, from: m.from, body: m.body }, 30000); P.doneTexts[s.kind] = x.text; if (x.undoToken) tokens.push(x.undoToken); did.push(s.short); } catch (e) { /* the save stands; the step can be tapped */ } } }
      toast((auto ? 'Saved automatically: ' : 'Saved: ') + (out.saved.length ? plural(out.saved.length, 'file') + ' + ' : '') + 'email to ' + lid + (did.length ? ' · ' + did.join(' · ') : '') + (out.skipped.length ? ' · ' + out.skipped.length + ' already there' : '') + (out.refused.length ? ' · ' + out.refused.map(x => x.name + ' (' + x.why + ')').join(', ') : ''), tokens.length ? () => undo(tokens) : null); remember();
    } catch (e) { toast(e.message, null, true); } finally { P.busy = false; render(); } }
  function saveAlso(id) { if (P.res && id === P.res.loadId) return toast('This email is already on ' + id + '.'); return save(false, id); }
  async function resolve(docId, choice) { forget(); try { const r = await api('resolve', { docId, choice }, 60000); P.conflicts = P.conflicts.filter(c => c.docId !== docId); render(); toast(r.status === 'saved' ? 'Saved as ' + r.name : 'Skipped', r.undoToken ? () => undo([r.undoToken]) : null); } catch (e) { toast(e.message); } }
  async function act(kind) { const r = P.res; if (!r || !r.loadId) return; const m = P.mail; let action = 'action', data = { kind, loadId: r.loadId, internetMessageId: m.internetMessageId, from: m.from, body: m.body };
    if (kind === 'cc') data.note = ($('#ccnote') || {}).value || '';
    if (kind === 'driver') { action = 'driver'; data = { loadId: r.loadId, fields: r.driver.fields, link: r.driver.link }; }
    if (kind === 'update') { action = 'update'; P.upd = { status: $('#ust').value, eta: $('#ueta').value.trim(), note: $('#unote').value.trim() }; data = Object.assign({ loadId: r.loadId, from: m.from, internetMessageId: m.internetMessageId }, P.upd); }
    if (kind === 'det' || kind === 'lumper') { action = 'issue'; const i = r.issueHint; data = { kind, loadId: r.loadId, subject: m.subject, arrived: i.arrived, stop: i.stop, internetMessageId: m.internetMessageId }; }
    const btn = document.querySelector(`[data-do="${kind}"]`); if (btn) { btn.disabled = true; btn.textContent = 'Working…'; }
    forget(); try { const res = await api(action, data, 30000); P.doneTexts[kind] = res.text; const L = P.res.load;
      if (kind === 'driver' && res.fields) { const f = res.fields; if (f.Driver || f['Driver Phone']) L.driver = [f.Driver || L.driver.split(' · ')[0], f['Driver Phone']].filter(Boolean).join(' · '); if (f.Truck) L.truck = f.Truck; if (f.Trailer) L.trailer = f.Trailer; if (L.next === 'Needs driver info') L.next = ''; }
      if (kind === 'update' && P.upd.eta) L.eta = P.upd.eta; if (L && L.missing) L.missing = L.missing.filter(x => !(kind === 'pod' && x === 'POD') && !(kind === 'inv' && x === 'Invoice'));
      render(); remember(); toast(res.already ? res.text : res.text, res.undoToken ? () => undo([res.undoToken], kind) : null); } catch (e) { toast(e.message); render(); } }
  async function ask(loadId, tplId) { try { const r = await api('load', { loadId }, 15000); const t = (r.drafts || []).find(x => x.id === tplId); if (!t) return toast('No template for that.'); if (!t.to.length) { copy(t.subject + '\n\n' + t.body); return toast('No carrier email on ' + loadId + '. The text is copied.'); } draft(t, false); } catch (e) { toast(e.message); } }
  async function undo(tokens, kind) { forget(); try { for (const t of tokens.slice().reverse()) await api('undo', { undoToken: t }, 30000); if (kind) delete P.doneTexts[kind]; toast('Undone'); P.tab = 'email'; tabs(); onItem(); } catch (e) { toast(e.message); } }

  let tt = 0; function toast(text, onUndo, long) { const el = $('#toast'); el.hidden = false; el.innerHTML = `<span>${esc(text)}</span>${onUndo ? '<button type="button">Undo</button>' : ''}`; if (onUndo) el.querySelector('button').onclick = () => { el.hidden = true; onUndo(); }; clearTimeout(tt); tt = setTimeout(() => { el.hidden = true; }, onUndo ? 12000 : long ? 9000 : 3500); }

  /* ---------- setup and settings ---------- */
  function setup(msg) { P.view = 'setup'; $('#gear').hidden = true; tabs();
    $('#app').innerHTML = `<section class="card"><b>Connect to the Freight Hub</b>${msg ? `<div class="err" style="margin-top:6px">${esc(msg)}</div>` : '<div class="line muted">Both come from the Hub: More › Outlook › Make my key.</div>'}
      <label class="f" for="u">Outlook API address</label><input class="inp" id="u" type="url" autocomplete="off" spellcheck="false" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(P.url)}">
      <label class="f" for="k">Your personal key</label><input class="inp" id="k" type="password" autocomplete="off" spellcheck="false" placeholder="fh_…">
      <button class="btn pri big" type="button" id="connect" style="margin-top:12px">Connect</button><div id="cmsg" class="line"></div></section>
      <section class="card"><b>Pin the panel</b><div class="line muted">Click the pin at the top of this panel so it stays open and reads every email you click.</div></section>`;
    $('#connect').onclick = async () => { const u = $('#u').value.trim(), k = $('#k').value.trim(); if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(u)) { $('#cmsg').innerHTML = '<span class="err">That should start with https://script.google.com/</span>'; return; }
      P.url = u; P.key = k; $('#cmsg').innerHTML = '<span class="spin" aria-hidden="true"></span>Connecting…';
      try { const r = await api('ping', {}, 20000); P.name = r.name; P.switches = r.switches; await saveSettings({ hubUrl: u, hubKey: k, hubName: r.name }); $('#cmsg').innerHTML = `<span class="pass">Connected as ${esc(r.name)}</span> · <button class="lk" type="button" id="runc">Run checks</button>`; $('#runc').onclick = () => { P.view = 'settings'; settings(true); }; setTimeout(() => { if (P.view === 'setup') back(); }, 2500); }
      catch (e) { P.key = ''; $('#cmsg').innerHTML = `<span class="err">${esc(e.message)}</span>`; } }; }
  function back() { P.view = 'main'; $('#gear').hidden = false; P.tab = 'email'; tabs(); onItem(); }
  async function settings(runNow) { P.view = 'settings'; tabs(); const me = (P.switches && P.switches.me) || { auto: false, autoFiles: false, hideSig: true }; const n = Object.keys(memRead()).length;
    $('#app').innerHTML = `<section class="card"><b>Settings</b><div class="line muted">Connected as ${esc(P.name || '…')}. These are yours only.</div>
      <label class="sw"><input type="checkbox" id="pdet" ${P.details ? 'checked' : ''}> Show load details open (pickup, delivery, dispatch…)</label>
      <label class="sw"><input type="checkbox" data-pref="auto" ${me.auto ? 'checked' : ''}> Save automatically when the match is certain</label>
      <label class="sw"><input type="checkbox" data-pref="autoFiles" ${me.autoFiles ? 'checked' : ''}> Include attachments in automatic saves</label>
      <label class="sw"><input type="checkbox" data-pref="hideSig" ${me.hideSig !== false ? 'checked' : ''}> Hide signature images</label>
      <div class="row" style="margin-top:6px"><span class="line muted">Remembers ${plural(n, 'email')} on this computer for 12 hours (no files, no pay).</span><button class="lk" type="button" id="clrmem">Clear</button></div></section>
      <section class="card"><div class="row"><b>Run checks</b><button class="btn" type="button" id="run" style="margin-left:auto">Run</button></div><div id="ck">${P.checks ? checksHTML(P.checks) : '<div class="line muted">Checks Outlook and the Hub connection.</div>'}</div></section>
      <div class="row"><button class="btn" type="button" id="done">Back</button><button class="btn" type="button" id="disc">Disconnect</button></div>`;
    $('#done').onclick = back; $('#run').onclick = runChecks; $('#disc').onclick = async () => { ls.del(MEM_KEY); MEMO.clear(); [IDX_KEY, THR_KEY, SIG_KEY].forEach(k => ls.del(k)); P.idx = null; await saveSettings({ hubKey: null, hubName: null }); P.key = ''; setup(); };
    $('#clrmem').onclick = () => { ls.del(MEM_KEY); MEMO.clear(); [IDX_KEY, THR_KEY, SIG_KEY].forEach(k => ls.del(k)); P.idx = null; toast('Memory cleared'); settings(); };
    $('#pdet').onchange = e => { P.details = e.target.checked; ls.set('fh.details', P.details ? '1' : '0'); };
    document.querySelectorAll('[data-pref]').forEach(i => { i.onchange = async () => { const prefs = {}; prefs[i.dataset.pref] = i.checked; try { const r = await api('prefs', { prefs }, 20000); P.switches = Object.assign({}, P.switches, { me: r.me }); toast('Saved'); } catch (e) { i.checked = !i.checked; toast(e.message); } }; });
    if (runNow) runChecks(); }
  function checksHTML(list) { return `<ul class="checks">${list.map(c => `<li><span class="${c.ok ? 'pass' : 'fail'}">${c.ok ? 'PASS' : 'FAIL'}</span><span>${esc(c.t)}</span></li>`).join('')}</ul><button class="btn" type="button" id="copyck" style="margin-top:8px">Copy results</button>`; }
  async function runChecks() { const out = []; const add = (ok, t) => out.push({ ok: !!ok, t });
    const ver = ['1.14', '1.13', '1.12', '1.11', '1.10', '1.9', '1.8', '1.7', '1.6', '1.5', '1.4', '1.3', '1.2', '1.1'].find(sup) || 'none';
    add(typeof Office !== 'undefined' && Office.context, 'Office.js loaded · Mailbox ' + ver + ' · panel 4.16.1');
    add(sup('1.5'), sup('1.5') ? 'Pinning supported' : 'Pinning needs Mailbox 1.5');
    try { const t0 = Date.now(); const r = await api('ping', {}, 20000); add(true, 'Hub reachable, key accepted (' + r.name + ', Hub ' + r.hub + ') · ' + (Date.now() - t0) + ' ms'); } catch (e) { add(false, 'Hub: ' + e.message); }
    const it = Office.context.mailbox.item; const att = it && (it.attachments || []).find(a => !a.isInline);
    if (!sup('1.8')) add(false, 'This Outlook can\'t hand files to add-ins. Save the email only.'); else if (!att) add(true, 'Attachment content readable (open an email with an attachment to test it)'); else { try { const c = await cb(it.getAttachmentContentAsync.bind(it), att.id); add(!!c.content, 'Attachment content readable (' + att.name + ', ' + c.format + ')'); } catch (e) { add(false, 'Attachment content: ' + (e.message || e)); } }
    add(sup('1.14') && !!(it && it.getAsFileAsync), sup('1.14') && it && it.getAsFileAsync ? 'getAsFileAsync available: emails save as .eml' : 'No getAsFileAsync: emails save as an HTML copy with headers');
    add(!!(it && it.displayReplyFormAsync), it && it.displayReplyFormAsync ? 'Reply drafts available' : 'Reply drafts missing');
    add(!!(it && it.notificationMessages), it && it.notificationMessages ? 'Status line on the email available' : 'Status line on the email not available');
    let memOk = false; try { localStorage.setItem('fh.t', '1'); memOk = localStorage.getItem('fh.t') === '1'; localStorage.removeItem('fh.t'); } catch (e) { memOk = false; } add(memOk, memOk ? 'Memory of recent emails works' : 'Memory of recent emails blocked (this session only)');
    try { rs().set('fhCheck', String(Date.now())); await new Promise((ok, bad) => rs().saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? ok() : bad(r.error))); add(!!rs().get('fhCheck'), 'roamingSettings saved and read back'); rs().remove('fhCheck'); rs().saveAsync(() => {}); } catch (e) { add(false, 'roamingSettings: ' + (e.message || e)); }
    P.checks = out; const ck = $('#ck'); if (ck) { ck.innerHTML = checksHTML(out); $('#copyck').onclick = () => { copy('Freight Hub for Outlook · Run checks\n' + out.map(c => (c.ok ? 'PASS ' : 'FAIL ') + c.t).join('\n')); toast('Copied'); }; }
    return out; }

  /* ---------- start ---------- */
  function theme() { try { const t = Office.context.officeTheme; const bg = t && t.bodyBackgroundColor; if (bg) { const n = parseInt(bg.replace('#', ''), 16); const l = (0.299 * (n >> 16 & 255) + 0.587 * (n >> 8 & 255) + 0.114 * (n & 255)); document.body.classList.toggle('dark', l < 110); } } catch (e) { /* light */ } }
  Office.onReady(() => { theme(); loadSettings(); $('#pinned').hidden = !sup('1.5');
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { if (P.view === 'main') onItem(); }); } catch (e) { /* not pinnable */ }
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, theme); } catch (e) { /* older Outlook */ }
    if (!P.url || !P.key) return setup(); $('#gear').hidden = false; P.view = 'main'; tabs(); onItem();
    idxLoad(); tabs(); setTimeout(idxRefresh, 1200); setInterval(() => { if (P.view === 'main') idxRefresh(); }, 120000); });   // v4.16.1: the load index + Today in one call, behind the first email; again every 2 min
  window.__fh = { P, runChecks, render, memRead, thrRead, sigStore };   // for the tests
})();
