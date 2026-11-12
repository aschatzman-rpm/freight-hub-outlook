/* Freight Hub for Outlook (v4.15) · the panel.
   Holds no data, no keys and no Hub address: those are typed in once and kept in Outlook roamingSettings (your mailbox).
   Reads only the email you have open. Nothing is saved, ticked or opened without a tap (unless you turn on auto-save). */
(function () {
  'use strict';
  const CLIENT = 'outlook-1.0';
  const TYPES = ['POD', 'Rate Con (Signed)', 'Invoice', 'BOL', 'Lumper', 'Scale', 'Other'];
  const MAX_ONE_CALL = 25 * 1024 * 1024;
  const $ = s => document.querySelector(s);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const rid = () => { const a = new Uint8Array(12); crypto.getRandomValues(a); return 'r' + Array.from(a, b => b.toString(16).padStart(2, '0')).join(''); };
  const P = { url: '', key: '', name: '', switches: null, view: 'main', seq: 0, mail: null, res: null, files: {}, showSig: false, conflicts: [], doneTexts: {}, busy: false, checks: null };

  /* ---------- settings in the mailbox (never localStorage, never a URL) ---------- */
  const rs = () => Office.context.roamingSettings;
  function loadSettings() { P.url = rs().get('hubUrl') || ''; P.key = rs().get('hubKey') || ''; P.name = rs().get('hubName') || ''; }
  function saveSettings(o) { return new Promise((ok, bad) => { Object.keys(o).forEach(k => { if (o[k] == null) rs().remove(k); else rs().set(k, o[k]); }); rs().saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? ok() : bad(new Error('Outlook couldn\'t save the settings'))); }); }

  /* ---------- talking to the Hub ---------- */
  async function api(action, data, ms, requestId) {
    if (!P.url || !P.key) throw Object.assign(new Error('Not connected'), { auth: true });
    const ctl = new AbortController(); const t = setTimeout(() => ctl.abort(), ms || 20000);
    try {
      const r = await fetch(P.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ key: P.key, action, requestId: requestId || rid(), client: CLIENT }, data || {})), redirect: 'follow', signal: ctl.signal });
      const j = await r.json(); if (j && j.ok === false) { const e = new Error(j.message || 'The Hub said no.'); e.auth = !!j.auth; throw e; } return j;
    } catch (e) { if (e.name === 'AbortError') throw new Error('Can\'t reach the Hub (it took too long).'); if (e instanceof TypeError) throw new Error('Can\'t reach the Hub.'); throw e; }
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
      const f = { id: a.id, name: a.name, size: a.size, contentType: a.contentType, isInline: !!a.isInline, sha256: '', base64: '' };
      if (sup('1.8') && a.size <= MAX_ONE_CALL) { try { const c = await cb(it.getAttachmentContentAsync.bind(it), a.id); if (c.format === 'base64' || c.format === Office.MailboxEnums.AttachmentContentFormat.Base64) { f.base64 = c.content; f.sha256 = await sha256(b64bytes(c.content)); } } catch (e) { /* unreadable: can't be saved */ } }
      m.atts.push(f); }
    return m;
  }
  async function emailCopy() { const it = Office.context.mailbox.item;
    if (sup('1.14') && it.getAsFileAsync) { try { return { eml: await cb(it.getAsFileAsync.bind(it)) }; } catch (e) { /* fall back */ } }
    let headers = ''; if (sup('1.8') && it.getAllInternetHeadersAsync) { try { headers = await cb(it.getAllInternetHeadersAsync.bind(it)); } catch (e) { headers = ''; } }
    let html = ''; try { html = await cb(it.body.getAsync.bind(it.body), Office.CoercionType.Html); } catch (e) { html = ''; }
    return { html, headers }; }

  /* ---------- each email ---------- */
  async function onItem() {
    const my = ++P.seq; P.res = null; P.files = {}; P.conflicts = []; P.doneTexts = {}; P.showSig = false; P.mail = null; if (P.view !== 'main') return;
    $('#app').innerHTML = '<div class="card muted"><span class="spin" aria-hidden="true"></span>Reading this email…</div>';
    let m; try { m = await readItem(); } catch (e) { m = null; }
    if (my !== P.seq) return; if (!m) { $('#app').innerHTML = '<div class="card muted">Open an email to see its load.</div>'; return; }
    P.mail = m;
    try { const r = await api('match', match(m), 20000); if (my !== P.seq) return; P.res = r; P.switches = r.switches || P.switches;
      (r.files || []).forEach(f => { P.files[f.id] = { type: f.type, tick: f.state === 'new' && !f.signature }; });
      render();
      if (P.switches && P.switches.me && P.switches.me.auto && r.match && r.match.certain && !r.already && r.loadId) save(true);
    } catch (e) { if (my !== P.seq) return; if (e.auth) return setup(e.message); $('#app').innerHTML = `<div class="card"><div class="err">${esc(e.message)}</div><button class="btn" id="retry" type="button" style="margin-top:10px">Retry</button></div>`; $('#retry').onclick = onItem; }
  }
  function match(m, loadId) { return { internetMessageId: m.internetMessageId, conversationId: m.conversationId, subject: m.subject, from: m.from, body: m.body, loadId: loadId || '', attachments: m.atts.map(a => ({ id: a.id, name: a.name, size: a.size, contentType: a.contentType, isInline: a.isInline, sha256: a.sha256 })) }; }
  async function pickLoad(id) { const my = ++P.seq; try { const r = await api('match', match(P.mail, id), 20000); if (my !== P.seq) return; P.res = r; P.files = {}; (r.files || []).forEach(f => { P.files[f.id] = { type: f.type, tick: f.state === 'new' && !f.signature }; }); render(); } catch (e) { toast(e.message); } }

  /* ---------- render (the mockup's order) ---------- */
  function render() { const r = P.res, m = P.mail; if (!r || !m) return; const sw = P.switches || {}; const me = sw.me || {}; const h = [];
    if (me.auto) h.push('<div class="note">Auto-save is on: certain matches save when you open them.</div>');
    if (r.load) { const L = r.load; const ok = r.match.certain || r.match.picked;
      h.push(`<section class="card" aria-label="Matched load"><div class="row"><span class="lab">Matched load</span><span class="chip ${ok ? 'ok' : 'warn'}">${esc(r.match.reason)}</span>${!r.match.certain ? '<button class="lk" type="button" data-act="change" style="margin-left:auto">Change</button>' : ''}</div>
        <div class="lid">${esc(L.id)}</div><div class="lane">${esc(L.lane)}</div><div class="status">${esc(L.status)}</div>
        <div class="kv">${[['Pickup', L.pickup], ['Delivery', L.delivery], ['Carrier', L.carrier], ['Dispatch', L.dispatch], ['Driver', L.driver]].filter(x => x[1]).map(x => `<span>${x[0]}</span><span>${esc(x[1])}</span>`).join('')}</div>
        ${L.missing && L.missing.length ? `<div class="missing">Missing: ${esc(L.missing.join(', '))}</div>` : ''}${L.issue ? `<div class="line">Open issue: ${esc(L.issue)}</div>` : ''}</section>`); }
    else if ((r.candidates || []).length) h.push(`<section class="card pick" aria-label="Pick the load"><b>Which load is this for?</b><div class="line muted">${esc(r.match.reason)}</div><div class="stack" style="margin-top:10px">${r.candidates.map(c => `<button class="btn act" type="button" data-pick="${esc(c.id)}"><b>${esc(c.id)}</b><div class="line muted">${esc(c.lane)} · ${esc(c.status)}</div></button>`).join('')}</div></section>`);
    else h.push(`<section class="card"><b>No load matched</b><div class="line muted">${r.party ? esc(r.match.reason) : 'The sender isn\'t a carrier or customer in the Hub, and there\'s no load # in the email. Nothing to do.'}</div></section>`);
    if (r.party) { const p = r.party; const lines = p.type === 'carrier' ? [p.mc && (/^(mc|dot)\b/i.test(p.mc) ? p.mc : 'MC / DOT ' + p.mc), p.dispatch && 'Dispatch: ' + p.dispatch, p.safety && 'Safety: ' + p.safety, p.active + ' active · ' + p.total + ' loads with RPM'] : [p.contact && 'Contact: ' + p.contact, p.active + ' active · ' + p.total + ' loads with RPM'];
      h.push(`<section class="card" aria-label="${p.type === 'carrier' ? 'Carrier' : 'Customer'}"><span class="lab">${p.type === 'carrier' ? 'Carrier' : 'Customer'}</span><div style="font-weight:700;margin-top:4px">${esc(p.name)}</div>${lines.filter(Boolean).map(l => `<div class="line">${esc(l)}</div>`).join('')}</section>`); }
    if (r.already) h.push(`<div class="saved" role="status"><span aria-hidden="true">✓</span><span>Already saved to Load ${esc(r.already.loadId)} · ${esc(when(r.already.at))} · ${esc(r.already.by)}. Opening this email again never saves it twice.</span></div>`);
    const files = r.files || []; const hide = me.hideSig !== false && !P.showSig; const shown = files.filter(f => !(hide && f.signature)); const hidden = files.length - shown.length;
    if (files.length) h.push(`<section class="card" aria-label="Attachments"><b>Attachments</b>${shown.map((f, i) => { const st = P.files[f.id] || {}; const dis = f.state === 'dup' || f.state === 'noload' || !!st.saved; const txt = st.saved ? st.saved : f.state === 'new' && st.type !== f.type ? 'New · will be saved as ' + st.type : f.text;
        return `<div class="att"><input type="checkbox" id="att${i}" data-tick="${esc(f.id)}" ${st.tick ? 'checked' : ''} ${dis ? 'disabled' : ''}><div style="flex:1;min-width:0"><label class="nm" for="att${i}">${esc(f.name)}</label><div class="st ${esc(st.saved ? 'new' : f.state)}">${esc(txt)}</div></div><button class="type" type="button" data-cycle="${esc(f.id)}" aria-label="File type: ${esc(st.type || f.type)}. Tap to change" ${dis ? 'disabled' : ''}>${esc(st.type || f.type)}</button></div>`; }).join('')}
      ${hidden || P.showSig ? `<button class="lk" type="button" data-act="sig">${P.showSig ? 'Hide signature images' : plural(hidden, 'signature image') + ' hidden · show'}</button>` : ''}</section>`);
    P.conflicts.forEach(c => h.push(`<section class="conflict" aria-label="File conflict"><b>${esc(c.name)}</b><p>${esc(c.text)}</p><div class="row"><button class="btn pri" type="button" data-res="new" data-doc="${esc(c.docId)}">Save as ${esc(c.next)}</button><button class="btn" type="button" data-res="replace" data-doc="${esc(c.docId)}">Replace</button><button class="btn" type="button" data-res="skip" data-doc="${esc(c.docId)}">Skip</button></div></section>`));
    const ticked = shown.filter(f => (P.files[f.id] || {}).tick && !(P.files[f.id] || {}).saved).length;
    if (r.load && (!r.already || ticked)) h.push(`<button class="btn pri big" type="button" data-act="save" ${P.busy ? 'disabled' : ''}>${P.busy ? 'Saving…' : r.already ? 'Save ' + plural(ticked, 'file') + ' to load' : ticked ? 'Save email + ' + plural(ticked, 'file') + ' to load' : 'Save email to load'}</button>`);
    if (r.load && (r.actions || []).length) h.push(`<section class="stack" aria-label="Actions">${r.actions.map(a => { const t = P.doneTexts[a.kind] || (a.done ? a.text : '');
        if (t) return `<div class="done"><span aria-hidden="true">✓</span><span>${esc(t)}</span></div>`;
        if (a.kind === 'cc') return `<div class="card"><label class="f" for="ccnote" style="margin-top:0">Check call note</label><textarea class="inp" id="ccnote">${esc(a.note || '')}</textarea><button class="btn act" type="button" data-do="cc" style="margin-top:8px">${esc(a.label)}</button></div>`;
        return `<button class="btn act" type="button" data-do="${esc(a.kind)}">${esc(a.label)}</button>`; }).join('')}</section>`);
    $('#app').innerHTML = h.join('');
  }
  const localISO = () => { const d = new Date(), z = n => String(n).padStart(2, '0'); return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate()) + 'T' + z(d.getHours()) + ':' + z(d.getMinutes()); };
  const plural = (n, w) => n + ' ' + w + (n === 1 ? '' : 's');
  function when(s) { const m = /^(\d{4})-(\d{2})-(\d{2})T?(\d{2})?:?(\d{2})?/.exec(String(s || '')); if (!m) return String(s || ''); const h = +(m[4] || 0); return (+m[2]) + '/' + (+m[3]) + (m[4] ? ' ' + ((h % 12) || 12) + ':' + m[5] + ' ' + (h < 12 ? 'AM' : 'PM') : ''); }

  /* ---------- taps ---------- */
  document.addEventListener('click', e => { const b = e.target.closest('button'); if (!b || b.disabled) return;
    if (b.id === 'gear') return P.view === 'settings' ? back() : settings();
    if (b.dataset.pick) return pickLoad(b.dataset.pick);
    if (b.dataset.cycle) { const f = P.files[b.dataset.cycle]; if (!f) return; f.type = TYPES[(TYPES.indexOf(f.type) + 1) % TYPES.length]; return render(); }
    if (b.dataset.res) return resolve(b.dataset.doc, b.dataset.res);
    if (b.dataset.do) return act(b.dataset.do);
    const a = b.dataset.act; if (a === 'save') return save(false); if (a === 'sig') { P.showSig = !P.showSig; return render(); }
    if (a === 'change') { P.res.candidates = P.res.candidates && P.res.candidates.length ? P.res.candidates : []; P.res.load = null; if (!P.res.candidates.length) P.res.match.reason = 'Type the load # in the email, or pick from the Hub.'; return render(); } });
  document.addEventListener('change', e => { const t = e.target; if (t.dataset && t.dataset.tick) { const f = P.files[t.dataset.tick]; if (f) f.tick = t.checked; render(); } });

  async function save(auto) { if (P.busy) return; const r = P.res, m = P.mail; if (!r || !r.loadId) return; const sw = P.switches || {}; const me = sw.me || {};
    const files = (r.files || []).filter(f => { const st = P.files[f.id] || {}; if (st.saved || f.state === 'dup' || f.state === 'noload') return false; if (auto) return me.autoFiles && f.state === 'new' && !f.signature; return st.tick; });
    P.busy = true; render(); const my = P.seq;
    try { const copy = r.already ? null : await emailCopy(); const base = { internetMessageId: m.internetMessageId, conversationId: m.conversationId, itemId: m.itemId, loadId: r.loadId, subject: m.subject, from: m.from, to: m.to, cc: m.cc, received: m.received, body: m.body, party: r.party ? { type: r.party.type, name: r.party.name } : null, auto: !!auto };
      const part = f => { const a = m.atts.find(x => x.id === f.id) || {}; const st = P.files[f.id] || {}; return { id: f.id, name: f.name, type: st.type || f.type, contentType: a.contentType, base64: a.base64, sha256: a.sha256, signature: f.signature, ticked: !!st.tick }; };
      const total = files.reduce((n, f) => n + (+f.size || 0), 0); const calls = total > MAX_ONE_CALL ? [[]].concat(files.map(f => [f])) : [files];
      const tokens = []; let out = { saved: [], skipped: [], conflicts: [], refused: [] };
      for (let i = 0; i < calls.length; i++) { const res = await api('save', Object.assign({}, base, { email: i === 0 ? copy : null, files: calls[i].map(part) }), 60000); if (res.undoToken) tokens.push(res.undoToken); ['saved', 'skipped', 'conflicts', 'refused'].forEach(k => { out[k] = out[k].concat(res[k] || []); }); }
      if (my !== P.seq) return;
      out.saved.forEach(s => { P.files[s.id] = Object.assign(P.files[s.id] || {}, { saved: 'Saved as ' + s.as, tick: false }); });
      out.skipped.forEach(s => { P.files[s.id] = Object.assign(P.files[s.id] || {}, { saved: s.text, tick: false }); });
      P.conflicts = out.conflicts; P.res.already = P.res.already || { loadId: r.loadId, at: localISO(), by: P.name || 'you' };
      toast((auto ? 'Saved automatically: ' : 'Saved: ') + (out.saved.length ? plural(out.saved.length, 'file') + ' + ' : '') + 'email to ' + r.loadId + (out.skipped.length ? ' · ' + out.skipped.length + ' already there' : '') + (out.refused.length ? ' · ' + out.refused.map(x => x.name + ' (' + x.why + ')').join(', ') : ''), tokens.length ? () => undo(tokens) : null);
    } catch (e) { toast(e.message, null, true); } finally { P.busy = false; render(); } }
  async function resolve(docId, choice) { try { const r = await api('resolve', { docId, choice }, 60000); P.conflicts = P.conflicts.filter(c => c.docId !== docId); render(); toast(r.status === 'saved' ? 'Saved as ' + r.name : 'Skipped', r.undoToken ? () => undo([r.undoToken]) : null); } catch (e) { toast(e.message); } }
  async function act(kind) { const r = P.res; if (!r || !r.loadId) return; const note = kind === 'cc' ? (($('#ccnote') || {}).value || '') : '';
    try { const res = await api('action', { kind, loadId: r.loadId, internetMessageId: P.mail.internetMessageId, note, from: P.mail.from, body: P.mail.body }, 30000); P.doneTexts[kind] = res.text; const L = P.res.load; if (L && L.missing) L.missing = L.missing.filter(x => !(kind === 'pod' && x === 'POD') && !(kind === 'inv' && x === 'Invoice')); render(); toast(res.already ? res.text + ' (already done)' : res.text, res.undoToken ? () => undo([res.undoToken], kind) : null); } catch (e) { toast(e.message); } }
  async function undo(tokens, kind) { try { for (const t of tokens.slice().reverse()) await api('undo', { undoToken: t }, 30000); if (kind) delete P.doneTexts[kind]; toast('Undone'); onItem(); } catch (e) { toast(e.message); } }

  let tt = 0; function toast(text, onUndo, long) { const el = $('#toast'); el.hidden = false; el.innerHTML = `<span>${esc(text)}</span>${onUndo ? '<button type="button">Undo</button>' : ''}`; if (onUndo) el.querySelector('button').onclick = () => { el.hidden = true; onUndo(); }; clearTimeout(tt); tt = setTimeout(() => { el.hidden = true; }, long || onUndo ? 9000 : 4000); }

  /* ---------- setup and ⚙ ---------- */
  function setup(msg) { P.view = 'setup'; $('#gear').hidden = true;
    $('#app').innerHTML = `<section class="card"><b>Connect to the Freight Hub</b>${msg ? `<div class="err" style="margin-top:6px">${esc(msg)}</div>` : '<div class="line muted">Both come from the Hub: More › Outlook › Make my key.</div>'}
      <label class="f" for="u">Outlook API address</label><input class="inp" id="u" type="url" autocomplete="off" spellcheck="false" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(P.url)}">
      <label class="f" for="k">Your personal key</label><input class="inp" id="k" type="password" autocomplete="off" spellcheck="false" placeholder="fh_…">
      <button class="btn pri big" type="button" id="connect" style="margin-top:12px">Connect</button><div id="cmsg" class="line"></div></section>
      <section class="card"><b>Pin the panel</b><div class="line muted">Click the pin at the top of this panel so it stays open and follows every email you open.</div></section>`;
    $('#connect').onclick = async () => { const u = $('#u').value.trim(), k = $('#k').value.trim(); if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(u)) { $('#cmsg').innerHTML = '<span class="err">That should start with https://script.google.com/</span>'; return; }
      P.url = u; P.key = k; $('#cmsg').innerHTML = '<span class="spin" aria-hidden="true"></span>Connecting…';
      try { const r = await api('ping', {}, 20000); P.name = r.name; P.switches = r.switches; await saveSettings({ hubUrl: u, hubKey: k, hubName: r.name }); $('#cmsg').innerHTML = `<span class="pass">Connected as ${esc(r.name)}</span> · <button class="lk" type="button" id="runc">Run checks</button>`; $('#runc').onclick = () => { P.view = 'settings'; settings(true); }; setTimeout(() => { if (P.view === 'setup') back(); }, 2500); }
      catch (e) { P.key = ''; $('#cmsg').innerHTML = `<span class="err">${esc(e.message)}</span>`; } }; }
  function back() { P.view = 'main'; $('#gear').hidden = false; onItem(); }
  async function settings(runNow) { P.view = 'settings'; const me = (P.switches && P.switches.me) || { auto: false, autoFiles: false, hideSig: true };
    $('#app').innerHTML = `<section class="card"><b>Settings</b><div class="line muted">Connected as ${esc(P.name || '…')}. These are yours only.</div>
      <label class="sw"><input type="checkbox" data-pref="auto" ${me.auto ? 'checked' : ''}> Save automatically when the match is certain</label>
      <label class="sw"><input type="checkbox" data-pref="autoFiles" ${me.autoFiles ? 'checked' : ''}> Include attachments in automatic saves</label>
      <label class="sw"><input type="checkbox" data-pref="hideSig" ${me.hideSig !== false ? 'checked' : ''}> Hide signature images</label></section>
      <section class="card"><div class="row"><b>Run checks</b><button class="btn" type="button" id="run" style="margin-left:auto">Run</button></div><div id="ck">${P.checks ? checksHTML(P.checks) : '<div class="line muted">Checks Outlook and the Hub connection.</div>'}</div></section>
      <div class="row"><button class="btn" type="button" id="done">Back</button><button class="btn" type="button" id="disc">Disconnect</button></div>`;
    $('#done').onclick = back; $('#run').onclick = runChecks; $('#disc').onclick = async () => { await saveSettings({ hubKey: null, hubName: null }); P.key = ''; setup(); };
    document.querySelectorAll('[data-pref]').forEach(i => { i.onchange = async () => { const prefs = {}; prefs[i.dataset.pref] = i.checked; try { const r = await api('prefs', { prefs }, 20000); P.switches = Object.assign({}, P.switches, { me: r.me }); toast('Saved'); } catch (e) { i.checked = !i.checked; toast(e.message); } }; });
    if (runNow) runChecks(); }
  function checksHTML(list) { return `<ul class="checks">${list.map(c => `<li><span class="${c.ok ? 'pass' : 'fail'}">${c.ok ? 'PASS' : 'FAIL'}</span><span>${esc(c.t)}</span></li>`).join('')}</ul><button class="btn" type="button" id="copyck" style="margin-top:8px">Copy results</button>`; }
  async function runChecks() { const out = []; const add = (ok, t) => out.push({ ok: !!ok, t });
    const ver = ['1.14', '1.13', '1.12', '1.11', '1.10', '1.9', '1.8', '1.7', '1.6', '1.5', '1.4', '1.3', '1.2', '1.1'].find(sup) || 'none';
    add(typeof Office !== 'undefined' && Office.context, 'Office.js loaded · ' + ((Office.context.diagnostics && (Office.context.diagnostics.hostName + ' ' + Office.context.diagnostics.version)) || 'Outlook') + ' · Mailbox ' + ver);
    add(sup('1.5'), sup('1.5') ? 'Pinning supported' : 'Pinning needs Mailbox 1.5');
    try { const t0 = Date.now(); const r = await api('ping', {}, 20000); add(true, 'Hub reachable, key accepted (' + r.name + ', Hub ' + r.hub + ') · ' + (Date.now() - t0) + ' ms'); } catch (e) { add(false, 'Hub: ' + e.message); }
    const it = Office.context.mailbox.item; const att = it && (it.attachments || []).find(a => !a.isInline);
    if (!sup('1.8')) add(false, 'This Outlook can\'t hand files to add-ins. Save the email only.'); else if (!att) add(true, 'Attachment content readable (open an email with an attachment to test it)'); else { try { const c = await cb(it.getAttachmentContentAsync.bind(it), att.id); add(!!c.content, 'Attachment content readable (' + att.name + ', ' + c.format + ')'); } catch (e) { add(false, 'Attachment content: ' + (e.message || e)); } }
    add(sup('1.14') && !!(it && it.getAsFileAsync), sup('1.14') && it && it.getAsFileAsync ? 'getAsFileAsync available: emails save as .eml' : 'No getAsFileAsync: emails save as an HTML copy with headers');
    add(!!(it && it.displayReplyFormAsync), it && it.displayReplyFormAsync ? 'displayReplyFormAsync available' : 'displayReplyFormAsync missing');
    try { rs().set('fhCheck', String(Date.now())); await new Promise((ok, bad) => rs().saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? ok() : bad(r.error))); add(!!rs().get('fhCheck'), 'roamingSettings saved and read back'); rs().remove('fhCheck'); rs().saveAsync(() => {}); } catch (e) { add(false, 'roamingSettings: ' + (e.message || e)); }
    P.checks = out; const ck = $('#ck'); if (ck) { ck.innerHTML = checksHTML(out); $('#copyck').onclick = () => { const txt = 'Freight Hub for Outlook · Run checks\n' + out.map(c => (c.ok ? 'PASS ' : 'FAIL ') + c.t).join('\n'); (navigator.clipboard ? navigator.clipboard.writeText(txt) : Promise.reject()).then(() => toast('Copied'), () => { const ta = document.createElement('textarea'); ta.value = txt; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove(); toast('Copied'); }); }; }
    return out; }

  /* ---------- start ---------- */
  function theme() { try { const t = Office.context.officeTheme; const bg = t && t.bodyBackgroundColor; if (bg) { const n = parseInt(bg.replace('#', ''), 16); const l = (0.299 * (n >> 16 & 255) + 0.587 * (n >> 8 & 255) + 0.114 * (n & 255)); document.body.classList.toggle('dark', l < 110); } } catch (e) { /* light */ } }
  Office.onReady(() => { theme(); loadSettings(); $('#pinned').hidden = !sup('1.5');
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { if (P.view === 'main') onItem(); }); } catch (e) { /* not pinnable */ }
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, theme); } catch (e) { /* older Outlook */ }
    if (!P.url || !P.key) return setup(); $('#gear').hidden = false; onItem(); });
  window.__fh = { P, runChecks, render };   // for the tests
})();
