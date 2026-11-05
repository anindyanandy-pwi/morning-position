/* AEN Morning Position — website logic. */
(function () {
  const C = window.AMP_CONFIG || {};
  const $ = id => document.getElementById(id);
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const store = { get: k => { try { return localStorage.getItem(k); } catch (e) { return null; } },
                  set: (k, v) => { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch (e) {} } };

  let session = store.get('amp_session'), me = null;

  /* ---------- server ---------- */
  const GIVE_UP_MS = 390000;   // 6 min 30 s: the website stops waiting for an answer (spec §19 item 3)
  async function call(fn, args) {
    let r; const t0 = Date.now();
    const ctrl = (typeof AbortController !== 'undefined') ? new AbortController() : null;
    const timer = ctrl ? setTimeout(function () { ctrl.abort(); }, GIVE_UP_MS) : null;
    try {
      r = await fetch(C.API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ session, fn, args: args || {} }), signal: ctrl ? ctrl.signal : undefined });
    } catch (e) {
      if (timer) clearTimeout(timer);
      // Tell a timeout apart from a real network failure (spec §19 item 6): Google kills a web-app call at
      // 6 minutes by dropping the connection, which looks like a network error — the time taken gives it away.
      if (ctrl && e && e.name === 'AbortError') throw new Error('The server took more than 6½ minutes and did not answer — Generate again.');
      if (Date.now() - t0 > 330000) throw new Error('The server took more than 6 minutes and was stopped — Generate again.');
      throw new Error('The server could not be reached. Check the internet connection and try again.');
    }
    if (timer) clearTimeout(timer);
    if (!r.ok) throw new Error('The server did not answer (' + r.status + ').');
    let j;
    // a 200 whose body is not JSON is Google's 6-minute kill page, not a reachability problem (spec §19 item 6)
    try { j = await r.json(); } catch (e) { throw new Error('The server took more than 6 minutes and was stopped — Generate again.'); }
    if (!j.ok) {
      if (/sign in again|Please sign in|no longer registered|expired\. Please/.test(j.error || '') && fn !== 'otpVerify') signedOut(j.error);
      throw new Error(j.error || 'Server error');
    }
    return j;
  }
  const say = (node, text, kind) => { node.textContent = text || ''; node.className = 'msg' + (kind ? ' ' + kind : ''); };
  const tell = (text, kind) => say($('msg'), text, kind);

  /* ---------- dates and names ---------- */
  const dmy = iso => iso ? iso.split('-').reverse().join('.') : '\u2014';
  const dotted = iso => iso.split('-').reverse().join('·');
  const addDays = (iso, n) => { const d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const weekday = iso => new Date(iso + 'T00:00:00Z').toLocaleDateString('en-IN', { weekday: 'long', timeZone: 'UTC' });
  const town = office => String(office || '').replace(/^Assistant Divisional Engineer[\s,/]*/i, '').replace(/\s*\(.*\)\s*$/, '').replace(/^Track\/(.*)$/, '$1 (Track)');
  const two = n => String(n).padStart(2, '0');
  const SECTION = { miniControl: 'AEN mini control', pwiMc: 'PWI mini control', gps: 'Winter patrolling', major: 'Major failures', minor: 'Minor failures', tdcLapse: 'TDC lapse',
    pqrsTdc: 'PQRS TDC fail', protection: 'Protection caution', badRoad: 'Bad road surface', cctv: 'CCTV', niLc: 'NI LC inspection', kpi: 'KPI and inspections',
    blocks: 'Blocks', machines: 'Machines', tms: 'TMS items', file: 'Saved file' };
  const gapText = g => (SECTION[g.section] || g.section) + ': ' + (g.kind === 'stale' ? 'older data used' : g.text);
  // "why a gap" (spec §13): reason codes in plain words
  const REASON = { 'not-entered': 'Not entered yet — the sheet is behind the date', 'date-form': 'Date written in an unusual form', 'tab-missing': 'Tab name not found',
    'columns-changed': 'A column heading has changed', 'not-placed': 'Rows not placed under any AEN', 'reader-not-run': 'The TMS reader has not uploaded it yet',
    'source-error': 'The source could not be read', 'column-repair': 'Column heading repaired',
    'date-from-rows': 'Block date taken from the rows' };

  /* ---------- sign-in ---------- */
  const showStep = code => { $('stepEmail').hidden = code; $('stepCode').hidden = !code; (code ? $('loginCode') : $('loginEmail')).focus(); };
  $('loginEmail').value = store.get('amp_email') || '';
  async function sendCode() {
    const email = $('loginEmail').value.trim();
    if (!email) return say($('msgEmail'), 'Enter your email.', 'err');
    $('sendCode').disabled = $('resend').disabled = true; say($('msgEmail'), 'Sending the code…'); say($('msgCode'), '');
    try {
      const j = await call('otpRequest', { email });
      store.set('amp_email', email); $('sentTo').textContent = j.sentTo; $('loginCode').value = '';
      say($('msgEmail'), ''); showStep(true); say($('msgCode'), 'Code sent.', 'ok');
    } catch (e) { say($('stepCode').hidden ? $('msgEmail') : $('msgCode'), e.message, 'err'); }
    $('sendCode').disabled = $('resend').disabled = false;
  }
  $('sendCode').onclick = sendCode;
  $('resend').onclick = sendCode;
  $('loginEmail').onkeydown = e => { if (e.key === 'Enter') sendCode(); };
  $('changeEmail').onclick = () => { showStep(false); say($('msgCode'), ''); };
  $('verifyCode').onclick = async () => {
    $('verifyCode').disabled = true;
    try {
      const j = await call('otpVerify', { email: $('loginEmail').value.trim(), code: $('loginCode').value.trim() });
      session = j.session; store.set('amp_session', session); await start();
    } catch (e) { say($('msgCode'), e.message, 'err'); }
    $('verifyCode').disabled = false;
  };
  $('loginCode').onkeydown = e => { if (e.key === 'Enter') $('verifyCode').click(); };

  function signedOut(why) {
    session = null; me = null; store.set('amp_session', null);
    $('app').hidden = true; $('signin').hidden = false; showStep(false);
    say($('msgEmail'), why || '', why ? 'err' : '');
  }
  $('signOut').onclick = () => { call('signOut').catch(() => {}); signedOut('Signed out.'); $('msgEmail').className = 'msg'; };

  /* ---------- after sign-in ---------- */
  const ROLE = { admin: 'Administrator', operator: 'Engg/CNL', aen: '' };
  const VIEWS = ['vToday', 'vAen', 'vReports', 'vUsers', 'vSettings', 'vPlace', 'vData', 'vSources'];
  const MENU = { admin: [['vToday', 'Today'], ['vReports', 'Reports'], ['vUsers', 'Users'], ['vSettings', 'Settings'], ['vPlace', 'Placements'], ['vSources', 'Sources'], ['vData', 'Data and checks']],
                 operator: [['vToday', 'Today'], ['vReports', 'Reports'], ['vSources', 'Sources']] };
  const LOAD = { vToday: () => loadToday(), vAen: () => loadAen(), vReports: () => loadReports(), vUsers: () => loadAdmin(), vSettings: () => loadAdmin(), vPlace: () => loadPlacements(), vData: () => loadAdmin(), vSources: () => loadSources() };

  async function start() {
    try { me = await call('me'); } catch (e) { if (session) signedOut(e.message); return; }
    $('signin').hidden = true; $('app').hidden = false;
    $('who').textContent = (me.role === 'aen' ? 'AEN/' + me.aen : ROLE[me.role]) + ' · ' + me.email;
    const items = MENU[me.role];
    $('menu').hidden = !items; $('menuItems').innerHTML = '';
    (items || []).forEach(([v, t]) => { const b = el('button', '', t); b.dataset.view = v; b.onclick = () => show(v); $('menuItems').appendChild(b); });
    if (me.role !== 'aen') {
      D = me.today; $('tPick').value = D;
      $('rAen').innerHTML = ''; me.offices.forEach(o => { const op = el('option', '', 'AEN/' + o.aen + ' · ' + town(o.office)); op.value = o.aen; $('rAen').appendChild(op); });
    }
    show(me.role === 'aen' ? 'vAen' : 'vToday');
  }
  function show(v) {
    VIEWS.forEach(id => $(id).hidden = id !== v);
    document.querySelectorAll('#menuItems button').forEach(b => b.setAttribute('aria-current', b.dataset.view === v ? 'page' : 'false'));
    tell('');
    LOAD[v]();
  }

  /* ---------- Word file ---------- */
  let logoBytes = null;   // the Indian Railways logo for the letterhead, fetched once
  const getLogo = async () => { if (!logoBytes) { try { logoBytes = new Uint8Array(await (await fetch('ir-logo-256.png')).arrayBuffer()); } catch (e) { logoBytes = null; } } return logoBytes; };
  // the top sheet's artwork (tools/render-cover-art.js); a picture that cannot be loaded is left out, the text still comes
  let coverArt = null;
  const getCover = async () => { if (!coverArt) { coverArt = {};
    for (const k of ['frame', 'seal', 'rule', 'contents']) { try { coverArt[k] = new Uint8Array(await (await fetch('cover-' + k + '.png')).arrayBuffer()); } catch (e) { coverArt[k] = null; } } }
    return coverArt; };
  async function saveDocx(R) {
    // made safe for Word 2007 (docx-compat.js): the offices' Word 2007 refused the file otherwise (05.11.2026)
    const blob = await word2007Safe(JSZip, await docx.Packer.toBlob(buildMorningDoc(docx, R, { logo: await getLogo(), cover: await getCover() })), 'blob');
    const a = el('a'); a.href = URL.createObjectURL(blob);
    a.download = 'Morning_Position_AEN-' + R.aen.replace('/', '-') + '_' + dmy(R.date) + '.docx';
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 30000);
  }
  async function download(aen, name, btn) {
    if (btn) btn.disabled = true;
    // §20 item 3: a failure is shown where the button is, not only in a passing message
    const host = btn && btn.parentNode;
    let note = host ? host.querySelector('.dl-err') : null;
    if (note) note.remove();
    let s = null;
    try { s = await call('snapshot', { aen, name }); }
    catch (e) { fail('Could not fetch the saved report: ' + e.message); if (btn) btn.disabled = false; return; }
    try { await saveDocx(s.report); }
    catch (e) { console.error(e); fail('Could not build the Word file: ' + (e && e.message || e)); }
    if (btn) btn.disabled = false;
    function fail(text) {
      tell(text, 'err');
      if (host) { note = el('div', 'small err dl-err', text); host.appendChild(note); }
    }
  }

  /* ---------- today (Engg/CNL and admin) ---------- */
  let D = null, today = null;
  const busy = {};          // aen -> 'run' | 'queue'
  const failed = {};        // aen -> message
  let stopping = false, preparing = false, runCount = 0;   // §19: Stop button and the honest counter
  $('tPick').onchange = () => { if ($('tPick').value) { D = $('tPick').value; loadToday(); } };

  // The ten tiles and the Generate buttons are ready at once; nothing waits for the server.
  // Which reports were already saved for the date is filled in quietly in the background.
  const fresh = {};         // aen -> saved in this page for date D (wins over the background answer)
  function blankDay() {
    return { date: D, aens: me.offices.map(o => ({ aen: o.aen, office: o.office, count: 0, latest: null })), sources: null };
  }
  async function loadToday() {
    if (!today || today.date !== D) { today = blankDay(); Object.keys(fresh).forEach(k => delete fresh[k]); }
    drawToday();
    const asked = D;
    try {
      const t = await call('today', { date: asked });
      if (asked !== D || !t || !Array.isArray(t.aens)) return;
      t.aens.forEach(x => { const a = today.aens.find(y => y.aen === x.aen); if (a && !fresh[a.aen]) { a.latest = x.latest; a.count = x.count; } });
      today.sources = t.sources; drawToday();
    } catch (e) { console.error(e); }          // status only; Generate works without it
  }
  function drawHero() {
    $('tDay').textContent = weekday(D) + ' · Morning position';
    $('tDate').textContent = dotted(D);
    const y = addDays(D, -1);
    $('tLine').textContent = 'Day of ' + dmy(y) + ' and the night following · cut-off ' + (me.cutoff || '');
  }
  function stateOf(a) {
    if (busy[a.aen] === 'run') return { cls: 'run', st: 'Generating…' };
    if (busy[a.aen] === 'queue') return { cls: 'wait', st: 'Queued' };
    if (failed[a.aen]) return { cls: 'bad', st: 'Not prepared', gaps: [failed[a.aen]] };
    if (!a.latest) return { cls: 'wait', st: 'Not prepared yet' };
    const all = a.latest.gaps || [], g = all.filter(x => x.kind !== 'note'), notes = all.filter(x => x.kind === 'note');
    return g.length ? { cls: 'gap', st: 'Ready, ' + g.length + (g.length > 1 ? ' gaps' : ' gap'), gaps: g, notes } : { cls: 'ok', st: 'Ready', notes };
  }
  function drawToday() {
    drawHero();
    let ready = 0, gap = 0;
    $('tiles').innerHTML = '';
    today.aens.forEach((a, i) => {
      const s = stateOf(a);
      if (s.cls === 'ok') ready++; if (s.cls === 'gap') gap++;
      const t = el('div', 'tile' + (s.cls === 'gap' || s.cls === 'bad' ? ' gap' : s.cls === 'run' ? ' run' : ''));
      const meta = el('div', 'meta'); meta.append(el('span', '', two(i + 1)), el('span', '', a.latest ? a.latest.at + (a.count > 1 ? ' · ×' + a.count : '') : '—'));
      t.append(meta, el('div', 'code', a.aen), el('div', 'town', town(a.office)), el('div', 'st ' + s.cls, s.st));
      // each gap is a link to its explanation; notes that are not gaps follow in grey (spec §13)
      const gp = el('div', 'gaps');
      (s.gaps || []).forEach(g => { const line = el('div');
        if (typeof g === 'string') line.textContent = g;
        else { const b = el('button', 'why-link', gapText(g)); b.title = asRead(a); b.onclick = () => openWhy(a, g); line.appendChild(b); }
        gp.appendChild(line); });
      (s.notes || []).forEach(g => { const b = el('button', 'why-link note', (SECTION[g.section] || g.section) + ': ' + (REASON[g.reason] || g.reason).toLowerCase());
        b.title = asRead(a); b.onclick = () => openWhy(a, g); gp.appendChild(b); });
      t.appendChild(gp);
      const acts = el('div', 'acts');
      if (!busy[a.aen]) {
        const g = el('button', a.latest ? 'btn-text' : 'btn-line', a.latest ? 'Generate again' : 'Generate'); g.onclick = () => runQueue([a.aen]);
        if (a.latest) { const dl = el('button', 'btn', 'Download'); dl.onclick = () => download(a.aen, a.latest.name, dl); acts.append(dl, g); }
        else acts.appendChild(g);
      }
      t.appendChild(acts); $('tiles').appendChild(t);
    });
    const left = today.aens.length - ready - gap;
    $('cReady').textContent = two(ready); $('cGap').textContent = two(gap); $('cLeft').textContent = two(left);
    // honest counter (spec §19 item 7): done / running / waiting, not "started of ten"
    const running = today.aens.filter(a => busy[a.aen] === 'run').length;
    const waiting = today.aens.filter(a => busy[a.aen] === 'queue').length;
    const done = Math.max(0, runCount - running - waiting);
    const ga = $('genAll');
    if (preparing) {
      $('tProgress').textContent = 'READING EVERY SOURCE ONCE…';
      ga.textContent = 'Reading sources…'; ga.disabled = false;
    } else if (stopping) {
      $('tProgress').textContent = 'STOPPING · ' + running + ' STILL RUNNING';
      ga.textContent = 'Stopping… (' + running + ' still running)'; ga.disabled = running === 0;
    } else if (running + waiting > 0) {
      $('tProgress').textContent = 'GENERATING · ' + done + ' DONE · ' + running + ' RUNNING · ' + waiting + ' WAITING';
      ga.textContent = 'Stop'; ga.disabled = false;
    } else {
      $('tProgress').textContent = left ? left + ' TO GO' : 'ALL READY';
      ga.textContent = 'Generate all ten'; ga.disabled = false;
    }
    drawSources(today.sources);
  }
  function drawSources(s) {
    const box = $('sources'); box.innerHTML = '';
    const line = (k, v) => { const r = el('div', 'kv'); r.append(el('span', '', k), el('span', 'mono', v)); box.appendChild(r); };
    const when = x => x ? (x.at.slice(0, 10) === D ? x.at.slice(11) : dmy(x.at.slice(0, 10)) + ' ' + x.at.slice(11)) : 'not yet';
    line('Meeting sheet', 'live');
    if (!s) { line('TMS data', '…'); return; }
    line('TMS items', when(s.items));
    line('TMS machines', when(s.machines));
    line('TMS spot attention', when(s.spotAttention));
    line('TMS worksites', when(s.worksites));
    line('TMS inspection plan', when(s.inspPlan));
    line('TMS Control', s.tmsControl && s.tmsControl.at ? when({ at: String(s.tmsControl.at).replace('T', ' ').slice(0, 16) }) : 'not found');
  }
  /** The explanation of one gap or note: the tab link first, then what was looked for, what was found and why (spec §13). */
  function openWhy(a, g) {
    const box = $('whyBody'); box.innerHTML = '';
    $('whyTitle').textContent = 'AEN/' + a.aen + ' · ' + (SECTION[g.section] || g.section);
    const w = g.why || {}, where = w.where || {};
    const row = (k, v) => { if (!v) return; const r = el('div', 'why-row'); r.append(el('div', 'label', k), v instanceof Node ? v : el('div', '', v)); box.appendChild(r); };
    const src = el('div'); src.appendChild(document.createTextNode(where.label || 'Source not recorded'));
    if (where.url) { src.appendChild(el('br')); const l = el('a', '', 'Open the tab ↗'); l.href = where.url; l.target = '_blank'; l.rel = 'noopener'; src.appendChild(l); }
    if (where.at) src.appendChild(el('div', 'small', 'Uploaded ' + where.at));
    row('Where', src);
    row('Reason', REASON[w.reason || g.reason] || (g.kind === 'error' ? REASON['source-error'] : ''));
    row('Looked for', w.lookedFor);
    row('Found', w.found || (g.kind === 'note' ? g.text : g.kind === 'error' ? g.text : ''));
    if (!g.why && g.kind !== 'note') row('Note', 'This report was prepared before explanations were recorded; generate it again to see why.');
    if (g.placements && me.role === 'admin') { const b = el('button', 'btn-line', 'Open Placements'); b.onclick = () => { $('why').close(); show('vPlace'); }; box.appendChild(b); }
    // what this panel is showing is what was read when the report was made, not the sheet as it stands now
    row('As read', asRead(a));
    const foot = el('div', 'why-foot');
    foot.appendChild(el('div', 'small', 'Sheet corrected since? Press Generate again.'));
    const acts = el('div', 'why-acts'), said = el('div', 'small');
    if (canGenerate() && g.section) {
      const rc = el('button', 'btn-line', 'Re-check');
      rc.onclick = async () => { rc.disabled = true; said.className = 'small'; said.textContent = 'Re-reading the source…';
        try {
          const v = await call('recheck', { aen: a.aen, date: today.date, section: g.section });
          said.className = 'small ' + (v.fixed ? 'ok' : '');
          said.textContent = v.fixed
            ? 'Fixed now: found a filled block for ' + dmy(v.dataDate) + ' — the Sources tab shows it; Generate again to put it in the report'
            : 'Still the same: ' + (v.found || v.error || (v.dataDate ? 'newest entry dated ' + dmy(v.dataDate) : 'nothing newer')) + ' (re-read at ' + v.at.slice(11) + ')';
        } catch (e) { said.className = 'small err'; said.textContent = e.message; }
        rc.disabled = false; };
      acts.appendChild(rc);
    }
    if (canGenerate()) { const gen = el('button', 'btn', 'Generate again'); gen.onclick = () => { $('why').close(); runQueue([a.aen]); }; acts.appendChild(gen); }
    foot.append(acts, said); box.appendChild(foot);
    $('why').showModal();
  }
  /** "as read at 11:26" — when the report this panel came from was made (spec §16 item 11). */
  const asRead = a => 'as read at ' + ((a.latest && a.latest.at) || '—');
  const canGenerate = () => me && (me.role === 'admin' || me.role === 'operator');
  $('whyClose').onclick = () => $('why').close();
  // "Generate all ten" reads every source once (the batch store, §19) and then builds the ten from it.
  // While it runs the same button is Stop. A single tile's Generate / Generate again reads fresh (no batch).
  $('genAll').onclick = () => { if (Object.keys(busy).length || preparing) stopAll(); else runQueue(today.aens.map(a => a.aen), true); };
  function stopAll() {
    stopping = true;                                     // start no more offices
    Object.keys(busy).forEach(a => { if (busy[a] === 'queue') delete busy[a]; });   // waiting ones go back to "Not prepared yet"
    drawToday();                                         // offices already running finish on their own
  }
  function runQueue(list, batch) { return runQueue_(list, batch).catch(e => { preparing = false; tell(e.message, 'err'); console.error(e); }); }
  async function runQueue_(list, batch) {
    tell('');
    list = list.filter(a => !busy[a]);
    if (!list.length) return;
    stopping = false; runCount = list.length;
    list.forEach(a => { busy[a] = 'queue'; delete failed[a]; });
    drawToday();
    const endRun = () => { stopping = false; preparing = false; runCount = 0; };
    let batchId = null;
    if (batch) {
      try {
        preparing = true; drawToday();
        const p = await prepareParts(D);
        preparing = false;
        if (stopping) { list.forEach(a => delete busy[a]); endRun(); drawToday(); return; }   // Stop pressed while reading
        if (p.stopped) { list.forEach(a => delete busy[a]); endRun(); tell(p.message || 'Reading the sources was stopped — Generate again', 'err'); drawToday(); return; }
        batchId = p.batchId;
      } catch (e) { list.forEach(a => delete busy[a]); endRun(); drawToday(); tell(e.message, 'err'); return; }
    }
    const one = async aen => {
      if (stopping) { delete busy[aen]; drawToday(); return; }
      busy[aen] = 'run'; drawToday();
      try {
        const j = await call('generate', batchId ? { date: D, aen, batch: batchId } : { date: D, aen });
        if (j.stopped) { failed[aen] = j.step || 'stopped at 5 minutes'; }   // §19: a self-stopped build shows as failed on the tile
        else { const a = today.aens.find(x => x.aen === aen); if (a) { a.latest = j.info; a.count = (a.count || 0) + 1; fresh[aen] = true; } }
      } catch (e) { failed[aen] = e.message; }
      delete busy[aen]; drawToday();
    };
    const todo = list.slice();
    const worker = async () => { while (todo.length && !stopping) await one(todo.shift()); };
    await Promise.all(Array.from({ length: batchId ? 5 : 3 }, worker));   // five offices at a time from the store
    endRun(); drawToday();
  }
  /**
   * Read the shared sources in four parts side by side (spec §19): meeting/IOW/TMS, PIR tabs, PWI tabs, work plans.
   * The work plans go first and alone: a server that knows parts answers with part 'plans' in a few seconds, and
   * the other three then start together under the same id; an older server ignores the part, reads everything
   * once and answers with its own id, which is used as it is.
   */
  async function prepareParts(D) {
    const id = D + '_' + Math.random().toString(36).slice(2, 10).replace(/[^a-z0-9]/g, 'x').padEnd(8, 'x');
    const first = await call('prepare', { date: D, part: 'plans', batchId: id });
    if (first.stopped || first.part !== 'plans') return first;              // older server: a whole read already done
    const rest = await Promise.all(['meeting', 'pir', 'pwi'].map(part => call('prepare', { date: D, part, batchId: id })));
    const bad = rest.find(r => r.stopped);
    return bad || { batchId: id, date: D };
  }

  /* ---------- ruled list of saved reports ---------- */
  function drawList(box, h, aen, skipToday) {
    box.innerHTML = '';
    const head = el('div', 'lrow head'); head.append(el('span', '', 'No.'), el('span', '', 'Report date'), el('span', 'hide-sm', 'Prepared'), el('span', ''));
    box.appendChild(head);
    const rows = h.snapshots.filter(s => !s.again && !(skipToday && s.date === me.today));
    if (!rows.length) { box.appendChild(el('div', 'empty', 'No reports yet.')); return; }
    rows.slice(0, 60).forEach((s, i) => {
      const r = el('div', 'lrow');
      const d = el('span', ''); d.style.fontWeight = '500'; d.textContent = dmy(s.date) + ' ';
      if (s.times > 1) d.appendChild(el('span', 'small', 'prepared ' + s.times + ' times'));
      const b = el('button', 'btn-line', 'Download'); b.onclick = () => download(aen, s.name, b);
      const end = el('span', 'end'); end.appendChild(b);
      r.append(el('span', 'no', two(i + 1)), d, el('span', 'mono hide-sm', s.at), end);
      box.appendChild(r);
    });
  }

  /* ---------- AEN office ---------- */
  async function loadAen() {
    $('aOffice').textContent = (me.office || '').replace(/\s*\(.*\)\s*$/, '');
    $('aCode').textContent = 'AEN/' + me.aen;
    const tb = $('aToday'); tb.innerHTML = ''; tb.appendChild(el('div', 'small', 'Reading…'));
    try {
      const h = await call('history', {});
      const t = h.snapshots.find(s => s.date === me.today);
      tb.innerHTML = '';
      const info = el('div'); info.append(el('div', 'label', 'Today'), el('div', 'd', dotted(me.today)));
      if (t) {
        info.appendChild(el('div', 'small', 'Ready since ' + t.at + ' · Word file'));
        const b = el('button', 'btn', 'Download report'); b.onclick = () => download(me.aen, t.name, b);
        tb.append(info, b);
      } else {
        info.appendChild(el('div', 'small', 'Not prepared yet. Engg/CNL prepares it each morning; please check again shortly.'));
        tb.appendChild(info);
      }
      drawList($('aList'), h, me.aen, true);
    } catch (e) { tb.innerHTML = ''; tell(e.message, 'err'); }
  }

  /* ---------- reports (Engg/CNL and admin) ---------- */
  $('rAen').onchange = () => loadReports();
  async function loadReports() {
    const aen = $('rAen').value; $('rList').innerHTML = '';
    try { drawList($('rList'), await call('history', { aen }), aen, false); } catch (e) { tell(e.message, 'err'); }
  }

  /* ---------- admin ---------- */
  let users = [], admins = [];
  async function loadAdmin() {
    try {
      const s = await call('adminState');
      users = s.users || []; admins = s.admins || [];
      drawUsers();
      $('cutoff').value = s.cutoff || '';
      $('keyState').textContent = s.uploadKeySet ? 'A key is set.' : 'No key yet.';
      $('dTmsc').textContent = s.tmsControl.ok ? 'Read correctly. Made ' + String(s.tmsControl.generatedAt).replace('T', ' ').slice(0, 16) + ' for ' + (s.tmsControl.aens || []).length + ' AENs.' : 'Not read: ' + s.tmsControl.error;
      const ub = $('dUploads'); ub.innerHTML = '';
      const KIND = { items: 'TMS items', machines: 'TMS machine progress', stations: 'Station list' };
      (s.uploads || []).forEach(u => { const tr = el('tr'); [u.at, KIND[u.kind] || u.kind, u.date ? dmy(u.date) : '—', u.rows == null ? '' : u.rows].forEach(v => tr.appendChild(el('td', '', v))); ub.appendChild(tr); });
      if (!(s.uploads || []).length) { const tr = el('tr'); const td = el('td', '', 'Nothing received yet.'); td.colSpan = 4; tr.appendChild(td); ub.appendChild(tr); }
      const iss = $('dIssues'); iss.innerHTML = '';
      if (s.selfCheck && s.selfCheck.at) { const p = (s.selfCheck.problems || []).length;
        const sc = el('p', 'small' + (p ? ' err' : ''), 'Source check ' + String(s.selfCheck.at).slice(11, 16) + ': '
          + (p ? p + ' source' + (p > 1 ? 's' : '') + " can't be read — " + s.selfCheck.problems.join(', ') : 'every source resolved'));
        iss.appendChild(sc); }
      Object.keys(s.issues || {}).sort().reverse().slice(0, 7).forEach(d => Object.keys(s.issues[d]).forEach(a => {
        // rows already answered on the Placements page are not shown again (spec §12)
        const placed = new Set(s.placed || []), open = (it => (it.unassigned || []).filter(x => { const i = x.indexOf(': ');
          return !placed.has(placeKey(i < 0 ? x : x.slice(0, i), i < 0 ? '' : x.slice(i + 2))); }));
        const it = s.issues[d][a], lines = [...open(it).map(x => 'Not placed: ' + x), ...(it.timeFixes || []).map(x => 'Time read from a date: ' + x),
          ...(it.dateSwaps || []).map(x => 'Date read as dd.mm after a Sheets swap: ' + x)];
        if (!lines.length) return;
        iss.appendChild(el('b', '', dmy(d) + ' · AEN/' + a));
        const ul = el('ul'); lines.forEach(l => ul.appendChild(el('li', '', l))); iss.appendChild(ul);
      }));
      if (!iss.querySelector('b')) iss.appendChild(el('p', 'empty', 'No placement or date issues to look at.'));
      drawSourcesTable(s.sourcesTable || {});
      drawRepairs(s.repairs || {});
    } catch (e) { tell(e.message, 'err'); console.error(e); }
  }
  /** Column repairs the program made on its own (spec §16): Keep stores the new name (note stops); Undo makes it a gap again. */
  function drawRepairs(rep) {
    const box = $('dRepairs'); box.innerHTML = '';
    const tabs = Object.keys(rep).filter(t => (rep[t].list || []).length || (rep[t].dates || []).length);
    if (!tabs.length) { box.appendChild(el('p', 'empty', 'No column changes — every heading was read as known.')); return; }
    tabs.forEach(t => {
      box.appendChild(el('b', '', rep[t].tab + (rep[t].at ? ' · ' + rep[t].at : '')));
      // a block date the rows decided (§16 item 10): read afresh every time, so there is nothing to keep or undo
      (rep[t].dates || []).forEach(d => {
        const row = el('div', 'repair-row');
        row.appendChild(el('span', '', d.why === 'assumed'
          ? 'two blocks headed ' + dmy(d.was) + ', no date in the rows — the upper one taken as ' + dmy(d.now)
          : 'heading read ' + dmy(d.was) + ', rows dated ' + dmy(d.now).slice(0, 5) + ' — taken as ' + dmy(d.now).slice(0, 5)
            + (d.of ? ' (' + d.agree + ' of ' + d.of + ' rows)' : '')));
        row.appendChild(el('span', 'repair-acts small', 'read afresh each time'));
        box.appendChild(row);
      });
      (rep[t].list || []).forEach(r => {
        const row = el('div', 'repair-row');
        row.appendChild(el('span', '', '"' + r.from + '" → "' + r.to + '"' + (r.via && r.via !== 'alt' ? ' (by ' + r.via + ')' : '')));
        const acts = el('span', 'repair-acts');
        const keep = el('button', 'btn-line', 'Keep'); keep.onclick = () => decideRepair('keepRepair', r.tab || rep[t].tab, r.col, r.to, keep);
        const undo = el('button', 'btn-text', 'Undo'); undo.onclick = () => decideRepair('undoRepair', r.tab || rep[t].tab, r.col, r.to, undo);
        acts.append(keep, undo); row.appendChild(acts); box.appendChild(row);
      });
    });
  }
  async function decideRepair(fn, tab, col, to, btn) {
    btn.disabled = true;
    try { const s = await call(fn, { tab, col, to }); drawRepairs(s.repairs || {}); }
    catch (e) { btn.disabled = false; tell(e.message, 'err'); }
  }
  /**
   * Sources down, the ten AENs across: the date each AEN's latest report found in each source (spec §13). Green: current by
   * the section's own rule; amber: one day older than that; red: older, none or unreadable. Each cell links to the tab.
   */
  /* ---------- Sources (spec §17): every source, linked, read-only ---------- */
  const STATE = { current: 'Current', repaired: 'Repaired', nothingForAen: 'Nothing for this AEN',
    olderData: 'Older data', readerNotRun: 'Reader not run today', cantRead: "Can't read", notGenerated: 'Not generated yet' };
  let sources = null;
  async function loadSources() {
    const box = $('sTable');
    if (!sources || sources.date !== D) { box.innerHTML = ''; box.appendChild(el('p', 'empty', 'Reading the list of tabs…')); }
    try { sources = await call('sources', { date: D }); } catch (e) { box.innerHTML = ''; box.appendChild(el('p', 'empty', e.message)); return; }
    drawSheets(sources); drawSourceRows(sources);
    try { const st = await call('adminState'); drawSourcesTable(st.sourcesTable || {}); } catch (e) { /* the per-AEN table is extra */ }
  }
  function drawSheets(v) {
    const box = $('sSheets'); box.innerHTML = '';
    v.sheets.forEach(sh => {
      const row = el('div', 'sheet-row');
      row.appendChild(el('span', '', sh.name));
      if (sh.url) { const a = el('a', '', 'Open sheet ↗'); a.href = sh.url; a.target = '_blank'; a.rel = 'noopener'; row.appendChild(a); }
      if (sh.error) row.appendChild(el('span', 'small err', sh.error));
      box.appendChild(row);
    });
    if (v.account) box.appendChild(el('p', 'small', 'All of them are read with ' + v.account + ', and only read.'));
    if (v.selfCheck && v.selfCheck.at) { const p = v.selfCheck.problems || 0;
      box.appendChild(el('p', 'small' + (p ? ' err' : ''), 'Source check at ' + String(v.selfCheck.at).slice(11, 16) + ': '
        + (p ? p + ' source' + (p > 1 ? 's' : '') + " can't be read — see the red rows below" : 'every source resolved'))); }
    if (v.generated) { const n = v.generated.length;
      box.appendChild(el('p', 'small', n + ' of the ' + v.offices + ' offices prepared for ' + dmy(v.date)
        + (n ? ' \u2014 ' + v.generated.join(', ') + '. A source only they read reports what they found; "Not generated yet" means no report has reached it yet.'
             : '. Nothing has been read yet today, so every state below says so.'))); }
  }
  function drawSourceRows(v) {
    const box = $('sTable'); box.innerHTML = '';
    const tb = el('table', 't src-list'), thead = el('thead'), hd = el('tr');
    ['Report section', 'Source', 'Open', 'Newest data found', 'Last read', 'State'].forEach(h => hd.appendChild(el('th', '', h)));
    thead.appendChild(hd); tb.appendChild(thead);
    const body = el('tbody');
    let part = null;
    v.rows.forEach(r => {
      if (r.no !== part) { part = r.no; const sep = el('tr', 'part-row'); const td = el('td', '', r.no + '  ' + r.part);
        td.colSpan = 6; sep.appendChild(td); body.appendChild(sep); }
      const tr = el('tr');
      tr.appendChild(el('td', 'src-sec', r.name + (r.aen ? ' · ' + r.aen : '')));
      const nm = el('td', 'src-name'); nm.appendChild(document.createTextNode(r.source));
      if (r.renamed) nm.appendChild(el('div', 'small', 'renamed since the spec; read under this name'));
      if (r.menu) nm.appendChild(el('div', 'small', r.menu));
      tr.appendChild(nm);
      const op = el('td');
      if (r.url) { const a = el('a', '', 'Open ↗'); a.href = r.url; a.target = '_blank'; a.rel = 'noopener'; op.appendChild(a); }
      else op.appendChild(el('span', 'small', r.tms || r.tmsc ? 'read by the TMS reader' : '—'));
      tr.appendChild(op);
      const dt = el('td', 'src-c ' + dateClass(r));
      dt.appendChild(el('span', '', r.date ? dmy(r.date) : (r.read ? 'none' : '—')));
      if (!r.sameEverywhere) { const u = el('div', 'small');
        u.textContent = r.perAen.map(x => x.aen + ' ' + (x.date ? dmy(x.date).slice(0, 5) : '—')).join(' · '); dt.appendChild(u); }
      tr.appendChild(dt);
      tr.appendChild(el('td', 'src-c', r.lastRead ? String(r.lastRead).slice(11, 16) : '—'));
      const st = el('td', 'src-state');
      const label = sourceStateLabel(r);
      const b = el('button', 'why-link state-' + r.state, label);
      b.onclick = () => openSourceWhy(r);
      st.appendChild(b);
      // §21 item 2: a red row says what Google said, not only "Can't read"
      if (r.state === 'cantRead' && r.error) st.appendChild(el('div', 'small err', r.error));
      if (r.repairs || r.dateMoves) st.appendChild(el('div', 'small', [r.repairs ? r.repairs + ' heading' : '', r.dateMoves ? 'block date from the rows' : ''].filter(Boolean).join(' · ')));
      const said = el('div', 'small');
      const rc = el('button', 'btn-text', 'Re-check');
      rc.onclick = async () => { rc.disabled = true; said.className = 'small'; said.textContent = 'Re-reading…';
        try { const x = await call('recheck', { aen: r.whyAen || me.offices[0].aen, date: D, section: r.section });
          said.className = 'small ' + (x.fixed ? 'ok' : '');
          said.className = 'small ' + (x.fixed ? 'ok' : 'err');
          said.textContent = x.fixed ? 'Fixed now' + (x.dataDate ? ': ' + dmy(x.dataDate) : '')
            : 'Still the same' + (x.error ? ': ' + x.error : x.dataDate ? ': newest ' + dmy(x.dataDate) : '') + (x.at ? ' (' + String(x.at).slice(11) + ')' : '');
          // §21 item 3: the server has already put the new read in place of the old one; redraw the rows from it
          if (x.fixed) setTimeout(() => { if (sources && sources.date === D) loadSources(); }, 1200);
        } catch (e) { said.className = 'small err'; said.textContent = e.message; }
        rc.disabled = false; };
      st.append(rc, said);
      tr.appendChild(st);
      body.appendChild(tr);
    });
    tb.appendChild(body); box.appendChild(tb);
  }
  /** "Older for BWN, KWAE — 02.11" when only some AENs are behind; "Older data — 02.11" when all are (§21 item 4). */
  function sourceStateLabel(r) {
    if (r.state !== 'olderData') return STATE[r.state] || r.state;
    const d = r.behindDate || r.date, when = d ? ' — ' + dmy(d).slice(0, 5) : '';
    const few = r.behind && r.behind.length && r.perAen && r.behind.length < r.perAen.length;
    return (few ? 'Older for ' + r.behind.join(', ') : STATE.olderData) + when;
  }
  const dateClass = r => ({ current: 'green', repaired: 'green', olderData: 'amber', readerNotRun: 'amber',
    cantRead: 'red', nothingForAen: 'neutral', notGenerated: 'grey' }[r.state] || 'grey');
  /** The same explanation the tile shows, opened from a source row. */
  async function openSourceWhy(r) {
    if (!r.whyAen || !r.whySnapshot) { tell('No saved report has read this source yet — generate one first.', 'err'); return; }
    try {
      const j = await call('snapshot', { aen: r.whyAen, name: r.whySnapshot });
      const sec = (j.report.sections || {})[r.section] || {};
      const g = { section: r.section, kind: sec.ok ? 'stale' : 'error',
        text: sec.ok ? 'latest available data used' : sec.error, why: (sec.why || {}).gap || null };
      openWhy({ aen: r.whyAen, latest: { at: r.whySnapshot.slice(11, 13) + ':' + r.whySnapshot.slice(13, 15) } }, g);
    } catch (e) { tell(e.message, 'err'); }
  }
  function drawSourcesTable(st) {
    const box = $('dSources'); box.innerHTML = '';
    const aens = me.offices.map(o => o.aen).filter(a => st[a]);
    if (!aens.length) { box.appendChild(el('p', 'empty', 'No report has been prepared since this table was added.')); return; }
    const order = [], cell = {};
    aens.forEach(a => (st[a].rows || []).forEach(r => { if (!order.includes(r.source)) order.push(r.source); const k = r.source + '|' + a;
      if (!cell[k] || r.error || (r.stale && !cell[k].error)) cell[k] = r; }));
    const tb = el('table', 't src-t'), hd = el('tr'); hd.appendChild(el('th', '', 'Source'));
    aens.forEach(a => { const th = el('th', ''); th.append(el('div', '', a), el('div', 'small', dmy(st[a].date).slice(0, 5) + ' ' + String(st[a].generatedAt || '').slice(11))); hd.appendChild(th); });
    const thead = el('thead'); thead.appendChild(hd); tb.appendChild(thead);
    const body = el('tbody');
    order.forEach(src => { const tr = el('tr'); tr.appendChild(el('td', 'src-name', src));
      aens.forEach(a => { const r = cell[src + '|' + a], td = el('td', 'src-c');
        if (!r) { td.textContent = ''; tr.appendChild(td); return; }
        const behind = r.date && r.stale && r.date >= addDays(st[a].date, -2);
        td.className = 'src-c ' + (r.error || !r.date ? 'red' : !r.stale ? 'green' : behind ? 'amber' : 'red');
        const txt = r.error ? 'not read' : r.date ? dmy(r.date).slice(0, 5) : 'none';
        const v = r.url ? el('a', '', txt) : el('span', '', txt); if (r.url) { v.href = r.url; v.target = '_blank'; v.rel = 'noopener'; }
        td.appendChild(v); if (r.at) td.appendChild(el('div', 'small', String(r.at).slice(11, 16)));
        td.title = [r.error, r.reason ? REASON[r.reason] : '', r.at ? 'uploaded ' + r.at : ''].filter(Boolean).join(' · ');
        tr.appendChild(td); });
      body.appendChild(tr); });
    tb.appendChild(body); box.appendChild(tb);
  }
  /** The server's placementKey(): where without dates + the row's text, upper case, spaces collapsed. */
  const placeKey = (where, text) => String(where || '').toUpperCase().replace(/\b\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}\b/g, '').replace(/\s+/g, ' ').trim()
    + '|' + String(text || '').toUpperCase().replace(/\s+/g, ' ').trim();

  /* ---------- admin: Placements (spec §12) ---------- */
  async function loadPlacements() {
    try { drawPlacements(await call('placements')); } catch (e) { tell(e.message, 'err'); console.error(e); }
  }
  async function answer(where, text, ans) {
    try { drawPlacements(await call('savePlacement', { where, text, answer: ans })); tell(ans == null ? 'Answer removed.' : 'Saved. It is used in every report from now on.', 'ok'); }
    catch (e) { tell(e.message, 'err'); }
  }
  function drawPlacements(v) {
    const ob = $('pOpen'), db = $('pDone'); ob.innerHTML = ''; db.innerHTML = '';
    const none = (tb, n, t) => { const tr = el('tr'), td = el('td', 'empty', t); td.colSpan = n; tr.appendChild(td); tb.appendChild(tr); };
    (v.open || []).forEach(o => {
      const tr = el('tr');
      tr.appendChild(el('td', '', o.where)); tr.appendChild(el('td', '', o.text || '—'));
      tr.appendChild(el('td', 'small', o.days.slice(-3).map(dmy).join(', ') + (o.days.length > 3 ? ' …' : '')));
      const td = el('td'), row = el('div', 'row'), sel = el('select', 'field');
      sel.appendChild(el('option', '', 'Belongs to …')); sel.options[0].value = '';
      (v.aens || []).forEach(a => { const op = el('option', '', 'AEN/' + a); op.value = a; sel.appendChild(op); });
      sel.onchange = () => { if (sel.value) answer(o.where, o.text, sel.value); };
      const ig = el('button', 'btn-line', 'Ignore'), sw = el('button', 'btn-line', 'Date swap');
      ig.onclick = () => answer(o.where, o.text, 'ignore'); sw.onclick = () => answer(o.where, o.text, 'swap');
      [sel, ig, sw].forEach(x => row.appendChild(x)); td.appendChild(row); tr.appendChild(td); ob.appendChild(tr);
    });
    if (!(v.open || []).length) none(ob, 4, 'Nothing waiting. Every row is placed.');
    (v.answered || []).forEach(o => {
      const tr = el('tr');
      [o.where, o.text || '—', o.aen ? 'AEN/' + o.aen : o.ignore ? 'Ignore' : 'Date swap', (o.by || '') + (o.at ? ' · ' + o.at : '')].forEach(x => tr.appendChild(el('td', '', x)));
      const td = el('td'), b = el('button', 'btn-line', 'Undo'); b.onclick = () => answer(o.where, o.text, null); td.appendChild(b); tr.appendChild(td); db.appendChild(tr);
    });
    if (!(v.answered || []).length) none(db, 5, 'No answers yet.');
  }

  function drawUsers() {
    const tb = $('usersBody'); tb.innerHTML = '';
    admins.forEach((a, i) => {
      const tr = el('tr');
      [a, i ? 'Administrator (backup)' : 'Administrator', 'All', i ? '' : 'P.Way Planning', 'Yes', ''].forEach(v => tr.appendChild(el('td', 'fixed', v)));
      tb.appendChild(tr);
    });
    users.forEach((u, i) => {
      const tr = el('tr');
      const email = Object.assign(el('input', 'field'), { type: 'email', value: u.email || '', placeholder: 'email' }); email.setAttribute('aria-label', 'Email');
      const role = el('select', 'field'); role.setAttribute('aria-label', 'Role');
      [['aen', 'AEN office (download)'], ['operator', 'Engg/CNL (prepare)']].forEach(([v, t]) => { const o = el('option', '', t); o.value = v; o.selected = (u.role || 'aen') === v; role.appendChild(o); });
      const aen = el('select', 'field'); aen.setAttribute('aria-label', 'AEN');
      const ch = el('option', '', 'Choose'); ch.value = ''; aen.appendChild(ch);
      me.offices.forEach(o => { const op = el('option', '', 'AEN/' + o.aen); op.value = o.aen; op.selected = o.aen === u.aen; aen.appendChild(op); });
      const name = Object.assign(el('input', 'field'), { value: u.name || '' }); name.setAttribute('aria-label', 'Name or post');
      const act = Object.assign(el('input'), { type: 'checkbox', checked: u.active !== false }); act.setAttribute('aria-label', 'Active');
      const del = el('button', 'btn-text', 'Remove'); del.style.color = 'var(--bad)';
      const sync = () => { aen.disabled = role.value === 'operator'; if (aen.disabled) aen.value = ''; }; role.onchange = sync; sync();
      const keep = () => Object.assign(u, { email: email.value.trim(), role: role.value, aen: aen.value, name: name.value.trim(), active: act.checked });
      [email, role, aen, name, act].forEach(x => x.addEventListener('change', keep));
      del.onclick = () => { users.splice(i, 1); drawUsers(); };
      [email, role, aen, name, act, del].forEach(x => { const td = el('td'); td.appendChild(x); tr.appendChild(td); });
      tb.appendChild(tr);
    });
  }
  $('addUser').onclick = () => { users.push({ email: '', role: 'aen', aen: '', active: true }); drawUsers(); const f = $('usersBody').querySelectorAll('input[type=email]'); f[f.length - 1].focus(); };
  $('saveUsers').onclick = async () => {
    $('saveUsers').disabled = true;
    try { const j = await call('saveUsers', { users: users.filter(u => u.email) }); users = j.users; drawUsers(); tell('Users saved.', 'ok'); }
    catch (e) { tell(e.message, 'err'); }
    $('saveUsers').disabled = false;
  };
  $('saveCutoff').onclick = async () => {
    try { await call('setCutoff', { cutoff: $('cutoff').value }); me.cutoff = $('cutoff').value; tell('Cut-off saved.', 'ok'); } catch (e) { tell(e.message, 'err'); }
  };
  $('newKey').onclick = async () => {
    try { const j = await call('newUploadKey'); $('keyOut').hidden = false; $('keyOut').textContent = j.uploadKey; $('keyState').textContent = 'New key made. Copy it now.'; }
    catch (e) { tell(e.message, 'err'); }
  };

  /* ---------- go ---------- */
  if (!C.API_URL) { say($('msgEmail'), 'This site is not configured yet.', 'err'); $('sendCode').disabled = true; }
  else if (session) start();
})();
