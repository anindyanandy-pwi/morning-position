/* §25 preview app. One site; the screen changes by role, enforced on the server. Talks to the live server through
 * AMP_CONFIG.API_URL. New routes (seen / reportMistake / mistakes / email) are feature-detected: where the live
 * server does not have them yet, the control says so instead of failing. */
(function () {
  const $ = id => document.getElementById(id);
  const el = (t, c, x) => { const n = document.createElement(t); if (c) n.className = c; if (x != null) n.textContent = x; return n; };
  const store = { get: k => { try { return localStorage.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (e) {} }, del: k => { try { localStorage.removeItem(k); } catch (e) {} } };
  // §30 switch-over: both sites may be live at once (old portal at one path, redesign at another). The session token
  // is the same for both; the old portal keeps it under 'amp_session', this app under 'amp_sess'. Read either, and
  // write BOTH on every sign-in/sign-out, so moving between the two sites never signs anyone out.
  const SESS_KEYS = ['amp_sess', 'amp_session'];
  const getSession = () => store.get('amp_sess') || store.get('amp_session') || '';
  const setSession = v => SESS_KEYS.forEach(k => store.set(k, v));
  const delSession = () => SESS_KEYS.forEach(k => store.del(k));
  const dmy = iso => iso ? iso.split('-').reverse().join('.') : '—';
  const addDays = (iso, n) => { const d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const weekday = iso => new Date(iso + 'T00:00:00Z').toLocaleDateString('en-IN', { weekday: 'short', timeZone: 'UTC' });
  const town = o => String(o || '').replace(/^Assistant Divisional Engineer[\s,/]*/i, '').replace(/\s*\(.*\)\s*$/, '');
  let me = null, missing = {};   // missing[fn]=true when the live server lacks a route
  // section id -> number/title (for the Report-a-mistake picker and the gap labels). The report itself is the .docx.
  window.SECTION_TITLES = {
    miniControl: { no: 1, title: 'AEN mini control' }, pwiMc: { no: 1, title: 'PWI mini control' }, iowMc: { no: 1, title: 'IOW mini control' },
    gps: { no: 2, title: 'Winter patrolling' }, major: { no: 3, title: 'Major failures' }, minor: { no: 3, title: 'Minor failures' }, pir: { no: 3, title: 'PIR / failures detail' },
    tdcLapse: { no: 4, title: 'Caution past TDC' }, pqrsTdc: { no: 4, title: 'PQRS TDC fail' }, protection: { no: 4, title: 'Protection caution' },
    badRoad: { no: 5, title: 'Bad road surface at LC gates' }, niLc: { no: 6, title: 'NI LC inspection' }, cctv: { no: 7, title: 'CCTV recording check' },
    spot: { no: 8, title: 'Spot attention due' }, workPlan: { no: 9, title: 'Work planned (next day)' }, inspPlan: { no: 10, title: 'Inspections due' },
    kpi: { no: 11, title: 'KPI and inspections' }, blocks: { no: 12, title: 'Blocks availed and output' }, machines: { no: 13, title: 'Machine utilisation' },
    tms: { no: 14, title: 'TMS dashboard' }, tmsWorksites: { no: 15, title: 'Worksite entry in TMS' }, pwiSites: { no: 16, title: 'Worksite safety arrangement' }, iowSites: { no: 16, title: 'IOW worksite safety' }
  };
  const ST = window.SECTION_TITLES, secTitle = id => (ST[id] ? ST[id].title : id);

  // §19 item 3/6: Google kills a web-app call at ~6 min by dropping the connection (looks like a network error) or by
  // returning a 200 whose body is its own HTML kill page (not JSON). The time taken tells a timeout apart from a real
  // network failure, so "Generate again" is shown instead of "check the internet". Same wording as the old portal.
  const GIVE_UP_MS = 390000;   // 6 min 30 s
  async function call(fn, args) {
    let r; const t0 = Date.now();
    const ctrl = (typeof AbortController !== 'undefined') ? new AbortController() : null;
    const timer = ctrl ? setTimeout(() => ctrl.abort(), GIVE_UP_MS) : null;
    try {
      r = await fetch(AMP_CONFIG.API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ session: getSession(), fn, args: args || {} }), signal: ctrl ? ctrl.signal : undefined });
    } catch (e) {
      if (timer) clearTimeout(timer);
      if (ctrl && e && e.name === 'AbortError') throw new Error('The server took more than 6½ minutes and did not answer — Generate again.');
      if (Date.now() - t0 > 330000) throw new Error('The server took more than 6 minutes and was stopped — Generate again.');
      throw new Error('The server could not be reached. Check the internet connection and try again.');
    }
    if (timer) clearTimeout(timer);
    if (!r.ok) throw new Error('The server did not answer (' + r.status + ').');
    let j; try { j = await r.json(); } catch (e) { throw new Error('The server took more than 6 minutes and was stopped — Generate again.'); }
    if (!j.ok) { if (/unknown fn|not allowed|no such/i.test(j.error || '')) missing[fn] = true;
      if (/sign in again|Please sign in|no longer registered|expired/i.test(j.error || '') && fn !== 'otpVerify') signedOut(); throw new Error(j.error || 'Server error'); }
    return j;
  }
  function signedOut() { delSession(); me = null; clearCaches(); show('login'); }

  /* ---------- shared report blob (same as Download) ---------- */
  // one builder for both sites (web/docx-build.js): logo + cover artwork, Word-2007 safe
  const docxBlob = R => AMPDocx.docxBlob(R);
  const docxName = R => AMPDocx.docxName(R);

  /* ---------- views ---------- */
  function show(name) { ['login', 'app'].forEach(n => $(n).hidden = n !== (name === 'login' ? 'login' : 'app')); }

  // switch-over day (§30): the token is the same for both sites. Mirror whichever key is set into the other, so a
  // user who signed in on the old portal is already signed in here, and one who signs in here stays signed in there.
  // Both keys are kept (not dropped) so moving back and forth between the two sites never forces a fresh sign-in.
  function carryOverSession() {
    try { const t = getSession(); if (t) setSession(t); } catch (e) {}
  }
  async function boot() {
    MX.pressAll(document);
    carryOverSession();
    try { const j = await call('me'); me = j; store.set('amp_email', me.email || ''); afterSignIn(); }
    catch (e) { show('login'); initLogin(); }
    registerSW();
  }
  // who may prepare reports — admin and Engg/CNL (operator), the same rule the server enforces (canGenerate_). Both get
  // the Today board with Generate; only the admin gets the Checks and Settings tabs (users, cut-off, upload key).
  const canGen = () => me && (me.role === 'admin' || me.role === 'operator');
  function afterSignIn() {
    show('app');
    const who = me.role === 'admin' ? 'Admin' : me.role === 'operator' ? 'Engg/CNL' : 'AEN/' + me.aen;
    $('whoami').textContent = (me.email || '') + ' · ' + who;
    $('signout').hidden = false;
    if (canGen()) adminBoard('today'); else aenHome();
  }

  /* ================= AEN home (§25.2) ================= */
  async function aenHome() {
    const root = $('screen'); root.innerHTML = ''; root.className = 'screen aen';
    $('menu').innerHTML = ''; $('menu').hidden = true;         // §25.2 fix: no admin nav for an AEN — no empty bar
    const head = el('div', 'aen-head');
    head.appendChild(el('div', 'aen-office', 'AEN/' + me.aen + (me.office ? ' · ' + town(me.office) : '')));
    const strip = el('div', 'date-strip'); strip.id = 'dateStrip'; head.appendChild(strip);
    root.appendChild(head);
    const panel = el('div', 'aen-report'); panel.id = 'aenReport'; root.appendChild(panel);
    panel.appendChild(skeletonReport());
    const cached = await readLast();                 // §25.4: the last report this user opened, shown at once
    if (cached && cached.date) openReader(cached, false, true);   // marked "loading newer…"; swapped below when fresh
    let today, hist;
    try { [today, hist] = await Promise.all([call('today'), call('history', { aen: me.aen }).catch(() => ({ snapshots: [] }))]); }
    catch (e) { panel.innerHTML = ''; panel.appendChild(el('div', 'note err', e.message)); return; }
    const mine = (today.aens || []).find(a => a.aen === me.aen) || {};
    const auto = today.auto || {};
    // date strip from history
    const dates = [...new Set((hist.snapshots || []).map(s => s.date))].slice(0, 30);
    buildDateStrip(strip, dates, today.date);
    renderAenDay(panel, today.date, mine, auto, hist);
  }
  function buildDateStrip(strip, dates, active) {
    strip.innerHTML = '';
    if (!dates.includes(active)) dates.unshift(active);
    dates.forEach(d => { const b = el('button', 'date-pill' + (d === active ? ' on' : '')); b.dataset.press = '1';
      b.innerHTML = '<span class="dp-wd">' + weekday(d) + '</span><span class="dp-d">' + d.split('-').reverse().slice(0, 2).join('.') + '</span>';
      b.onclick = () => openAenDate(d); strip.appendChild(b); });
    MX.pressAll(strip);
  }
  async function renderAenDay(panel, date, mine, auto, hist) {
    panel.innerHTML = '';
    if (!mine.latest) {
      const n = el('div', 'aen-wait');
      n.appendChild(el('div', 'big', 'Not ready yet'));
      n.appendChild(el('div', 'small', auto && auto.enabled && auto.next ? 'The next automatic check is at ' + auto.next + '.' : 'Ask the control room, or press Generate on a PC.'));
      panel.appendChild(n); return;
    }
    const bar = el('div', 'aen-bar');
    const madeTxt = mine.latest.auto ? 'Made automatically ' + mine.latest.auto + autoCheckTail(auto) : 'Ready since ' + mine.latest.at;
    bar.appendChild(el('div', 'aen-made', madeTxt));
    const acts = el('div', 'aen-acts');
    const open = el('button', 'btn sm', 'Open the report'); open.dataset.press = '1';
    const dl = el('button', 'btn-line sm', 'Download Word'); dl.dataset.press = '1';
    const pr = el('button', 'btn-line sm', 'Print'); pr.dataset.press = '1';
    const seen = el('button', 'btn-line sm', 'Seen'); seen.dataset.press = '1';
    acts.append(open, dl, pr, seen); bar.appendChild(acts); panel.appendChild(bar);
    const mount = el('div', 'aen-inline'); mount.id = 'aenInline'; panel.appendChild(mount);

    const report = await loadReport(me.aen, mine.latest.name);
    open.onclick = () => openReader(report, false);
    dl.onclick = async () => { const done = MX.busy(dl, 'Preparing…'); try { saveBlob(await docxBlob(report), docxName(report)); done('Saved ✓'); } catch (e) { done('Failed', 'err'); } };
    pr.onclick = () => openReader(report, true);
    seen.onclick = () => markSeen(seen, date);
    // show the gap summary behind the reader; §25.2: on landing, the report is open in the reader (Close returns here)
    mount.appendChild(inlineSummary(report));
    openReader(report, false);
  }
  function autoCheckTail(auto) { if (!auto || !auto.results) return ''; const done = (auto.slots || []).filter(s => auto.results[s] && auto.results[s].done); if (done.length <= 1) return ''; const last = done[done.length - 1]; return ' · checked ' + last + (changedAt(auto, last) ? '' : ', no change'); }
  function changedAt() { return false; }
  let readerAt = null;
  function openReader(report, printAfter, stale) {
    readerAt = report.aen + '|' + report.date + '|' + (report.generatedAt || '');
    Reader.open({ title: 'AEN/' + report.aen + ' · ' + dmy(report.date) + (stale ? ' · loading newer…' : ''), report, filename: docxName(report),
      blob: () => docxBlob(report),                         // the Download: the exact Word file (no display-only breaks)
      measure: () => AMPDocx.measureBlob(report),            // on screen: one flat page to measure the real heights
      display: (spec) => AMPDocx.displayBlob(report, spec),  // then one section per page at the measured breaks
      reportMistake: (section, text) => sendMistake(report.aen, report.date, section, text), onSeen: () => {} });
    if (printAfter) setTimeout(() => window.print(), 500);
  }
  async function sendMistake(aen, date, section, text) {
    return await call('reportMistake', { aen, date, section, text });   // the reader's sheet shows any error
  }
  async function openAenDate(d) {
    const h = await call('history', { aen: me.aen }).catch(() => ({ snapshots: [] }));
    const s = (h.snapshots || []).filter(x => x.date === d).sort((a, b) => a.name < b.name ? 1 : -1)[0];
    if (!s) return MX.toast('No report for ' + dmy(d));
    const report = await loadReport(me.aen, s.name);
    openReader(report, false);
  }
  // §25.2 fix: the "at a glance" line lists the actual gaps (section + the date used), not "see the report"
  function inlineSummary(R) {
    const box = el('div', 'summary');
    box.appendChild(el('div', 'summary-h', 'Today at a glance'));
    const gaps = gapsFromReport(R);
    if (!gaps.length) { box.appendChild(el('div', 'small ok', 'All sections on current data.')); return box; }
    box.appendChild(el('div', 'small', gaps.length + ' section' + (gaps.length > 1 ? 's' : '') + ' on older or missing data:'));
    const ul = el('ul', 'gap-list');
    gaps.forEach(g => { const li = el('li', 'gap-item' + (g.error ? ' err' : ' amber'));
      li.textContent = secTitle(g.section) + ' — ' + (g.error ? 'could not read' : 'older data' + (g.date ? ', ' + dmy(g.date) : ''));
      li.dataset.press = '1'; li.onclick = () => whyGap({ section: g.section, date: g.date, error: g.error, stale: g.stale, reason: g.reason }); ul.appendChild(li); });
    box.appendChild(ul); MX.pressAll(box);
    return box;
  }
  function gapsFromReport(R) {
    const S = R.sections || {}, out = [];
    Object.keys(S).forEach(k => { const s = S[k]; if (!s || !ST[k]) return;
      const stale = !!(s.stale || s.dayStale || s.nightStale || s.notAvailedStale || (s.work && s.work.stale) || (s.block && s.block.stale));
      if (s.ok === false) out.push({ section: k, error: s.error || 'could not read' });
      else if (stale) out.push({ section: k, stale: true, date: s.dataDate || null });
    });
    return out.sort((a, b) => (ST[a.section].no - ST[b.section].no));
  }

  /* ================= Admin board (§25.3) ================= */
  // Live generate state — the SAME server calls the old portal uses (prepare → generate ×N → finish). §19: a batch
  // reads every source once, builds five offices at a time from that store, and the shared lists are brought up to
  // date by finish; while it runs the Generate-all button is Stop; the counter is honest (done / running / waiting).
  let BT = null, BA = null, curTab = 'today';   // the board's today + auto, kept so a generate can repaint in place
  const busy = {}, failed = {}, fresh = {};     // aen -> 'run'|'queue' · aen -> error · aen -> built in this page
  let stopping = false, preparing = false, runCount = 0;

  async function adminBoard(tab) {
    const root = $('screen'); root.innerHTML = ''; root.className = 'screen admin';
    curTab = tab || 'today'; buildMenu(curTab);
    if (tab && tab !== 'today') return adminTab(tab);
    root.appendChild(skeletonBoard());
    let today, st;
    try { [today, st] = await Promise.all([call('today'), call('adminState').catch(() => ({}))]); }
    catch (e) { root.innerHTML = ''; root.appendChild(el('div', 'note err', e.message)); return; }
    paintBoard(root, today, st);
  }
  // build the whole Today board. The morning snapshot (timeline, sources) is static; the dynamic region below it
  // (roll-up, problems, office rows, Generate/Stop bar) is repainted in place by paintToday() as offices build.
  function paintBoard(root, today, st) {
    BT = today; BA = today.auto || {};
    root.innerHTML = '';
    root.appendChild(timeline(BA, st));
    root.appendChild(freshness(st));
    const dyn = el('div', 'board-dyn'); dyn.id = 'boardDyn'; root.appendChild(dyn);
    paintToday();
    MX.pressAll(root);
  }
  function paintToday() {
    const dyn = $('boardDyn'); if (!dyn || !BT) return;
    dyn.innerHTML = '';
    dyn.appendChild(rollup(BT));
    dyn.appendChild(causes(BT));
    dyn.appendChild(officeList(BT, BA));
    dyn.appendChild(genBar());
    MX.pressAll(dyn);
  }
  // after a batch finishes the shared lists changed; refetch today + adminState once and rebuild (no skeleton flash).
  async function refreshBoard() {
    if (curTab !== 'today') return;
    const root = $('screen'); if (!root || !root.classList.contains('admin')) return;
    let today, st;
    try { [today, st] = await Promise.all([call('today'), call('adminState').catch(() => ({}))]); } catch (e) { return; }
    if (!today || !Array.isArray(today.aens)) return;
    // keep this page's just-built reports if the server view lags; failed offices keep their red state via failed{}
    paintBoard(root, today, st);
  }
  function rollup(today) {
    const n = (today.aens || []); const ready = n.filter(a => a.latest).length; const gaps = n.filter(a => a.latest && (a.latest.gaps || []).length).length;
    const box = el('div', 'rollup');
    box.appendChild(stat(ready, 'ready')); box.appendChild(stat(gaps, 'with gaps')); box.appendChild(stat(10 - ready, 'to do'));
    return box;
  }
  function stat(n, lab) { const s = el('div', 'stat'); s.appendChild(el('div', 'stat-n num', String(n))); s.appendChild(el('div', 'stat-l small', lab)); return s; }
  function timeline(auto, st) {
    const box = el('div', 'tl card'); box.appendChild(el('div', 'tl-h', 'This morning'));
    const line = el('div', 'tl-line');
    const tms = (st && st.tmsControl && st.tmsControl.at) ? String(st.tmsControl.at).slice(11, 16) : null;
    const steps = [['07:30', 'cut-off', true], ['TMS', tms ? 'data ' + tms : 'data —', !!tms]];
    (auto.slots || ['08:30', '09:00', '09:30', '10:00']).forEach(sl => { const r = auto.results && auto.results[sl]; steps.push([sl, slotWords(r, auto.next === sl), r && r.done, r]); });
    steps.forEach(([t, s2, done, r]) => { const node = el('div', 'tl-step' + (done ? ' done' : '')); node.appendChild(el('div', 'tl-t mono', t)); node.appendChild(el('div', 'tl-s small', s2));
      if (r && r.failed && r.failed.length) node.title = r.failed.map(f => f.aen + ': ' + f.error).join('\n'); line.appendChild(node); });
    box.appendChild(line);
    if (!auto.enabled) box.appendChild(el('div', 'small faint', 'Automatic reports are off.'));
    return box;
  }
  // §24.2 slot result in words, not "1✓ 1×"
  function slotWords(r, isNext) {
    if (!r || !r.done) return isNext ? 'next' : '—';
    const parts = [];
    if (r.saved) parts.push(r.saved + ' saved');
    if (r.unchanged) parts.push(r.unchanged + ' unchanged');
    if (r.failed && r.failed.length) parts.push(r.failed.length + ' failed');
    return parts.length ? parts.join(' · ') : 'done';
  }
  // §25.3 freshness: one chip per source SHEET, tabs folded in, the oldest date shown, worst colour; full names on tap
  function freshness(st) {
    const box = el('div', 'fresh card');
    const h = el('div', 'fresh-head'); h.appendChild(el('div', 'tl-h', 'Sources'));
    const legend = el('div', 'fresh-legend small'); legend.innerHTML = '<span class="lg green">today</span><span class="lg amber">yesterday</span><span class="lg red">older</span>';
    h.appendChild(legend); box.appendChild(h);
    const strip = el('div', 'fresh-strip'); const tbl = (st && st.sourcesTable) || {};
    const sheets = {};               // sheet name -> { tabs:Set, oldest, latest }
    Object.keys(tbl).forEach(aen => (tbl[aen].rows || []).forEach(r => {
      const parts = String(r.source || r.id || '').split('›'); const sheet = parts[0].trim() || r.id; const tab = (parts[1] || '').trim();
      const g = sheets[sheet] || (sheets[sheet] = { sheet, tabs: new Set(), oldest: null, hasMissing: false });
      if (tab) g.tabs.add(tab);
      if (!r.date) g.hasMissing = true; else if (!g.oldest || r.date < g.oldest) g.oldest = r.date;
    }));
    const all = Object.values(sheets);
    const newest = all.map(g => g.oldest).filter(Boolean).sort().slice(-1)[0] || '';
    if (!all.length) strip.appendChild(el('div', 'small faint', 'No source dates yet.'));
    all.sort((a, b) => (a.oldest || '0') < (b.oldest || '0') ? -1 : 1).forEach(g => {
      const d = g.oldest, cls = g.hasMissing && !d ? 'red' : !d ? 'red' : d === newest ? 'green' : d >= addDays(newest, -1) ? 'amber' : 'red';
      const chip = el('button', 'fresh-chip ' + cls); chip.dataset.press = '1';
      const tabN = g.tabs.size; const label = tidySheet(g.sheet) + (tabN > 1 ? ' · ' + tabN + ' tabs' : '') + ' · ' + (d ? d.split('-').reverse().slice(0, 2).join('.') : '—');
      chip.textContent = label; chip.title = g.sheet + (tabN ? ' — tabs: ' + [...g.tabs].join(', ') : '') + (d ? ' — oldest ' + dmy(d) : ' — not read');
      chip.onclick = () => MX.toast(chip.title);
      strip.appendChild(chip);
    });
    box.appendChild(strip); return box;
  }
  function tidySheet(s) { return String(s).replace(/\bsheet\b/i, '').replace(/\bexport\b/i, '').replace(/\s{2,}/g, ' ').trim(); }
  // §25.3 problems grouped by cause — the server already normalised the office out (today.causes)
  function causes(today) {
    const box = el('div', 'causes card'); box.appendChild(el('div', 'tl-h', 'Problems grouped by cause'));
    const list = today.causes || [];
    if (!list.length) { box.appendChild(el('div', 'small ok', 'No problems across the offices.')); return box; }
    list.forEach(g => { const row = el('div', 'cause wide'); row.dataset.press = '1';
      row.appendChild(el('div', 'cause-t', secTitle(g.section) + ': ' + g.text + ' — affects ' + g.offices.length + ' office' + (g.offices.length > 1 ? 's' : '')));
      row.appendChild(el('div', 'cause-o small faint', g.offices.join(', ')));
      // §21: open the gap — Re-check the source and Generate again for the affected offices (Engg/CNL and admin)
      row.onclick = () => whyGap({ section: g.section, reason: g.text, stale: g.kind === 'stale', error: g.kind === 'error' ? g.text : null, offices: g.offices.slice() });
      box.appendChild(row); });
    return box;
  }
  function officeList(today, auto) {
    const box = el('div', 'offices');
    const autoOn = !!(auto && auto.enabled && auto.next);           // §30: only mention "next check" when auto is on
    (today.aens || []).forEach(a => {
      const row = el('div', 'office wide'); row.dataset.press = '1';
      const left = el('div', 'office-l'); left.appendChild(el('span', 'office-n', 'AEN/' + a.aen));
      const s = stateOfOffice(a);
      const pill = el('span', 'pill ' + s.pill, s.label); if (s.title) pill.title = s.title; left.appendChild(pill);
      const right = el('div', 'office-r small faint');
      right.textContent = a.latest ? (a.latest.auto ? 'auto ' + a.latest.auto : a.latest.at) + (autoOn ? ' · next ' + auto.next : '')
        : (autoOn ? 'next check ' + auto.next : 'not prepared yet');
      const acts = el('div', 'office-acts');
      if (a.latest) { const open = el('button', 'btn-line sm', 'Open'); open.dataset.press = '1'; open.onclick = ev => { ev.stopPropagation(); openAdminReport(a.aen, a.latest.name); }; acts.appendChild(open); }
      if (!busy[a.aen]) { const g = el('button', (a.latest ? 'btn-line' : 'btn') + ' sm', a.latest ? 'Generate again' : 'Generate'); g.dataset.press = '1';
        g.onclick = ev => { ev.stopPropagation(); runQueue([a.aen]); }; acts.appendChild(g); }
      const end = el('div', 'office-end'); end.append(right, acts);
      row.append(left, end);
      if (a.latest) row.onclick = () => openAdminReport(a.aen, a.latest.name);
      box.appendChild(row);
    });
    return box;
  }
  function stateOfOffice(a) {
    if (busy[a.aen] === 'run') return { pill: 'run', label: 'generating…' };
    if (busy[a.aen] === 'queue') return { pill: 'wait', label: 'queued' };
    if (failed[a.aen]) return { pill: 'failed', label: 'failed', title: failed[a.aen] };
    if (!a.latest) return { pill: 'to', label: 'to do' };
    const gl = (a.latest.gaps || []).length;
    return gl ? { pill: 'gaps', label: 'gaps ' + gl } : { pill: 'ready', label: 'ready' };
  }
  async function openAdminReport(aen, name) {
    const report = await loadReport(aen, name);
    openReader(report, false);
  }

  /* ---------- Generate (§19): the Generate-all / Stop bar and the honest counter ---------- */
  function genBar() {
    const wrap = el('div', 'admin-gen'), row = el('div', 'gen-row');
    const running = (BT.aens || []).filter(a => busy[a.aen] === 'run').length;
    const waiting = (BT.aens || []).filter(a => busy[a.aen] === 'queue').length;
    const active = preparing || running + waiting > 0;
    const btn = el('button', 'btn sm', preparing ? 'Reading sources…' : active ? 'Stop' : 'Generate all ten'); btn.dataset.press = '1';
    btn.disabled = stopping && running === 0;
    btn.onclick = () => { if (preparing || running + waiting > 0 || stopping) stopAll(); else runQueue((BT.aens || []).map(a => a.aen), true); };
    row.appendChild(btn);
    const done = Math.max(0, runCount - running - waiting);
    let txt;
    if (preparing) txt = 'Reading every source once…';
    else if (stopping) txt = 'Stopping · ' + running + ' still running';
    else if (active) txt = done + ' done · ' + running + ' running · ' + waiting + ' waiting';
    else { const left = (BT.aens || []).filter(a => !a.latest).length; txt = left ? left + ' to go' : 'All ready'; }
    row.appendChild(el('div', 'gen-progress small', txt));
    wrap.appendChild(row); return wrap;
  }
  function stopAll() { stopping = true; Object.keys(busy).forEach(a => { if (busy[a] === 'queue') delete busy[a]; }); paintToday(); }
  function runQueue(list, batch) { return runQueue_(list, batch).catch(e => { preparing = false; MX.toast(e.message); paintToday(); }); }
  async function runQueue_(list, batch) {
    const D = BT.date;
    list = list.filter(a => !busy[a]);
    if (!list.length) return;
    stopping = false; runCount = list.length;
    list.forEach(a => { busy[a] = 'queue'; delete failed[a]; });
    paintToday();
    const endRun = () => { stopping = false; preparing = false; runCount = 0; };
    let batchId = null;
    if (batch) {
      try {
        preparing = true; paintToday();
        const p = await prepareParts(D);
        preparing = false;
        if (stopping) { list.forEach(a => delete busy[a]); endRun(); paintToday(); return; }               // Stop while reading
        if (p.stopped) { list.forEach(a => delete busy[a]); endRun(); MX.toast(p.message || 'Reading the sources was stopped — Generate again'); paintToday(); return; }
        batchId = p.batchId;
      } catch (e) { list.forEach(a => delete busy[a]); endRun(); paintToday(); MX.toast(e.message); return; }
    }
    const one = async aen => {
      if (stopping) { delete busy[aen]; paintToday(); return; }
      busy[aen] = 'run'; paintToday();
      try {
        const j = await call('generate', batchId ? { date: D, aen, batch: batchId } : { date: D, aen });
        if (j.stopped) { failed[aen] = j.step || 'stopped at 5 minutes'; }                  // §19: a self-stopped build shows as failed
        else { const a = (BT.aens || []).find(x => x.aen === aen); if (a) { a.latest = j.info; a.count = (a.count || 0) + 1; fresh[aen] = true; } }
      } catch (e) { failed[aen] = e.message; }                                              // lock / timeout message surfaced on the tile
      delete busy[aen]; paintToday();
    };
    const todo = list.slice();
    const worker = async () => { while (todo.length && !stopping) await one(todo.shift()); };
    await Promise.all(Array.from({ length: batchId ? 5 : 3 }, worker));   // five offices at a time from the store
    // the offices wrote only their own files; the shared lists (index, Sources, repairs, admin notes) are brought up
    // to date once, here — also after Stop. An older server has no finish: nothing to do then.
    if (batchId) { try { await call('finish', { date: D }); } catch (e) {} }
    endRun();
    const nfail = Object.keys(failed).length;
    MX.toast(nfail ? nfail + ' office' + (nfail > 1 ? 's' : '') + ' did not build — see the red state' : (batch ? 'All offices ready' : 'Done'));
    await refreshBoard();                                                 // refresh causes / sources / timeline from the saved state
  }
  // §19: read the shared sources in four parts side by side; the work plans go first and alone (a server that knows
  // parts answers in a few seconds), the other three then start together under the same id. An older server ignores
  // the part, reads everything once and answers with its own id, used as it is.
  async function prepareParts(D) {
    const id = D + '_' + Math.random().toString(36).slice(2, 10).replace(/[^a-z0-9]/g, 'x').padEnd(8, 'x');
    const first = await call('prepare', { date: D, part: 'plans', batchId: id });
    if (first.stopped || first.part !== 'plans') return first;            // older server: a whole read already done
    const rest = await Promise.all(['meeting', 'pir', 'pwi'].map(part => call('prepare', { date: D, part, batchId: id })));
    const bad = rest.find(r => r.stopped);
    return bad || { batchId: id, date: D };
  }

  /* ---- admin tabs: Reports / Checks / Settings ---- */
  function buildMenu(active) {
    const m = $('menu'); m.innerHTML = ''; m.hidden = false;
    const tabs = me.role === 'admin' ? [['today', 'Today'], ['reports', 'Reports'], ['checks', 'Checks'], ['settings', 'Settings']]
                                     : [['today', 'Today'], ['reports', 'Reports']];     // Engg/CNL: no admin pages
    tabs.forEach(([k, lab]) => {
      const b = el('button', 'menu-b sm' + (k === active ? ' on' : '')); b.textContent = lab; b.dataset.press = '1'; b.onclick = () => adminBoard(k); m.appendChild(b);
    });
    MX.pressAll(m);
  }
  async function adminTab(tab) {
    const root = $('screen');
    if ((tab === 'settings' || tab === 'checks') && me.role !== 'admin') return adminBoard('today');   // admin-only
    if (tab === 'settings') return settingsTab(root);
    if (tab === 'checks') return checksTab(root);
    if (tab === 'reports') return reportsTab(root);
  }
  /* ---- Reports: per-office history, View + Download (admin and Engg/CNL) ---- */
  async function reportsTab(root) {
    const card = el('div', 'card'); card.appendChild(el('div', 'tl-h', 'Saved reports'));
    let aens = (BT && BT.aens) ? BT.aens.slice() : null;
    if (!aens) { try { const t = await call('today'); aens = t.aens; } catch (e) { root.appendChild(el('div', 'note err', e.message)); return; } }
    const sel = el('select', 'field'); aens.forEach(a => { const o = el('option', '', 'AEN/' + a.aen + (a.office ? ' · ' + town(a.office) : '')); o.value = a.aen; sel.appendChild(o); });
    const box = el('div', 'rep-list');
    const load = async () => { box.innerHTML = ''; box.appendChild(el('div', 'small faint', 'Reading…'));
      try { drawHistory(box, sel.value, await call('history', { aen: sel.value })); } catch (e) { box.innerHTML = ''; box.appendChild(el('div', 'small err', e.message)); } };
    sel.onchange = load; card.append(sel, box); root.appendChild(card); MX.pressAll(card); load();
  }
  function drawHistory(box, aen, h) {
    box.innerHTML = '';
    const rows = (h.snapshots || []).filter(s => !s.again);
    if (!rows.length) { box.appendChild(el('div', 'small faint', 'No reports yet.')); return; }
    rows.slice(0, 60).forEach(s => {
      const r = el('div', 'rep-row');
      const d = el('div', 'rep-d', dmy(s.date) + (s.times > 1 ? ' · ' + s.times + '×' : ''));
      const at = el('div', 'rep-at small faint', (s.at || '') + (s.auto ? ' · auto ' + s.auto : ''));
      const acts = el('div', 'office-acts');
      const v = el('button', 'btn-line sm', 'View'); v.dataset.press = '1'; v.onclick = async () => { const done = MX.busy(v, '…'); try { const R = await loadReport(aen, s.name); done(); openReader(R, false); } catch (e) { done('Failed', 'err'); MX.toast(e.message); } };
      const dl = el('button', 'btn sm', 'Download'); dl.dataset.press = '1'; dl.onclick = async () => { const done = MX.busy(dl, 'Preparing…'); try { const R = await loadReport(aen, s.name); saveBlob(await docxBlob(R), docxName(R)); done('Saved ✓'); } catch (e) { done('Failed', 'err'); MX.toast('Could not build the Word file: ' + (e.message || e)); } };
      acts.append(v, dl);
      const end = el('div', 'office-end'); end.append(at, acts);
      r.append(d, end); box.appendChild(r);
    });
    MX.pressAll(box);
  }

  /* ---- Settings (admin): auto, email, cut-off, upload key, users ---- */
  async function settingsTab(root) {
    let st; try { st = await call('adminState'); } catch (e) { root.appendChild(el('div', 'note err', e.message)); return; }
    const au = st.auto || { enabled: false, installed: false, slots: [], results: {} };
    const wrap = el('div');
    // Automatic reports + Email
    const c1 = el('div', 'card'); c1.appendChild(el('div', 'tl-h', 'Automatic reports'));
    const line = el('div', 'row');
    const tog = el('button', 'btn-line', au.enabled ? 'Turn off' : 'Turn on'); tog.dataset.press = '1';
    if (au.installed === false) {
      // §30: this release installs no run times, so there is nothing to switch on — say so and disable the toggle.
      tog.disabled = true; tog.classList.add('is-disabled');
      line.appendChild(tog); line.appendChild(el('span', 'small faint', 'Not set up in this release (no run times installed).'));
    } else {
      tog.onclick = async () => { const done = MX.busy(tog); try { const j = await call('setAuto', { value: au.enabled ? 'off' : 'on' }); done(); MX.toast('Automatic reports ' + (j.auto.enabled ? 'on' : 'off')); settingsTab2(root); } catch (e) { done(); MX.toast(e.message); } };
      line.appendChild(tog); line.appendChild(el('span', 'small', au.enabled ? 'On · checks at ' + (au.slots || []).join(', ') : 'Off'));
    }
    c1.appendChild(line);
    const email = el('div', 'row'); email.style.marginTop = '12px';
    const et = el('button', 'btn-line', 'Email when ready: ' + (st.emailOn ? 'on' : 'off')); et.dataset.press = '1';
    et.onclick = async () => { const done = MX.busy(et); try { const j = await call('setEmail', { value: st.emailOn ? 'off' : 'on' }); st.emailOn = j.emailOn; et.textContent = 'Email when ready: ' + (j.emailOn ? 'on' : 'off'); done(); MX.toast('Email ' + (j.emailOn ? 'on' : 'off')); } catch (e) { done('Failed', 'err'); MX.toast(e.message); } };
    email.appendChild(et); email.appendChild(el('span', 'small faint', 'Default off.')); c1.appendChild(email);
    wrap.appendChild(c1);
    // Cut-off
    const c2 = el('div', 'card'); c2.appendChild(el('div', 'tl-h', 'Cut-off time'));
    const crow = el('div', 'row');
    const cin = el('input', 'field'); cin.type = 'time'; cin.value = st.cutoff || ''; cin.style.maxWidth = '150px'; cin.style.marginBottom = '0';
    const csave = el('button', 'btn-line sm', 'Save'); csave.dataset.press = '1';
    csave.onclick = async () => { const done = MX.busy(csave); try { await call('setCutoff', { cutoff: cin.value }); done('Saved ✓'); MX.toast('Cut-off saved.'); } catch (e) { done('Failed', 'err'); MX.toast(e.message); } };
    crow.append(cin, csave); c2.append(crow, el('div', 'small faint', 'The day/night window ends at this time.')); wrap.appendChild(c2);
    // Upload key + received uploads (§27.1: the Engg/CNL downloader uploads the caution file, Engg-Control mail and TMS data with it)
    const c3 = el('div', 'card'); c3.appendChild(el('div', 'tl-h', 'Downloader upload key'));
    c3.appendChild(el('div', 'small', st.uploadKeySet ? 'A key is set. The Engg/CNL downloader uses it to upload the caution file, Engineering-Control mail files and TMS data.' : 'No key yet — make one and paste it into the downloader.'));
    const krow = el('div', 'row'); const kbtn = el('button', 'btn-line sm', 'Make a new key'); kbtn.dataset.press = '1';
    const kout = el('div', 'mono small'); kout.style.wordBreak = 'break-all'; kout.style.marginTop = '8px';
    kbtn.onclick = async () => { const done = MX.busy(kbtn); try { const j = await call('newUploadKey'); done('Made ✓'); kout.textContent = j.uploadKey; MX.toast('New key made — copy it now; it is shown once.'); } catch (e) { done('Failed', 'err'); MX.toast(e.message); } };
    krow.appendChild(kbtn); c3.append(krow, kout);
    const KIND = { items: 'TMS items', machines: 'TMS machine progress', stations: 'Station list', caution: 'Caution file', mail: 'Engg-Control mail' };
    const ups = (st.uploads || []);
    if (ups.length) { const tb = el('div', 'uploads'); tb.appendChild(el('div', 'small faint', 'Received'));
      ups.slice(0, 12).forEach(u => tb.appendChild(el('div', 'up-row small', u.at + ' · ' + (KIND[u.kind] || u.kind) + (u.date ? ' · ' + dmy(u.date) : '') + (u.rows != null ? ' · ' + u.rows + ' rows' : '')))); c3.appendChild(tb); }
    wrap.appendChild(c3);
    // Users
    wrap.appendChild(usersCard(st));
    root.appendChild(wrap); MX.pressAll(wrap);
  }
  function settingsTab2(root) { root.innerHTML = ''; buildMenu('settings'); settingsTab(root); }
  function usersCard(st) {
    const card = el('div', 'card'); card.appendChild(el('div', 'tl-h', 'Users'));
    const admins = st.admins || [], aens = st.aens || (BT ? (BT.aens || []).map(a => a.aen) : []);
    let users = (st.users || []).map(u => Object.assign({}, u));
    const list = el('div', 'users');
    const draw = () => {
      list.innerHTML = '';
      admins.forEach(a => { const r = el('div', 'user-row fixed'); r.append(el('div', 'small', a), el('div', 'small faint', 'Administrator')); list.appendChild(r); });
      users.forEach((u, i) => {
        const r = el('div', 'user-row');
        const em = el('input', 'field'); em.type = 'email'; em.value = u.email || ''; em.placeholder = 'email'; em.oninput = () => u.email = em.value.trim();
        const role = el('select', 'field'); [['aen', 'AEN (download)'], ['operator', 'Engg/CNL (prepare)']].forEach(([v, t]) => { const o = el('option', '', t); o.value = v; if ((u.role || 'aen') === v) o.selected = true; role.appendChild(o); });
        const aen = el('select', 'field'); const ch = el('option', '', 'Choose AEN'); ch.value = ''; aen.appendChild(ch);
        aens.forEach(c => { const o = el('option', '', 'AEN/' + c); o.value = c; if (c === u.aen) o.selected = true; aen.appendChild(o); });
        const nm = el('input', 'field'); nm.value = u.name || ''; nm.placeholder = 'name or post'; nm.oninput = () => u.name = nm.value;
        const del = el('button', 'btn-line sm', 'Remove'); del.dataset.press = '1'; del.onclick = () => { users.splice(i, 1); draw(); };
        const sync = () => { u.role = role.value; aen.disabled = role.value === 'operator'; if (aen.disabled) { aen.value = ''; u.aen = ''; } };
        role.onchange = sync; aen.onchange = () => u.aen = aen.value; sync();
        r.append(em, role, aen, nm, del); list.appendChild(r);
      });
      MX.pressAll(list);
    };
    draw();
    const acts = el('div', 'row'); acts.style.marginTop = '10px';
    const add = el('button', 'btn-line sm', 'Add user'); add.dataset.press = '1'; add.onclick = () => { users.push({ email: '', role: 'aen', aen: '', active: true }); draw(); };
    const save = el('button', 'btn sm', 'Save users'); save.dataset.press = '1';
    save.onclick = async () => { const done = MX.busy(save, 'Saving…'); try { const j = await call('saveUsers', { users: users.filter(u => u.email) }); users = (j.users || []).map(u => Object.assign({}, u)); draw(); done('Saved ✓'); MX.toast('Users saved.'); } catch (e) { done('Failed', 'err'); MX.toast(e.message); } };
    acts.append(add, save); card.append(list, acts);
    return card;
  }

  /* ---- Checks (admin): mistakes, column repairs (Keep/Undo), placements ---- */
  async function checksTab(root) {
    let st; try { st = await call('adminState'); } catch (e) { root.appendChild(el('div', 'note err', e.message)); return; }
    const wrap = el('div');
    const c1 = el('div', 'card'); c1.appendChild(el('div', 'tl-h', 'Mistakes reported'));
    try { const j = await call('mistakes'); const mlist = j.mistakes || [];
      if (!mlist.length) c1.appendChild(el('div', 'small ok', 'None reported.'));
      mlist.forEach(m2 => { const row = el('div', 'mistake'); row.append(el('div', '', m2.text), el('div', 'small faint', 'AEN/' + m2.aen + ' · ' + dmy(m2.date) + ' · ' + secTitle(m2.section) + ' · ' + (m2.by || ''))); c1.appendChild(row); });
    } catch (e) { c1.appendChild(el('div', 'small err', 'Could not read mistakes: ' + e.message)); }
    wrap.appendChild(c1);
    wrap.appendChild(repairsCard(st.repairs || {}));
    const c3 = el('div', 'card'); c3.append(el('div', 'tl-h', 'Placements'), el('div', 'small faint', 'Rows the sheet did not place under an AEN.'));
    const pbox = el('div'); c3.appendChild(pbox); wrap.appendChild(c3);
    loadPlacements(pbox);
    root.appendChild(wrap); MX.pressAll(wrap);
  }
  // §16: Keep stores the new heading as a known name (the note stops); Undo rejects it (flagged as a gap from then on)
  function repairsCard(rep) {
    const card = el('div', 'card'); card.appendChild(el('div', 'tl-h', 'Column changes the program made'));
    const tabs = Object.keys(rep).filter(t => (rep[t].list || []).length);
    if (!tabs.length) { card.appendChild(el('div', 'small ok', 'No column changes — every heading was read as known.')); return card; }
    tabs.forEach(t => {
      card.appendChild(el('div', 'small faint', rep[t].tab + (rep[t].at ? ' · ' + rep[t].at : '')));
      (rep[t].list || []).forEach(rp => {
        const row = el('div', 'repair-row');
        row.appendChild(el('div', '', '"' + rp.from + '" → "' + rp.to + '"' + (rp.via && rp.via !== 'alt' ? ' (by ' + rp.via + ')' : '')));
        const acts = el('div', 'office-acts');
        const keep = el('button', 'btn-line sm', 'Keep'); keep.dataset.press = '1'; keep.onclick = () => decideRepair('keepRepair', rp.tab || rep[t].tab, rp.col, rp.to, keep);
        const undo = el('button', 'btn-line sm', 'Undo'); undo.dataset.press = '1'; undo.onclick = () => decideRepair('undoRepair', rp.tab || rep[t].tab, rp.col, rp.to, undo);
        acts.append(keep, undo); row.appendChild(acts); card.appendChild(row);
      });
    });
    return card;
  }
  async function decideRepair(fn, tab, col, to, btn) {
    const done = MX.busy(btn);
    try { await call(fn, { tab, col, to }); done(fn === 'keepRepair' ? 'Kept ✓' : 'Undone ✓'); MX.toast(fn === 'keepRepair' ? 'Heading kept as known.' : 'Heading will be flagged again.'); }
    catch (e) { done('Failed', 'err'); MX.toast(e.message); }
  }
  async function loadPlacements(box) {
    box.innerHTML = ''; box.appendChild(el('div', 'small faint', 'Reading…'));
    let v; try { v = await call('placements'); } catch (e) { box.innerHTML = ''; box.appendChild(el('div', 'small err', e.message)); return; }
    drawPlacements(box, v);
  }
  function drawPlacements(box, v) {
    box.innerHTML = '';
    const open = v.open || [], answered = v.answered || [];
    if (!open.length) box.appendChild(el('div', 'small ok', 'Nothing waiting — every row is placed.'));
    open.forEach(o => {
      const row = el('div', 'place-row');
      row.append(el('div', '', o.where + (o.text ? ' · ' + o.text : '')), el('div', 'small faint', (o.days || []).slice(-3).map(dmy).join(', ')));
      const acts = el('div', 'office-acts');
      const sel = el('select', 'field'); sel.style.marginBottom = '0'; const ch = el('option', '', 'Belongs to…'); ch.value = ''; sel.appendChild(ch);
      (v.aens || []).forEach(a => { const op = el('option', '', 'AEN/' + a); op.value = a; sel.appendChild(op); });
      sel.onchange = () => { if (sel.value) answerPlacement(box, o.where, o.text, sel.value); };
      const ig = el('button', 'btn-line sm', 'Ignore'); ig.dataset.press = '1'; ig.onclick = () => answerPlacement(box, o.where, o.text, 'ignore');
      const sw = el('button', 'btn-line sm', 'Date swap'); sw.dataset.press = '1'; sw.onclick = () => answerPlacement(box, o.where, o.text, 'swap');
      acts.append(sel, ig, sw); row.appendChild(acts); box.appendChild(row);
    });
    if (answered.length) { box.appendChild(el('div', 'small faint', 'Answered'));
      answered.forEach(o => { const row = el('div', 'place-row');
        row.appendChild(el('div', 'small', o.where + (o.text ? ' · ' + o.text : '') + ' → ' + (o.aen ? 'AEN/' + o.aen : o.ignore ? 'ignore' : 'date swap')));
        const u = el('button', 'btn-line sm', 'Undo'); u.dataset.press = '1'; u.onclick = () => answerPlacement(box, o.where, o.text, null);
        row.appendChild(u); box.appendChild(row); }); }
    MX.pressAll(box);
  }
  async function answerPlacement(box, where, text, answer) {
    try { const v = await call('savePlacement', { where, text, answer }); drawPlacements(box, v); MX.toast(answer == null ? 'Answer removed.' : 'Saved — used in every report from now on.'); }
    catch (e) { MX.toast(e.message); }
  }

  /* ---- Seen (§25.4) ---- */
  async function markSeen(btn, date) {
    const done = MX.busy(btn, 'Seen ✓');
    try { await call('seen', { aen: me.aen, date }); done('Seen ✓'); }
    catch (e) { done('Failed', 'err'); MX.toast(e.message); }
  }
  function whyGap(gap) {
    const node = el('div', 'why');
    node.appendChild(el('div', 'why-h', secTitle(gap.section)));
    node.appendChild(el('div', '', gap.reason || gap.error || (gap.stale ? 'Older data was used because the sheet is behind the date.' : 'See the Sources tab.')));
    if (gap.date) node.appendChild(el('div', 'small faint', 'Data date: ' + dmy(gap.date)));
    if (gap.offices && gap.offices.length) node.appendChild(el('div', 'small faint', 'Affects AEN/' + gap.offices.join(', AEN/')));
    // §21: Re-check re-reads the source now; the server puts the fresh read in place of the old one. Only Engg/CNL and
    // admin may re-check, the same rule the old portal applies (an AEN user downloads, it does not re-check).
    const aen = gap.aen || (gap.offices && gap.offices[0]);
    if (canGen() && gap.section && aen) {
      const on = gap.on || (BT && BT.date);
      const said = el('div', 'small'); said.style.marginTop = '8px';
      const rc = el('button', 'btn-line sm', 'Re-check the source'); rc.dataset.press = '1';
      rc.onclick = async () => { const done = MX.busy(rc, 'Re-reading…');
        try { const v = await call('recheck', { aen, date: on, section: gap.section }); done();
          said.className = 'small ' + (v.fixed ? 'ok' : '');
          said.textContent = v.fixed
            ? 'Fixed now' + (v.dataDate ? ': found data for ' + dmy(v.dataDate) : '') + ' — Generate again to put it in the report'
            : 'Still the same' + (v.found ? ': ' + v.found : v.error ? ': ' + v.error : v.dataDate ? ': newest ' + dmy(v.dataDate) : '') + (v.at ? ' (' + String(v.at).slice(11, 16) + ')' : '');
        } catch (e) { done('Failed', 'err'); said.className = 'small err'; said.textContent = e.message; } };
      const gen = el('button', 'btn sm', 'Generate again'); gen.dataset.press = '1';
      gen.onclick = () => { runQueue((gap.offices && gap.offices.length) ? gap.offices : [aen]); };
      const acts = el('div', 'why-acts'); acts.style.cssText = 'display:flex;gap:8px;margin-top:10px;flex-wrap:wrap'; acts.append(rc, gen);
      node.append(acts, said);
      setTimeout(() => MX.pressAll(node), 0);
    }
    MX.sheet(node);
  }

  /* ---------- report caches ---------- */
  const rcache = {};
  async function loadReport(aen, name) { const k = aen + '/' + name; if (rcache[k]) return rcache[k]; const j = await call('snapshot', { aen, name }); rcache[k] = j.report; cacheLast(j.report); return j.report; }
  const saveBlob = (blob, name) => AMPDocx.saveBlob(blob, name);

  /* ---------- skeletons ---------- */
  function skeletonReport() { const w = el('div', 'skelwrap'); for (let i = 0; i < 5; i++) w.appendChild(MX.skeleton('line')); return w; }
  function skeletonBoard() { const w = el('div', 'skelwrap'); for (let i = 0; i < 10; i++) w.appendChild(MX.skeleton('bar')); return w; }

  /* ---------- login ---------- */
  function initLogin() {
    $('loginEmail').value = store.get('amp_email') || '';
    $('sendCode').onclick = sendCode; $('verify').onclick = verify;
    $('loginEmail').onkeydown = e => { if (e.key === 'Enter') sendCode(); };
    $('loginCode').onkeydown = e => { if (e.key === 'Enter') verify(); };
    MX.pressAll($('login'));
  }
  async function sendCode() {
    const email = $('loginEmail').value.trim(); if (!email) return msg('msgLogin', 'Enter your email.', 'err');
    const done = MX.busy($('sendCode'), 'Sending…');
    try { const j = await call('otpRequest', { email }); store.set('amp_email', email); $('stepEmail').hidden = true; $('stepCode').hidden = false; msg('msgLogin', 'Code sent to ' + (j.sentTo || email), 'ok'); done(); $('loginCode').focus(); }
    catch (e) { done(); msg('msgLogin', e.message, 'err'); }
  }
  async function verify() {
    const code = $('loginCode').value.trim(); const email = $('loginEmail').value.trim();
    const done = MX.busy($('verify'), 'Checking…');
    try { const j = await call('otpVerify', { email, code }); setSession(j.session); done(); me = j.user || null; const m = await call('me'); me = m; afterSignIn(); }
    catch (e) { done(); msg('msgLogin', e.message, 'err'); }
  }
  function msg(id, t, k) { const n = $(id); n.textContent = t; n.className = 'msg' + (k ? ' ' + k : ''); }
  $('signout') && ($('signout').onclick = () => { call('signOut').catch(() => {}); delSession(); clearCaches(); location.reload(); });

  /* ---------- PWA ---------- */
  function registerSW() { if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {}); }
  function lastKey() { return 'last-' + encodeURIComponent((me && me.email) || 'x'); }   // per signed-in user
  function cacheLast(R) { try { if ('caches' in window && me) caches.open('amp-report').then(c => c.put(lastKey(), new Response(JSON.stringify({ aen: R.aen, name: R.date + '_' + (R.generatedAt || '').slice(11).replace(/:/g,''), date: R.date, report: R }), { headers: { 'Content-Type': 'application/json' } }))); } catch (e) {} }
  async function readLast() { try { if (!('caches' in window) || !me) return null; const c = await caches.open('amp-report'); const r = await c.match(lastKey()); if (!r) return null; const j = await r.json(); return j && j.report && j.report.aen === me.aen ? j.report : null; } catch (e) { return null; } }
  function clearCaches() { try { if ('caches' in window) caches.keys().then(ks => ks.forEach(k => caches.delete(k))); if (navigator.serviceWorker) navigator.serviceWorker.getRegistrations().then(rs => rs.forEach(r => r.active && r.active.postMessage('clear'))); } catch (e) {} }

  window.addEventListener('DOMContentLoaded', boot);
  window.__AMP = { call, get me() { return me; } };   // for the screenshot harness
})();
