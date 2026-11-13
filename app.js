/* §25 preview app. One site; the screen changes by role, enforced on the server. Talks to the live server through
 * AMP_CONFIG.API_URL. New routes (seen / reportMistake / mistakes / email) are feature-detected: where the live
 * server does not have them yet, the control says so instead of failing. */
(function () {
  const $ = id => document.getElementById(id);
  const el = (t, c, x) => { const n = document.createElement(t); if (c) n.className = c; if (x != null) n.textContent = x; return n; };
  const store = { get: k => { try { return localStorage.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (e) {} }, del: k => { try { localStorage.removeItem(k); } catch (e) {} } };
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

  async function call(fn, args) {
    const r = await fetch(AMP_CONFIG.API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ session: store.get('amp_sess') || '', fn, args: args || {} }) });
    let j; try { j = await r.json(); } catch (e) { throw new Error('The server took too long — try again.'); }
    if (!j.ok) { if (/unknown fn|not allowed|no such/i.test(j.error || '')) missing[fn] = true;
      if (/sign in again|Please sign in|no longer registered|expired/i.test(j.error || '')) signedOut(); throw new Error(j.error || 'Server error'); }
    return j;
  }
  function signedOut() { store.del('amp_sess'); me = null; clearCaches(); show('login'); }

  /* ---------- shared report blob (same as Download) ---------- */
  // one builder for both sites (web/docx-build.js): logo + cover artwork, Word-2007 safe
  const docxBlob = R => AMPDocx.docxBlob(R);
  const docxName = R => AMPDocx.docxName(R);

  /* ---------- views ---------- */
  function show(name) { ['login', 'app'].forEach(n => $(n).hidden = n !== (name === 'login' ? 'login' : 'app')); }

  // switch-over day (§30): the old portal stored the session under 'amp_session'. Carry it over to the new key once,
  // so nobody has to sign in again — the token is the same and the server accepts it. Then drop the old key.
  function carryOverSession() {
    try { if (!store.get('amp_sess')) { const old = localStorage.getItem('amp_session'); if (old) store.set('amp_sess', old); } } catch (e) {}   // 13.11: keep amp_session so /classic/ stays signed in
  }
  async function boot() {
    MX.pressAll(document);
    carryOverSession();
    try { const j = await call('me'); me = j; store.set('amp_email', me.email || ''); afterSignIn(); }
    catch (e) { show('login'); initLogin(); }
    registerSW();
  }
  function afterSignIn() {
    show('app');
    $('whoami').textContent = (me.email || '') + ' · ' + (me.role === 'admin' ? 'Admin' : 'AEN/' + me.aen);
    $('signout').hidden = false;
    if (me.role === 'admin') adminBoard('today'); else aenHome();
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
    try { return await call('reportMistake', { aen, date, section, text }); }
    catch (e) { missing.reportMistake = true; throw new Error('Reporting mistakes is not on the live server yet (preview).'); }
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
  async function adminBoard(tab) {
    const root = $('screen'); root.innerHTML = ''; root.className = 'screen admin';
    buildMenu(tab || 'today');
    if (tab && tab !== 'today') return adminTab(tab);
    root.appendChild(skeletonBoard());
    let today, st;
    try { [today, st] = await Promise.all([call('today'), call('adminState').catch(() => ({}))]); }
    catch (e) { root.innerHTML = ''; root.appendChild(el('div', 'note err', e.message)); return; }
    root.innerHTML = '';
    const auto = today.auto || {};
    root.appendChild(rollup(today));
    root.appendChild(timeline(auto, st));
    root.appendChild(freshness(st));
    root.appendChild(causes(today));
    root.appendChild(officeList(today, auto));
    const gen = el('div', 'admin-gen');
    const g1 = el('button', 'btn-line sm', 'Generate all ten (classic portal)'); g1.dataset.press = '1'; g1.onclick = () => { location.href = 'classic/'; };   // 13.11: generate from /classic/ until wired
    gen.appendChild(g1); root.appendChild(gen);
    MX.pressAll(root);
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
      box.appendChild(row); });
    return box;
  }
  function officeList(today, auto) {
    const box = el('div', 'offices');
    (today.aens || []).forEach((a, i) => {
      const row = el('button', 'office wide'); row.dataset.press = '1';
      const left = el('div', 'office-l'); left.appendChild(el('span', 'office-n', 'AEN/' + a.aen));
      const stt = a.latest ? ((a.latest.gaps || []).length ? 'gaps' : 'ready') : 'to do';
      left.appendChild(el('span', 'pill ' + stt, stt + ((a.latest && a.latest.gaps || []).length ? ' ' + a.latest.gaps.length : '')));
      const right = el('div', 'office-r small faint');
      right.textContent = a.latest ? (a.latest.auto ? 'auto ' + a.latest.auto : a.latest.at) + (auto && auto.next ? ' · next ' + auto.next : '') : (auto && auto.next ? 'next check ' + auto.next : '');
      row.append(left, right);
      row.onclick = () => a.latest ? openAdminReport(a.aen, a.latest.name) : MX.toast('AEN/' + a.aen + ' has no report yet.');
      box.appendChild(row);
    });
    return box;
  }
  async function openAdminReport(aen, name) {
    const report = await loadReport(aen, name);
    openReader(report, false);
  }

  /* ---- admin tabs: Reports / Checks / Settings ---- */
  function buildMenu(active) {
    const m = $('menu'); m.innerHTML = ''; m.hidden = false;
    [['today', 'Today'], ['reports', 'Reports'], ['checks', 'Checks'], ['settings', 'Settings']].forEach(([k, lab]) => {
      const b = el('button', 'menu-b sm' + (k === active ? ' on' : '')); b.textContent = lab; b.dataset.press = '1'; b.onclick = () => adminBoard(k); m.appendChild(b);
    });
    MX.pressAll(m);
  }
  async function adminTab(tab) {
    const root = $('screen');
    if (tab === 'settings') return settingsTab(root);
    if (tab === 'checks') return checksTab(root);
    if (tab === 'reports') { root.appendChild(el('div', 'note', 'Reports: pick an office on Today to open its history.')); return; }
  }
  async function settingsTab(root) {
    const card = el('div', 'card'); card.appendChild(el('div', 'tl-h', 'Automatic reports'));
    let st; try { st = await call('adminState'); } catch (e) { root.appendChild(el('div', 'note err', e.message)); return; }
    const au = st.auto || { enabled: false, installed: false, slots: [], results: {} };
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
    card.appendChild(line);
    const email = el('div', 'row'); email.style.marginTop = '12px';
    const et = el('button', 'btn-line', 'Email when ready: ' + (st.emailOn ? 'on' : 'off')); et.dataset.press = '1';
    et.onclick = async () => { if (missing.setEmail) return MX.toast('Email is not on the live server yet (preview).'); const done = MX.busy(et); try { const j = await call('setEmail', { value: st.emailOn ? 'off' : 'on' }); done(); MX.toast('Email ' + (j.emailOn ? 'on' : 'off')); } catch (e) { done(); missing.setEmail = true; MX.toast('Email is not on the live server yet (preview).'); } };
    email.appendChild(et); email.appendChild(el('span', 'small faint', 'Default off.')); card.appendChild(email);
    root.appendChild(card);
  }
  function settingsTab2(root) { root.innerHTML = ''; buildMenu('settings'); settingsTab(root); }
  async function checksTab(root) {
    const card = el('div', 'card'); card.appendChild(el('div', 'tl-h', 'Mistakes reported'));
    try { const j = await call('mistakes'); const list = j.mistakes || [];
      if (!list.length) card.appendChild(el('div', 'small ok', 'None reported.'));
      list.forEach(m2 => { const row = el('div', 'mistake'); row.appendChild(el('div', '', m2.text)); row.appendChild(el('div', 'small faint', 'AEN/' + m2.aen + ' · ' + dmy(m2.date) + ' · ' + secTitle(m2.section) + ' · ' + (m2.by || ''))); card.appendChild(row); });
    } catch (e) { missing.mistakes = true; card.appendChild(el('div', 'small faint', 'Reporting mistakes is not on the live server yet (preview).')); }
    root.appendChild(card);
  }

  /* ---- Seen (§25.4) ---- */
  async function markSeen(btn, date) {
    const done = MX.busy(btn, 'Seen ✓');
    try { await call('seen', { aen: me.aen, date }); done('Seen ✓'); }
    catch (e) { missing.seen = true; done('Noted'); MX.toast('Seen is not on the live server yet (preview).'); }
  }
  function whyGap(gap) {
    const node = el('div', 'why');
    node.appendChild(el('div', 'why-h', secTitle(gap.section)));
    node.appendChild(el('div', '', gap.reason || gap.error || (gap.stale ? 'Older data was used because the sheet is behind the date.' : 'See the Sources tab.')));
    if (gap.date) node.appendChild(el('div', 'small faint', 'Data date: ' + dmy(gap.date)));
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
    try { const j = await call('otpVerify', { email, code }); store.set('amp_sess', j.session); done(); me = j.user || null; const m = await call('me'); me = m; afterSignIn(); }
    catch (e) { done(); msg('msgLogin', e.message, 'err'); }
  }
  function msg(id, t, k) { const n = $(id); n.textContent = t; n.className = 'msg' + (k ? ' ' + k : ''); }
  $('signout') && ($('signout').onclick = () => { call('signOut').catch(() => {}); store.del('amp_sess'); clearCaches(); location.reload(); });

  /* ---------- PWA ---------- */
  function registerSW() { if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {}); }
  function lastKey() { return 'last-' + encodeURIComponent((me && me.email) || 'x'); }   // per signed-in user
  function cacheLast(R) { try { if ('caches' in window && me) caches.open('amp-report').then(c => c.put(lastKey(), new Response(JSON.stringify({ aen: R.aen, name: R.date + '_' + (R.generatedAt || '').slice(11).replace(/:/g,''), date: R.date, report: R }), { headers: { 'Content-Type': 'application/json' } }))); } catch (e) {} }
  async function readLast() { try { if (!('caches' in window) || !me) return null; const c = await caches.open('amp-report'); const r = await c.match(lastKey()); if (!r) return null; const j = await r.json(); return j && j.report && j.report.aen === me.aen ? j.report : null; } catch (e) { return null; } }
  function clearCaches() { try { if ('caches' in window) caches.keys().then(ks => ks.forEach(k => caches.delete(k))); if (navigator.serviceWorker) navigator.serviceWorker.getRegistrations().then(rs => rs.forEach(r => r.active && r.active.postMessage('clear'))); } catch (e) {} }

  window.addEventListener('DOMContentLoaded', boot);
  window.__AMP = { call, get me() { return me; } };   // for the screenshot harness
})();
