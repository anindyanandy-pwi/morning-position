/* AEN Morning Position — website logic. */
(function () {
  const C = window.AMP_CONFIG || {};
  const $ = id => document.getElementById(id);
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const store = { get: k => { try { return localStorage.getItem(k); } catch (e) { return null; } },
                  set: (k, v) => { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch (e) {} } };

  let session = store.get('amp_session'), me = null;

  /* ---------- server ---------- */
  async function call(fn, args) {
    let r;
    try {
      r = await fetch(C.API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ session, fn, args: args || {} }) });
    } catch (e) { throw new Error('The server could not be reached. Check the internet connection and try again.'); }
    if (!r.ok) throw new Error('The server did not answer (' + r.status + ').');
    const j = await r.json();
    if (!j.ok) {
      if (/sign in again|Please sign in|no longer registered|expired\. Please/.test(j.error || '') && fn !== 'otpVerify') signedOut(j.error);
      throw new Error(j.error || 'Server error');
    }
    return j;
  }
  const say = (node, text, kind) => { node.textContent = text || ''; node.className = 'msg' + (kind ? ' ' + kind : ''); };
  const tell = (text, kind) => say($('msg'), text, kind);

  /* ---------- dates and names ---------- */
  const dmy = iso => iso.split('-').reverse().join('.');
  const dotted = iso => iso.split('-').reverse().join('·');
  const addDays = (iso, n) => { const d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const weekday = iso => new Date(iso + 'T00:00:00Z').toLocaleDateString('en-IN', { weekday: 'long', timeZone: 'UTC' });
  const town = office => String(office || '').replace(/^Assistant Divisional Engineer[\s,/]*/i, '').replace(/\s*\(.*\)\s*$/, '').replace(/^Track\/(.*)$/, '$1 (Track)');
  const two = n => String(n).padStart(2, '0');
  const SECTION = { miniControl: 'Mini control', gps: 'Winter patrolling', major: 'Major failures', minor: 'Minor failures', tdcLapse: 'TDC lapse',
    pqrsTdc: 'PQRS TDC fail', protection: 'Protection caution', badRoad: 'Bad road surface', cctv: 'CCTV', kpi: 'KPI and inspections',
    blocks: 'Blocks', machines: 'Machines', tms: 'TMS items', file: 'Saved file' };
  const gapText = g => (SECTION[g.section] || g.section) + ': ' + (g.kind === 'stale' ? 'older data used' : g.text);

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
  const VIEWS = ['vToday', 'vAen', 'vReports', 'vUsers', 'vSettings', 'vData'];
  const MENU = { admin: [['vToday', 'Today'], ['vReports', 'Reports'], ['vUsers', 'Users'], ['vSettings', 'Settings'], ['vData', 'Data and checks']],
                 operator: [['vToday', 'Today'], ['vReports', 'Reports']] };
  const LOAD = { vToday: () => loadToday(), vAen: () => loadAen(), vReports: () => loadReports(), vUsers: () => loadAdmin(), vSettings: () => loadAdmin(), vData: () => loadAdmin() };

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
  async function saveDocx(R) {
    const blob = await docx.Packer.toBlob(buildMorningDoc(docx, R, { logo: await getLogo() }));
    const a = el('a'); a.href = URL.createObjectURL(blob);
    a.download = 'Morning_Position_AEN-' + R.aen.replace('/', '-') + '_' + dmy(R.date) + '.docx';
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 30000);
  }
  async function download(aen, name, btn) {
    if (btn) btn.disabled = true;
    try { const s = await call('snapshot', { aen, name }); await saveDocx(s.report); }
    catch (e) { tell(e.message, 'err'); }
    if (btn) btn.disabled = false;
  }

  /* ---------- today (Engg/CNL and admin) ---------- */
  let D = null, today = null;
  const busy = {};          // aen -> 'run' | 'queue'
  const failed = {};        // aen -> message
  $('tPick').onchange = () => { if ($('tPick').value) { D = $('tPick').value; loadToday(); } };

  async function loadToday() {
    drawHero();
    if (!today || today.date !== D) {
      today = null; $('tiles').innerHTML = ''; $('tProgress').textContent = 'READING…';
      $('genAll').disabled = true; $('genAll').textContent = 'Reading the day…';
    }
    try { const t = await call('today', { date: D }); if (t.date === D) { today = t; drawToday(); } }
    catch (e) { tell(e.message, 'err'); console.error(e); $('tProgress').textContent = ''; $('genAll').textContent = 'Generate all ten'; $('genAll').disabled = !today; }
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
    const g = a.latest.gaps || [];
    return g.length ? { cls: 'gap', st: 'Ready, ' + g.length + (g.length > 1 ? ' gaps' : ' gap'), gaps: g.map(gapText) } : { cls: 'ok', st: 'Ready' };
  }
  function drawToday() {
    drawHero();
    let ready = 0, gap = 0;
    if (!today || !Array.isArray(today.aens)) throw new Error('The server sent an incomplete answer. Please reload the page.');
    $('tiles').innerHTML = '';
    today.aens.forEach((a, i) => {
      const s = stateOf(a);
      if (s.cls === 'ok') ready++; if (s.cls === 'gap') gap++;
      const t = el('div', 'tile' + (s.cls === 'gap' || s.cls === 'bad' ? ' gap' : s.cls === 'run' ? ' run' : ''));
      const meta = el('div', 'meta'); meta.append(el('span', '', two(i + 1)), el('span', '', a.latest ? a.latest.at + (a.count > 1 ? ' · ×' + a.count : '') : '—'));
      t.append(meta, el('div', 'code', a.aen), el('div', 'town', town(a.office)), el('div', 'st ' + s.cls, s.st));
      const gp = el('div', 'gaps'); (s.gaps || []).forEach((g, k) => { if (k) gp.appendChild(el('br')); gp.appendChild(document.createTextNode(g)); }); t.appendChild(gp);
      const acts = el('div', 'acts');
      if (!busy[a.aen]) {
        if (a.latest) {
          const dl = el('button', 'btn', 'Download'); dl.onclick = () => download(a.aen, a.latest.name, dl);
          const again = el('button', 'btn-text', 'Prepare again'); again.onclick = () => runQueue([a.aen]);
          acts.append(dl, again);
        } else { const g = el('button', 'btn-line', 'Generate'); g.onclick = () => runQueue([a.aen]); acts.appendChild(g); }
      }
      t.appendChild(acts); $('tiles').appendChild(t);
    });
    const left = today.aens.length - ready - gap;
    $('cReady').textContent = two(ready); $('cGap').textContent = two(gap); $('cLeft').textContent = two(left);
    const running = today.aens.find(a => busy[a.aen] === 'run');
    $('tProgress').textContent = running ? 'AEN/' + running.aen + ' IN HAND' : (left ? left + ' TO GO' : 'ALL READY');
    const remaining = today.aens.filter(a => !a.latest).length;
    $('genAll').textContent = remaining === today.aens.length || remaining === 0 ? 'Generate all ten' : 'Generate remaining ' + remaining;
    $('genAll').disabled = Object.keys(busy).length > 0;
    drawSources(today.sources);
  }
  function drawSources(s) {
    const box = $('sources'); box.innerHTML = '';
    const line = (k, v) => { const r = el('div', 'kv'); r.append(el('span', '', k), el('span', 'mono', v)); box.appendChild(r); };
    const when = x => x ? (x.at.slice(0, 10) === D ? x.at.slice(11) : dmy(x.at.slice(0, 10)) + ' ' + x.at.slice(11)) : 'not yet';
    line('Meeting sheet', 'live');
    line('TMS items', when(s.items));
    line('TMS machines', when(s.machines));
    line('TMS Control', s.tmsControl.at ? when({ at: String(s.tmsControl.at).replace('T', ' ').slice(0, 16) }) : 'not found');
  }
  $('genAll').onclick = () => {
    if (!today) return;                      // the day is still being read; the button is disabled meanwhile
    const remaining = today.aens.filter(a => !a.latest).map(a => a.aen);
    runQueue(remaining.length ? remaining : today.aens.map(a => a.aen));
  };
  function runQueue(list) { return runQueue_(list).catch(e => { tell(e.message, 'err'); console.error(e); }); }
  async function runQueue_(list) {
    if (!today) return;
    list = list.filter(a => !busy[a]);
    list.forEach(a => { busy[a] = 'queue'; delete failed[a]; });
    drawToday();
    for (const aen of list) {
      busy[aen] = 'run'; drawToday();
      try {
        const j = await call('generate', { date: D, aen });
        const a = today.aens.find(x => x.aen === aen);
        if (a) { a.latest = j.info; a.count = (a.count || 0) + 1; }
      } catch (e) { failed[aen] = e.message; }
      delete busy[aen]; drawToday();
    }
    try { today = await call('today', { date: D }); drawToday(); } catch (e) {}
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
      Object.keys(s.issues || {}).sort().reverse().slice(0, 7).forEach(d => Object.keys(s.issues[d]).forEach(a => {
        const it = s.issues[d][a], lines = [...(it.unassigned || []).map(x => 'Not placed: ' + x), ...(it.timeFixes || []).map(x => 'Time read from a date: ' + x)];
        if (!lines.length) return;
        iss.appendChild(el('b', '', dmy(d) + ' · AEN/' + a));
        const ul = el('ul'); lines.forEach(l => ul.appendChild(el('li', '', l))); iss.appendChild(ul);
      }));
      if (!iss.children.length) iss.appendChild(el('p', 'empty', 'Nothing to look at.'));
    } catch (e) { tell(e.message, 'err'); console.error(e); }
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
