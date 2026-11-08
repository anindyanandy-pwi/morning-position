/* §25.1 (amended 08.11) full-screen report reader — the report only: the exact .docx via docx-preview, the same
 * blob as Download. No Data mode. Zoom Fit/100/150, pinch, double-tap, swipe/arrow pages, rubber-band at the ends.
 * A "Report a mistake" button picks a section from the report and takes one line of text (§25.4). window.Reader.open(ctx).
 * ctx: { title, report, blob()->Promise<Blob>, filename, reportMistake(section,text)->Promise, onSeen }. */
(function () {
  const el = (t, c, x) => { const n = document.createElement(t); if (c) n.className = c; if (x != null) n.textContent = x; return n; };
  let host = null, state = null;

  function open(ctx) {
    close();
    state = Object.assign({ zoom: 'fit', scale: 1 }, ctx);
    host = el('div', 'rd'); host.setAttribute('role', 'dialog'); host.setAttribute('aria-label', 'Report');
    const head = el('div', 'rd-head');
    const title = el('div', 'rd-title', ctx.title || 'Report');
    const acts = el('div', 'rd-acts');
    // wide screens: the actions inline; narrow (<=600px): Download + ⋯ (Report a mistake, Print) + Close — never clipped
    const flag = el('button', 'btn-line sm rd-wide-only', 'Report a mistake'); flag.dataset.press = '1'; flag.onclick = () => mistakeSheet();
    const dl = el('button', 'btn sm', 'Download'); dl.dataset.press = '1'; dl.onclick = () => doDownload(dl);
    const pr = el('button', 'btn-line sm rd-wide-only', 'Print'); pr.dataset.press = '1'; pr.onclick = () => window.print();
    const more = el('button', 'btn-line sm rd-narrow-only', '⋯'); more.setAttribute('aria-label', 'More'); more.dataset.press = '1'; more.onclick = () => moreSheet();
    const cl = el('button', 'btn-line sm rd-close', 'Close'); cl.dataset.press = '1'; cl.onclick = close;
    acts.append(flag, pr, dl, more, cl);
    head.append(title, acts);
    const body = el('div', 'rd-body rd-page'); body.id = 'rdBody';
    host.append(head, body); document.body.appendChild(host); document.body.style.overflow = 'hidden'; document.body.classList.add('rd-open');
    requestAnimationFrame(() => host.classList.add('in'));
    window.MX && MX.pressAll(host);
    document.addEventListener('keydown', onKey);
    renderPage(body);
    if (ctx.onSeen) ctx.onSeen();
  }
  function close() {
    if (!host) return; document.removeEventListener('keydown', onKey);
    host.classList.remove('in'); const h = host; host = null; document.body.style.overflow = ''; document.body.classList.remove('rd-open');
    setTimeout(() => h.remove(), 220); state = null;
  }
  function onKey(e) {
    if (e.key === 'Escape') return close();
    const b = document.getElementById('rdBody'); if (!b) return;
    if (e.key === 'ArrowDown' || e.key === 'PageDown' || e.key === ' ') { e.preventDefault(); b.scrollBy({ top: b.clientHeight * 0.9, behavior: smooth() }); }
    if (e.key === 'ArrowUp' || e.key === 'PageUp') { e.preventDefault(); b.scrollBy({ top: -b.clientHeight * 0.9, behavior: smooth() }); }
  }
  function smooth() { return window.MX && MX.reduce() ? 'auto' : 'smooth'; }

  async function doDownload(btn) {
    const done = window.MX ? MX.busy(btn, 'Preparing…') : () => {};
    try { saveBlob(await state.blob(), state.filename || 'report.docx'); done('Saved ✓'); }
    catch (e) { done('Failed', 'err'); window.MX && MX.toast('Could not build the Word file: ' + (e.message || e)); }
  }
  function saveBlob(blob, name) { const a = el('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 30000); }

  const settle = async root => {
    await Promise.all([...root.querySelectorAll('img')].map(im => (im.decode ? im.decode() : Promise.resolve()).catch(() => {})));
    if (document.fonts && document.fonts.ready) { try { await document.fonts.ready; } catch (e) {} }   // fonts decide the wraps
  };
  const RENDER = { inWrapper: true, ignoreLastRenderedPageBreak: true, experimental: true };

  /* ---------- the .docx on screen ----------
     §25.1 (08.11): docx-preview does not flow pages. So we render the report once flat (off-screen), measure the
     REAL rendered height of every block and table row, work out where the A4 pages break (a heading stays with the
     table under it; a long table splits between rows, its header repeated), then render one section per page at
     exactly those breaks. Pagination comes from what is actually drawn, never from an estimate. */
  async function renderPage(body) {
    const stage = el('div', 'rd-stage'); const scaleEl = el('div', 'rd-scale'); stage.appendChild(scaleEl); body.appendChild(stage);
    body.appendChild(zoomBar());
    const sk = window.MX ? MX.skeleton('page') : el('div'); scaleEl.appendChild(sk);
    try {
      let blob;
      if (state.measure && state.display) {
        const m = await state.measure();                      // { blob, meta }
        const mroot = el('div', 'rd-measure'); document.body.appendChild(mroot);
        try {
          await window.docxPreview.renderAsync(m.blob, mroot, null, RENDER);
          await settle(mroot);
          tightenBody(mroot);                                 // Word single spacing on non-table paras, before measuring
          const spec = computeBreaks(mroot, m.meta || []);
          blob = await state.display(spec);
        } finally { mroot.remove(); }
      } else {
        blob = await (state.renderBlob || state.blob)();
      }
      scaleEl.innerHTML = '';
      await window.docxPreview.renderAsync(blob, scaleEl, null, RENDER);
      await settle(scaleEl);
      tightenBody(scaleEl);
      seatCoverFrame(scaleEl);
      spaceCoverRuns(scaleEl);
      dotLeaders(scaleEl);
      fit(scaleEl, body); enableGestures(stage, scaleEl, body);
      window.MX && MX.vibrate(10);
    } catch (e) {
      scaleEl.innerHTML = ''; const m = el('div', 'rd-msg err', 'Could not show the report on screen: ' + (e.message || e));
      const b = el('button', 'btn', 'Download instead'); b.onclick = () => doDownload(b); m.appendChild(b); scaleEl.appendChild(m);
    }
  }
  // §25.1 (08.11 r4): docx-preview gives non-table body paragraphs (section headings, italic notes) a far looser
  // line-height than Word's single spacing (~1.5× vs ~1.15×), which is where the reader's extra page came from.
  // Set each such paragraph's line-height to the single line of its largest run, as Word does for auto spacing.
  // Scoped to the body sections (not the cover, whose lines set their own exact height). The .docx is untouched.
  function tightenBody(root) {
    [...root.querySelectorAll('section.docx')].slice(1).forEach(sec => {
      sec.querySelectorAll('p').forEach(p => {
        if (p.closest('table') || p.closest('footer')) return;
        let max = 0; p.querySelectorAll('span').forEach(s => { const f = parseFloat(getComputedStyle(s).fontSize) || 0; if (f > max) max = f; });
        if (!max) max = parseFloat(getComputedStyle(p).fontSize) || 0;
        if (max) p.style.lineHeight = (Math.round(max * 1.15 * 10) / 10) + 'px';   // Arial single spacing
      });
    });
  }

  /* ---------- measured pagination ---------- */
  // Walk the flat body once rendered, filling A4-high pages from the real block heights. Returns the break spec the
  // builder applies: before = flat indices that start a new page; cuts = { tableIndex: [rowsToSplitAt] }.
  function computeBreaks(root, meta) {
    const secs = [...root.querySelectorAll('section.docx')];
    const bodySec = secs[secs.length - 1]; if (!bodySec) return { before: [], cuts: {} };
    const cs = getComputedStyle(bodySec);
    const padT = parseFloat(cs.paddingTop) || 0, padB = parseFloat(cs.paddingBottom) || 0;
    const pageH = parseFloat(cs.minHeight) || 1122.5;
    const avail = pageH - padT - padB - 1;                   // usable content height per A4 page
    const article = [...bodySec.children].find(c => c.tagName === 'ARTICLE') || bodySec;
    const blocks = [...article.children].filter(c => c.tagName !== 'FOOTER');
    const mh = e => { const r = e.getBoundingClientRect(); const s = getComputedStyle(e); return r.height + (parseFloat(s.marginTop) || 0) + (parseFloat(s.marginBottom) || 0); };
    const bh = blocks.map(mh);
    const n = blocks.length;
    const isTable = i => (meta[i] ? meta[i].table : blocks[i].tagName === 'TABLE');
    const headN = i => Math.max(1, (meta[i] && meta[i].head) || 1);
    const rowsOf = i => [...blocks[i].querySelectorAll('tr')];
    // a "heading group" is one or more non-table paragraphs (a section heading plus its italic notes) immediately
    // followed by a table. keep-with-next: the group must not be left at the foot of a page with its table pushed to
    // the next — but it only needs the table's header row + ONE body row with it, not the whole table.
    const isHeadingP = i => !isTable(i) && /^\d{1,2}\.\s/.test((blocks[i].textContent || '').trim());
    const groupStart = i => isHeadingP(i);                    // a numbered section heading, not a trailing note
    const groupNeed = i => {
      let j = i, h = 0; const lim = i + 6;
      while (j < n && !isTable(j) && j < lim) { h += bh[j]; j++; }
      if (j >= n || !isTable(j)) return null;                 // no table follows → not a heading group
      if (isWrapper(j)) {
        // a side-by-side wrapper can split, so the heading only needs its nested header + 2 body rows beside it, not
        // the whole wrapper (querySelectorAll('tr') would wrongly return the tall outer row first).
        const cols = wrapCols(j);
        if (cols) return h + Math.max.apply(null, cols.map(c => c.preH + c.headH + (c.rh[c.head] || 0) + (c.rh[c.head + 1] || 0)));
        return h + Math.min(bh[j], avail * 0.5);
      }
      const rh = rowsOf(j).map(r => r.getBoundingClientRect().height), hc = headN(j);
      return h + rh.slice(0, hc).reduce((a, b) => a + b, 0) + (rh[hc] || 0);
    };

    // a side-by-side wrapper: one outer row whose content cells each hold one nested table (meta.pair lists the
    // columns' nested {rows, head}). Measure the nested rows so the wrapper can be split between them, both columns
    // broken on the same page, each nested header repeated — the reader's way to cut the TMS dashboard and §ab pairs.
    const isWrapper = i => isTable(i) && meta[i] && meta[i].rows === 1 && meta[i].pair && blocks[i].querySelectorAll('table').length > 0;
    const wrapCols = i => {
      const outerRow = blocks[i].querySelector(':scope > tbody > tr') || blocks[i].querySelector(':scope > tr');
      if (!outerRow) return null;
      const cells = [...outerRow.children].filter(td => td.querySelector('table'));   // content cells (gutter has none)
      const pair = meta[i].pair; if (cells.length !== pair.length) return null;
      const top = blocks[i].getBoundingClientRect().top, cols = [];
      for (let c = 0; c < cells.length; c++) {
        const inner = [...cells[c].querySelectorAll('table')].filter(t => !t.querySelector('table'));
        if (inner.length !== 1 || !pair[c][0]) return null;         // only split a column with exactly one nested table
        const tb = inner[0], trs = [...tb.querySelectorAll('tr')];
        const head = Math.min(pair[c][0].head || 0, trs.length - 1);
        const rh = trs.map(r => r.getBoundingClientRect().height);
        cols.push({ head: head, body: trs.length - head, rh: rh, headH: rh.slice(0, head).reduce((a, b) => a + b, 0), preH: tb.getBoundingClientRect().top - top });
      }
      return cols;
    };
    // body rows of column c that fit in `budget`, starting at body-row `start`, with `pre` above the header (piece 1 only)
    const fitRows = (c, budget, start, pre) => { let acc = pre + c.headH, k = start; for (; k < c.body; k++) { if (acc + c.rh[c.head + k] > budget && k > start) break; acc += c.rh[c.head + k]; } return k; };
    const splitWrap = (i, used0) => {
      const cols = wrapCols(i); if (!cols || !cols.every(c => c.body >= 4)) return null;   // need >=2 each side
      if (cols.some(c => fitRows(c, avail - used0, 0, c.preH) < 2)) return null;            // no room on this page
      const starts = cols.map(() => 0), lists = cols.map(() => []); let pg = 0, lastH = cols.map(() => 0);
      while (starts.some((s, ci) => s < cols[ci].body)) {
        if (++pg > 60) return null;
        const budget = pg === 1 ? avail - used0 : avail;
        cols.forEach((c, ci) => {
          if (starts[ci] >= c.body) { lastH[ci] = 0; return; }
          let k = fitRows(c, budget, starts[ci], pg === 1 ? c.preH : 0);
          if (k < c.body && c.body - k < 2) k = c.body - 2;         // keep >=2 in the tail
          if (k - starts[ci] < 2) k = Math.min(starts[ci] + 2, c.body);
          let h = (pg === 1 ? c.preH : 0) + c.headH; for (let r = starts[ci]; r < k; r++) h += c.rh[c.head + r];
          lastH[ci] = k > starts[ci] ? h : 0;
          if (k < c.body) lists[ci].push(k);
          starts[ci] = k;
        });
      }
      return { lists: lists, usedAfter: Math.max.apply(null, lastH) };
    };

    const before = new Set(), cuts = {};
    let used = 0;
    for (let i = 0; i < n; i++) {
      if (used > 0 && groupStart(i)) {
        const need = groupNeed(i);
        if (need != null && used + need > avail) { before.add(i); used = 0; }
      }
      if (used + bh[i] <= avail) { used += bh[i]; continue; }   // fits on this page
      // a table that does not fit the space left is split between rows (its header repeated) to fill the page as Word
      // does, keeping >=2 body rows on each page so no stray single row is left. A plain table (rendered row count ==
      // the builder's, no nested tables) splits here; a side-by-side wrapper is split further below.
      if (isTable(i) && meta[i] && rowsOf(i).length === meta[i].rows && blocks[i].querySelectorAll('table').length === 0) {
        const rows = rowsOf(i), rh = rows.map(r => r.getBoundingClientRect().height), hc = Math.min(headN(i), rows.length - 1);
        const headH = rh.slice(0, hc).reduce((a, b) => a + b, 0), body = rows.length - hc;
        if (body >= 4) {
          // fill the page with body rows (header repeated on each page), each page kept to >=2 body rows; a stray
          // last row is avoided by borrowing one back from the previous page, so no page gets a single row.
          const segs = [[]]; let seg = used + headH;
          for (let r = hc; r < rows.length; r++) {
            const cur = segs[segs.length - 1];
            if (seg + rh[r] > avail && cur.length >= 2) { segs.push([]); seg = headH; }
            segs[segs.length - 1].push(r); seg += rh[r];
          }
          while (segs.length > 1 && segs[segs.length - 1].length < 2 && segs[segs.length - 2].length > 2) {
            const moved = segs[segs.length - 2].pop(); segs[segs.length - 1].unshift(moved);
          }
          // split only when it fills this page (first segment >=2 rows) and leaves a clean tail (last segment >=2)
          if (segs.length >= 2 && segs[0].length >= 2 && segs[segs.length - 1].length >= 2) {
            cuts[i] = segs.slice(1).map(s => s[0]);                     // first row index of each continuation page
            used = headH + segs[segs.length - 1].reduce((a, r) => a + rh[r], 0);
            continue;
          }
        }
      }
      // a side-by-side wrapper is split between its nested rows to fill the page (the TMS dashboard, §ab pairs)
      if (isWrapper(i)) {
        let res = splitWrap(i, used);                          // try to split filling the current page
        if (!(res && res.lists.some(l => l.length)) && used > 0) { before.add(i); used = 0; res = splitWrap(i, 0); }
        if (res && res.lists.some(l => l.length)) { cuts[i] = { pair: res.lists }; used = res.usedAfter; continue; }
      }
      if (used > 0) { before.add(i); used = 0; }               // move the whole block to a fresh page
      used += bh[i];
    }
    return { before: [...before], cuts };
  }

  // §25.1 (08.11): the cover's full-page frame is a floating image anchored to the PAGE at offset 0, so in Word it
  // sits edge-to-edge behind the text. docx-preview instead lays it out from the content box (inside the page
  // margins), which shifts the whole frame right — its left rule lands on the contents text and its right edge is
  // clipped. This is a docx-preview placement gap, not the Word file: so, for display only, we re-seat that one
  // full-page image to fill its page box, behind the text, exactly where Word draws it. The .docx is never touched.
  function seatCoverFrame(scaleEl) {
    scaleEl.querySelectorAll('section.docx').forEach(sec => {
      const sr = sec.getBoundingClientRect();
      [...sec.querySelectorAll('img')].forEach(im => {
        const r = im.getBoundingClientRect();
        if (im.naturalHeight > im.naturalWidth && r.width > sr.width * 0.8 && r.height > sr.height * 0.8) {
          for (let w = im.parentElement; w && w !== sec; w = w.parentElement) if (getComputedStyle(w).position !== 'static') w.style.position = 'static';
          im.style.position = 'absolute'; im.style.left = '0'; im.style.top = '0';
          im.style.width = '100%'; im.style.height = '100%'; im.style.margin = '0'; im.style.zIndex = '-1';
        }
      });
    });
  }
  // §25.1 (08.11): the cover's two tracked lines carry character spacing in the .docx (w:spacing), which docx-preview
  // does not render. Re-apply it on screen as letter-spacing (em, so it survives the fit-scale). Display only.
  function spaceCoverRuns(scaleEl) {
    const sec = scaleEl.querySelector('section.docx'); if (!sec) return;
    sec.querySelectorAll('p').forEach(p => {
      const t = (p.textContent || '').trim();
      if (/^EASTERN RAILWAY\b.*HOWRAH DIVISION$/.test(t)) p.style.letterSpacing = '0.32em';
      else if (/^AEN\/\S+\s+·\s+\S/.test(t) && t === t.toUpperCase()) p.style.letterSpacing = '0.08em';
    });
  }
  // §25.1 (08.11) item 3: the cover's contents list has dot leaders in Word (……… before each page number) via a
  // right tab stop with a DOT leader. docx-preview positions the number but does not draw the dots, so draw them on
  // screen: the gap before a page number is a whitespace-only span — give it a dotted baseline. Display only.
  function dotLeaders(scaleEl) {
    const sec = scaleEl.querySelector('section.docx'); if (!sec) return;
    sec.querySelectorAll('p').forEach(p => {
      const spans = [...p.children].filter(c => c.tagName === 'SPAN'); if (spans.length < 3) return;
      const last = spans[spans.length - 1], gap = spans[spans.length - 2];
      if (!/^\s*\d+\s*$/.test(last.textContent || '')) return;        // the page number
      if ((gap.textContent || '').trim() !== '') return;              // the tab gap before it
      gap.style.borderBottom = '1.6px dotted rgba(40,30,22,.45)';
      gap.style.margin = '0 4px'; gap.style.transform = 'translateY(-3px)';
    });
  }
  function pageWidth(scaleEl) { let w = 0; scaleEl.querySelectorAll('section.docx').forEach(s => w = Math.max(w, s.getBoundingClientRect().width / (state.scale || 1))); if (!w) { const wr = scaleEl.querySelector('.docx-wrapper'); w = wr ? wr.getBoundingClientRect().width / (state.scale || 1) : scaleEl.getBoundingClientRect().width; } return w; }
  function applyScale(scaleEl, scale) {
    state.scale = scale; scaleEl.style.transformOrigin = 'top center'; scaleEl.style.transform = 'scale(' + scale + ')';
    const wr = scaleEl.querySelector('.docx-wrapper'); const h = wr ? wr.getBoundingClientRect().height : scaleEl.getBoundingClientRect().height;
    scaleEl.parentNode.style.height = h + 'px';
  }
  function fit(scaleEl, body) { const w = pageWidth(scaleEl) || 1; applyScale(scaleEl, Math.min(1.5, (body.clientWidth - 24) / w)); markZoom('fit'); }
  function setZoom(z) { const body = document.getElementById('rdBody'), scaleEl = body && body.querySelector('.rd-scale'); if (!scaleEl) return; state.zoom = z;
    if (z === 'fit') fit(scaleEl, body); else { applyScale(scaleEl, z === '150' ? 1.5 : 1); markZoom(z); } window.MX && MX.vibrate(10); }
  const ZOOMS = ['fit', '100', '150'], ZLAB = ['Fit', '100%', '150%'];
  function markZoom(z) { const bar = document.querySelector('.rd-zoom'); if (!bar) return;
    bar.querySelectorAll('.rd-zoom-lab').forEach((n, i) => n.classList.toggle('on', ZOOMS[i] === z));
    const sl = bar.querySelector('.rd-zoom-slider'); const idx = ZOOMS.indexOf(z); if (sl && idx >= 0) sl.value = String(idx); }
  // §25.6: a small, semi-transparent weighted slider with notches at Fit / 100 / 150 (Android vibrate at notches)
  function zoomBar() {
    const bar = el('div', 'rd-zoom');
    const labels = el('div', 'rd-zoom-labels'); ZLAB.forEach((t, i) => { const s = el('span', 'rd-zoom-lab' + (i === 0 ? ' on' : '')); s.textContent = t; s.onclick = () => setZoom(ZOOMS[i]); labels.appendChild(s); });
    const sl = el('input', 'rd-zoom-slider'); sl.type = 'range'; sl.min = '0'; sl.max = '2'; sl.step = '1'; sl.value = '0'; sl.setAttribute('aria-label', 'Zoom');
    let last = 0; sl.oninput = () => { const i = +sl.value; if (i !== last) { window.MX && MX.vibrate(10); last = i; } setZoom(ZOOMS[i]); };
    bar.append(labels, sl); return bar;
  }
  function enableGestures(stage, scaleEl, body) {
    let lastTap = 0; stage.addEventListener('pointerup', () => { const t = Date.now(); if (t - lastTap < 300) setZoom(state.zoom === 'fit' ? '100' : 'fit'); lastTap = t; });
    let pts = new Map(), baseDist = 0, baseScale = 1;
    stage.addEventListener('pointerdown', e => pts.set(e.pointerId, e));
    stage.addEventListener('pointermove', e => { if (!pts.has(e.pointerId)) return; pts.set(e.pointerId, e);
      if (pts.size === 2) { const [a, b] = [...pts.values()]; const d = Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY);
        if (!baseDist) { baseDist = d; baseScale = state.scale || 1; } else { applyScale(scaleEl, Math.max(0.4, Math.min(2.2, baseScale * d / baseDist))); state.zoom = 'pinch'; markZoom(''); } } });
    const end = e => { pts.delete(e.pointerId); if (pts.size < 2) baseDist = 0; };
    stage.addEventListener('pointerup', end); stage.addEventListener('pointercancel', end);
  }

  /* ---------- overflow menu on narrow screens ---------- */
  function moreSheet() {
    const box = el('div', 'rd-more');
    const mk = el('button', 'btn-line', 'Report a mistake'); mk.dataset.press = '1'; mk.onclick = () => { closeSheet(); setTimeout(mistakeSheet, 60); };
    const pr = el('button', 'btn-line', 'Print'); pr.dataset.press = '1'; pr.onclick = () => { closeSheet(); window.print(); };
    box.append(mk, pr); setTimeout(() => window.MX && MX.pressAll(box), 0);
    var closeSheet = MX.sheet(box);
  }

  /* ---------- Report a mistake (§25.4 item 3) ---------- */
  function mistakeSheet() {
    const T = window.SECTION_TITLES || {};
    const present = Object.keys(state.report.sections || {}).filter(k => T[k]);
    const box = el('div', 'mk');
    box.appendChild(el('div', 'why-h', 'Report a mistake'));
    box.appendChild(el('div', 'small faint', 'Pick the section and write one line. It is sent to the control room; the report does not change.'));
    const sel = el('select', 'mk-sel'); present.forEach(k => { const o = el('option', '', (T[k].no) + '. ' + T[k].title); o.value = k; sel.appendChild(o); });
    const ta = el('textarea', 'mk-text'); ta.rows = 2; ta.placeholder = 'What is wrong?';
    const send = el('button', 'btn', 'Send'); send.dataset.press = '1';
    const msg = el('div', 'msg small');
    send.onclick = async () => {
      const text = ta.value.trim(); if (!text) { msg.textContent = 'Write one line first.'; msg.className = 'msg small err'; return; }
      const done = MX.busy(send, 'Sending…');
      try { await state.reportMistake(sel.value, text); done('Sent ✓'); msg.textContent = 'Thank you — sent to the control room.'; msg.className = 'msg small ok'; ta.value = ''; setTimeout(closeSheet, 1200); }
      catch (e) { done('Failed', 'err'); msg.textContent = e.message || 'Could not send.'; msg.className = 'msg small err'; }
    };
    box.append(sel, ta, send, msg);
    setTimeout(() => window.MX && MX.pressAll(box), 0);
    var closeSheet = MX.sheet(box);
  }

  window.Reader = { open, close, computeBreaks, seatCoverFrame, spaceCoverRuns, dotLeaders, tightenBody };
})();
