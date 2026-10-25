/*
 * Builds the Morning Position Word file from the report JSON the server returns.
 * Works in the browser (global `docx` from vendor/docx-9.8.1.iife.js) and in Node
 * (tests pass require('docx')). Layout as agreed: official A4, black ruled tables,
 * bold centred grey headers, numbers centred, then the two blank pages for Mini Control.
 */
(function (g) {
  function buildMorningDoc(X, R, opts) {
    opts = opts || {};
    const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
      AlignmentType, VerticalAlign, Footer, PageNumber, TableLayoutType, HeightRule, UnderlineType, TabStopType, ImageRun } = X;

    const CREDIT = 'Designed & developed by Anindya Nandy, SSE/P.Way/Planning, Howrah Division';
    const W = 10772, FONT = 'Arial', SZ = 19, GREY = 'D9D9D9';
    const C = AlignmentType.CENTER, L = AlignmentType.LEFT, RT = AlignmentType.RIGHT;
    const b = { style: BorderStyle.SINGLE, size: 6, color: '000000' }, nb = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const ALL = { top: b, bottom: b, left: b, right: b }, NONE = { top: nb, bottom: nb, left: nb, right: nb };
    const S = R.sections, D = R.date;

    /* ---------- small helpers ---------- */
    const dmy = iso => iso ? iso.slice(8, 10) + '.' + iso.slice(5, 7) + '.' + iso.slice(0, 4) : '';
    const dm = iso => iso ? iso.slice(8, 10) + '.' + iso.slice(5, 7) : '';
    const pct = v => v === null || v === undefined ? '–' : Math.round(v * 100) + '%';
    const val = v => v === null || v === undefined || v === '' ? '–' : (typeof v === 'number' ? (Number.isInteger(v) ? String(v) : v.toFixed(2)) : String(v));
    const tag = (r, text, size, bold, imgH) => { r._t = String(text); r._s = size || SZ; r._b = !!bold; if (imgH) r._img = imgH; return r; };
    const run = (t, o = {}) => tag(new TextRun({ text: String(t), font: FONT, size: o.size || SZ, bold: !!o.bold, italics: !!o.italics,
      underline: o.underline ? { type: UnderlineType.SINGLE } : undefined }), t, o.size, o.bold);
    const para = (t, o = {}) => {
      const kids = (Array.isArray(t) ? t : [t]).map(x => typeof x === 'string' ? run(x, o) : x);
      const p = new Paragraph({ children: kids, alignment: o.align || L, spacing: { before: o.before || 0, after: o.after || 0 }, keepNext: !!o.keepNext, tabStops: o.tabs });
      p._runs = kids; p._sp = (o.before || 0) + (o.after || 0); return p;
    };

    /* ---------- page filling estimate (twips) ----------
       Word lays the pages out, not us, so we estimate: line height and average letter width from the font size.
       Used only to decide whether the blank pages start on a fresh page or fill the rest of the last one. */
    const PH = 16838 - 567 - 680 - 60;                  // usable A4 height inside margins (and the footer)
    const lineH = s => Math.round(s * 11.6);              // half-points -> twips, Arial single spacing
    const charW = (s, b) => s * (b ? 5.45 : 5.1);          // average Arial letter width in twips
    const paraH = (p, width) => {
      if (!p || !p._runs) return lineH(SZ);
      let w = 0, s = 0, img = 0;
      p._runs.forEach(r => { if (r._img) img = Math.max(img, r._img); if (r._t) { w += r._t.length * charW(r._s, r._b); s = Math.max(s, r._s); } });
      const lines = Math.max(1, Math.ceil(w / Math.max(400, width)));
      return Math.max(img, lines * lineH(s || SZ)) + p._sp;
    };
    const cell = (content, width, o = {}) => { const kids = Array.isArray(content) ? content : [para(content, o)]; const c = new TableCell({
      width: { size: width, type: WidthType.DXA }, borders: o.borders || ALL, columnSpan: o.span,
      shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
      margins: { top: 20, bottom: 20, left: 60, right: 60 }, verticalAlign: o.valign || VerticalAlign.CENTER,
      children: kids });
      c._h = 40 + kids.reduce((a, k) => a + (k._h !== undefined ? k._h : paraH(k, width - 120)), 0); return c; };
    const TR = o => { const r = new TableRow(o); r._h = o.height && o.height.rule === HeightRule.EXACT ? o.height.value : Math.max(...o.children.map(c => c._h || lineH(SZ) + 40));
      r._head = !!o.tableHeader; r._cant = !!o.cantSplit; return r; };
    const table = (widths, rows) => { const tb = new Table({ width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA },
      columnWidths: widths, layout: TableLayoutType.FIXED, rows }); tb._rows = rows; tb._h = rows.reduce((a, r) => a + r._h, 0); return tb; };
    const H = (labels, widths, size) => TR({ tableHeader: true, cantSplit: true,
      children: labels.map((l, i) => cell(l, widths[i], { fill: GREY, bold: true, align: C, size })) });
    const Rw = (vals, widths, aligns, o = {}) => TR({ cantSplit: true,
      children: vals.map((v, i) => v instanceof TableCell ? v : cell(val(v), widths[i], { align: aligns[i] === 'c' ? C : L, size: o.size, bold: o.bold })) });
    const fit = (parts, total) => { const s = parts.reduce((a, c) => a + c, 0); const out = parts.map(p => Math.floor(p * total / s)); out[out.length - 1] += total - out.reduce((a, c) => a + c, 0); return out; };

    const section = (no, title, note) => para([run(no + '.  ', { bold: true, size: 21 }), run(title, { bold: true, underline: true, size: 21 }),
      ...(note ? [run('  (' + note + ')', { size: 17, italics: true })] : [])], { before: 110, after: 50, keepNext: true });
    const half = (W - 240) / 2;
    const pair = (left, right) => table([half, 240, half], [TR({ children: [
      cell(left, half, { borders: NONE, valign: VerticalAlign.TOP }), cell([para('')], 240, { borders: NONE }),
      cell(right, half, { borders: NONE, valign: VerticalAlign.TOP })] })]);
    // (pair's height: the taller of its two sides, each a list of paragraphs and tables)
    const nilLine = (text) => para([run(text, { size: SZ })], { after: 20 });
    const problem = (s) => para([run('Could not be read: ', { bold: true, size: 18 }), run(s.error, { italics: true, size: 18 })], { after: 20 });
    const staleNote = (iso, stale) => stale && iso ? 'data of ' + dmy(iso) + ', latest available' : null;
    const kv = (rows, w1, width) => { const w = [w1, width - w1]; return table(w, rows.map(([k, v]) => Rw([k, v], w, ['l', 'c']))); };

    /* ---------- heading: Eastern Railway letterhead (पूर्व रेलवे · IR logo · EASTERN RAILWAY) ---------- */
    const HIN = 'Nirmala UI';      // Hindi font present on every Windows PC
    const hrun = (t, o = {}) => tag(new TextRun({ text: t, font: { ascii: HIN, hAnsi: HIN, cs: HIN }, size: o.size || SZ, bold: !!o.bold, sizeComplexScript: o.size || SZ, boldComplexScript: !!o.bold }), t, (o.size || SZ) * 1.25, o.bold);
    const HINDI_OFFICE = { 'TR/HWH': 'सहायक मंडल अभियंता/ट्रैक/हावड़ा', LLH: 'सहायक मंडल अभियंता, लिलुआ', BDC: 'सहायक मंडल अभियंता, बंडेल',
      DKAE: 'सहायक मंडल अभियंता, डानकुनी', BWN: 'सहायक मंडल अभियंता, बर्धमान', NDAE: 'सहायक मंडल अभियंता, नवद्वीप धाम', KWAE: 'सहायक मंडल अभियंता, कटवा',
      AZ: 'सहायक मंडल अभियंता, अजीमगंज', BHP: 'सहायक मंडल अभियंता, बोलपुर', RPH: 'सहायक मंडल अभियंता, रामपुरहाट' };
    const rule = (sz) => ({ style: BorderStyle.SINGLE, size: sz, color: '000000' });
    const lw = [3900, W - 8300, 4400];
    const logoRun = tag(opts.logo ? new ImageRun({ type: 'png', data: opts.logo, transformation: { width: 70, height: 70 } }) : new TextRun(''), '', SZ, false, 1100);
    const mast = table(lw, [
      TR({ children: [
        cell([Object.assign(new Paragraph({ children: [hrun('प्रातःकालीन स्थिति', { size: 26, bold: true }), new TextRun({ text: '\tMorning Position', font: FONT, size: 30, bold: true }),
          new TextRun({ text: '\tDate: ' + dmy(D), font: FONT, size: 20, bold: true })],
          tabStops: [{ type: TabStopType.CENTER, position: Math.round(W / 2) }, { type: TabStopType.RIGHT, position: W - 120 }] }), { _h: 420 })], W, { span: 3, borders: NONE, valign: VerticalAlign.BOTTOM })] }),
      TR({ children: [
        cell([para([hrun('पूर्व रेलवे', { size: 56, bold: true })], { align: L })], lw[0], { borders: { top: rule(12), bottom: nb, left: nb, right: nb } }),
        cell([para([logoRun], { align: C, before: 40, after: 40 })], lw[1], { borders: { top: rule(12), bottom: nb, left: nb, right: nb } }),
        cell([para([new TextRun({ text: 'EASTERN RAILWAY', font: 'Arial Narrow', size: 44, bold: true })], { align: RT })], lw[2], { borders: { top: rule(12), bottom: nb, left: nb, right: nb } })] }),
      TR({ children: [
        cell([para([hrun(HINDI_OFFICE[R.aen] + ' का कार्यालय, हावड़ा मंडल', { size: 19 })], { align: C })], W, { span: 3, borders: { top: rule(6), bottom: nb, left: nb, right: nb } })] }),
      TR({ children: [
        cell([para([run('Office of the ' + R.office + ', Howrah Division', { size: 19, bold: true })], { align: C })], W, { span: 3, borders: { top: nb, bottom: rule(12), left: nb, right: nb } })] })
    ]);
    const letterhead = (title, sub) => [
      mast,
      ...(title && title !== 'MORNING POSITION' ? [para([run(title, { bold: true, underline: true, size: 26 })], { align: C, before: 120 })] : []),
      ...(sub ? [para([run(sub, { size: 18, italics: true })], { align: C, before: 60 })] : []),
      para([run('Units: ' + R.units.map(u => 'SSE/P.Way/' + u).join(', '), { size: 19 })], { before: 80, after: 60 })
    ];

    /* ---------- abstract ---------- */
    const ok = k => S[k] && S[k].ok;
    const tdcN = ok('tdcLapse') ? S.tdcLapse.rows.length : null;
    const excN = ok('miniControl') ? S.miniControl.rows.length : null;
    const failN = ok('major') && ok('minor') ? S.major.rows.length + S.minor.rows.length : null;
    const blk = ok('blocks') ? S.blocks : null, mc = ok('machines') ? S.machines : null;
    const nilOr = n => n === null ? '–' : (n === 0 ? 'Nil' : String(n));
    const abstractCells = [
      ['Caution past TDC', nilOr(tdcN)], ['Mini control exceptions', nilOr(excN)], ['Failures (' + dm(R.period.day) + ')', nilOr(failN)],
      ['LCs with bad road', ok('badRoad') && S.badRoad.row ? val(S.badRoad.row.total) : '–'],
      ['Blocks availed', blk ? blk.rows.length + ' (' + blk.rows.reduce((a, r) => a + (r.availed || 0), 0) + ' min)' : '–'],
      ['Blocks not availed', blk && blk.notAvailedDay !== null ? blk.notAvailedDay + ' (month: ' + blk.notAvailedMonth + ')' : '–'],
      ['M/C block utilisation', mc ? pct(mc.total.util) : '–'], ['M/C days lost', mc ? String(mc.lost.length) : '–']];
    const aw = fit(Array(8).fill(1), W);
    const abstract = table(aw, [H(abstractCells.map(c => c[0]), aw, 17), Rw(abstractCells.map(c => c[1]), aw, Array(8).fill('c'), { bold: true })]);

    /* ---------- 1. mini control exceptions + GPS ---------- */
    const s1 = [section(1, 'Mini control exceptions', ok('miniControl') ? staleNote(S.miniControl.dataDate, S.miniControl.stale) : null)];
    if (!ok('miniControl')) s1.push(problem(S.miniControl));
    else if (!S.miniControl.rows.length) s1.push(kv([['Mini control exceptions (' + dm(S.miniControl.dataDate) + ')', 'Nil']], 3200, half));
    else { const w = [1800, half - 1800]; s1.push(table(w, [H(['Item', 'Details of exception'], w), ...S.miniControl.rows.map(r => Rw([r.item, r.detail], w, ['l', 'l']))])); }
    if (!ok('gps')) s1.push(problem(S.gps));
    else if (!S.gps.started || !S.gps.row) s1.push(kv([['Winter patrolling (GPS)', 'Season not started']], 3200, half));
    else {
      const g2 = S.gps.row;
      s1.push(kv([['Patrol batches (GPS)', val(g2.total)], ['Defaulter batches', val(g2.defaulters)],
        ...g2.kinds.map(k => [k.name, val(k.n)]), ...(g2.action ? [['Action taken by AEN', g2.action]] : [])], 3200, half));
    }

    /* ---------- 2. failures ---------- */
    const s2 = [section(2, 'Failures', dmy(R.period.day))];
    const failTable = (label, s, cols) => {
      if (!s.ok) return [problem(s)];
      if (!s.rows.length) return [kv([[label, 'Nil']], 3200, half)];
      const w = fit([1, 1.6, 2.4], half);
      return [para([run(label, { bold: true, size: 18 })], { before: 30 }), table(w, [H(['Station / section', 'Line / km', 'Nature of failure'], w, 17),
        ...s.rows.map(r => Rw(cols(r), w, ['l', 'c', 'l'], { size: 17 }))])];
    };
    s2.push(...failTable('Major asset failure', S.major, r => [r.section, r.line, r.type + (r.duration ? ' (' + r.duration + ')' : '')]));
    s2.push(...failTable('Minor failure', S.minor, r => [r.station, [r.line, r.km].filter(Boolean).join(', '), r.type]));

    /* ---------- 3. cautions ---------- */
    const s3 = [section(3, 'Cautions', 'TDC lapse, PQRS 30 km/h TDC fail, protection caution')];
    let anyCaution = false;
    if (!ok('tdcLapse')) s3.push(problem(S.tdcLapse));
    else if (S.tdcLapse.rows.length) {
      anyCaution = true;
      const cw = [400, 750, 1600, 1050, 2100, 1000, 700, W - 7600];
      s3.push(para([run('Caution TDC lapse', { bold: true, size: 18 })], { after: 30 }), table(cw, [
        H(['Sl.', 'Speed', 'Section / Line', 'Km', 'Nature of work', 'TDC', 'Days past TDC', 'Remarks'], cw),
        ...S.tdcLapse.rows.map((r, i) => Rw([String(i + 1), r.speed, r.section + ', ' + r.line, r.km, r.work, r.tdc ? dmy(r.tdc).slice(0, 6) + r.tdc.slice(2, 4) : '–', val(r.daysPast), r.remarks], cw,
          ['c', 'c', 'l', 'c', 'l', 'c', 'c', 'l']))]));
    }
    if (!ok('pqrsTdc')) s3.push(problem(S.pqrsTdc));
    else if (S.pqrsTdc.rows.length) {
      anyCaution = true;
      const pw = [400, 2000, 1500, 1100, 1100, 1100, 1300, W - 8500];
      s3.push(para([run('PQRS – 30 km/h not relaxed in time', { bold: true, size: 18 })], { before: 60, after: 30 }), table(pw, [
        H(['Sl.', 'Section / Line', 'Location', '30 km/h imposed', 'Relaxed on', '10th day', 'Final removal', 'Reason for delay'], pw, 17),
        ...S.pqrsTdc.rows.map((r, i) => Rw([String(i + 1), r.section + ', ' + r.line, r.location, r.imposed && r.imposed.length === 10 ? dmy(r.imposed) : r.imposed,
          r.relaxed, r.tenthDay, r.removal, r.removalReason || r.relaxReason], pw, ['c', 'l', 'c', 'c', 'c', 'c', 'c', 'l'], { size: 17 }))]));
    }
    if (!ok('protection')) s3.push(problem(S.protection));
    else if (S.protection.rows.length) {
      anyCaution = true;
      const vw = [400, 1500, 2600, 1300, 2400, W - 8200];
      s3.push(para([run('Protection caution for PQRS, BCM & M/R', { bold: true, size: 18 })], { before: 60, after: 30 }), table(vw, [
        H(['Sl.', 'PWI', 'Section', 'Line', 'Type of work', 'Caution on adjacent line'], vw, 17),
        ...S.protection.rows.map((r, i) => Rw([String(i + 1), r.pwi, r.section, r.line, r.work, r.imposed], vw, ['c', 'l', 'l', 'c', 'l', 'c'], { size: 17 }))]));
    }
    if (!anyCaution && ok('tdcLapse') && ok('pqrsTdc') && ok('protection')) s3.push(nilLine('Nil'));

    /* ---------- 4. bad road, 5. CCTV ---------- */
    const s4 = [section(4, 'Bad road surface at LCs', ok('badRoad') && S.badRoad.dataDate ? 'as on ' + dmy(S.badRoad.dataDate) : null)];
    if (!ok('badRoad')) s4.push(problem(S.badRoad));
    else { const bw = fit([1, 1, 1, 1.1], half), r = S.badRoad.row;
      s4.push(table(bw, [H(['Total LCs', 'Open > 15 days', 'Open > 20 days', 'TDC'], bw), Rw([r.total, r.d15, r.d20, r.tdc], bw, ['c', 'c', 'c', 'c'])])); }
    let s5 = null;
    if (S.cctv) {
      s5 = [section(5, 'CCTV recording check, NI LC gates', ok('cctv') ? dmy(S.cctv.dataDate) + (S.cctv.stale ? ', latest available' : '') : null)];
      if (!ok('cctv')) s5.push(problem(S.cctv));
      else if (!S.cctv.row) s5.push(nilLine('No entry for this AEN'));
      else { const c = S.cctv.row;
        s5.push(kv([['Checked since start of drive', val(c.checked)], ['Exceptions', val(c.exceptions)], ['Scheduled trains checked (last 3 days)', val(c.fixed3)],
          ['Random trains checked (last 3 days)', val(c.random3)], ['Exception on date', c.today || 'Nil'], ['CCTV not working', c.working || 'Nil'],
          ...(c.zeroRemark ? [['Remarks against zero figure', c.zeroRemark]] : [])], 3200, half)); }
    }
    let n = 5 + (s5 ? 1 : 0);

    /* ---------- 6. KPI and inspections ---------- */
    const s6 = [section(n, 'KPI shortfall and inspections due / overdue', ok('kpi') && S.kpi.generatedAt ? 'TMS Control, ' + S.kpi.generatedAt : 'from TMS Control')];
    if (!ok('kpi')) s6.push(problem(S.kpi));
    else {
      const U = S.kpi.units, uw = Math.floor((W - 600 - 3800) / U.length), kw = [600, 3800, ...U.map((_, i) => i === U.length - 1 ? W - 4400 - uw * (U.length - 1) : uw)];
      const list = (a, max) => !a || !a.length ? 'Nil' : (a.length + (a.length ? ': ' + a.slice(0, max).join(', ') + (a.length > max ? ' …' : '') : ''));
      const items = [
        ['KPI marks, last month (out of 82)', u => val(u.lastMonth)], ['KPI marks, this month – forecast (out of 82)', u => val(u.thisMonth)],
        ['KPIs with shortfall', u => u.shortfalls.length ? u.shortfalls.join(', ') : 'Nil'],
        ['Inspections due this month (TMS plan)', u => list(u.due, 6)], ['Inspections planned earlier, not done', u => list(u.lateFromPlan, 6)],
        ['Inspections overdue', u => u.overdueCount ? u.overdueCount + ': ' + u.overdue.slice(0, 6).join(', ') + (u.overdue.length > 6 ? ' …' : '') : 'Nil'],
        ['Foot plate inspection – km not covered', u => val(u.footPlate)], ['Trolley inspection – km not covered', u => val(u.trolley)],
        ['Defects crossing 10 days', u => u.defects ? u.defects + ': ' + u.defectPlaces.slice(0, 4).join(', ') : 'Nil']];
      s6.push(table(kw, [H(['Sl.', 'Item', ...U.map(u => 'SSE/P.Way/' + u.label)], kw),
        ...items.map(([t, f], i) => Rw([String(i + 1), t, ...U.map(f)], kw, ['c', 'l', ...U.map(() => 'c')]))]));
    }
    n++;

    /* ---------- 7. blocks ---------- */
    const s7 = [section(n, 'Blocks availed', 'day of ' + dmy(R.period.day) + ' and night of ' + R.period.night.split('/').map(d => d.slice(8, 10)).join('/') + '.' + D.slice(5, 7) + '.' + D.slice(0, 4))];
    if (!ok('blocks')) s7.push(problem(S.blocks));
    else {
      const bw = [450, 1300, 2700, 2300, 950, W - 7700], rows = S.blocks.rows;
      const tot = rows.reduce((a, r) => a + (r.availed || 0), 0);
      const naText = 'Blocks not availed by Engg.: ' + (S.blocks.notAvailedDay === null ? '–' : S.blocks.notAvailedDay) + ' on ' + dmy(R.period.day) + '; ' + S.blocks.notAvailedMonth + ' this month';
      s7.push(table(bw, [H(['Sl.', 'Day / Night', 'Location', 'Nature of work', 'Block availed (min)', 'Progress'], bw),
        ...(rows.length ? rows.map((r, i) => Rw([String(i + 1), r.when, r.location, r.work, r.availed, r.progress], bw, ['c', 'c', 'l', 'l', 'c', 'l']))
          : [TR({ children: [cell('No block availed', W, { span: 6, align: C })] })]),
        TR({ children: [cell('Total', 6750, { span: 4, bold: true, align: RT }), cell(String(tot), 950, { bold: true, align: C }), cell(naText, W - 7700)] })]));
      const notes = [staleNote(S.blocks.dayDate, S.blocks.dayStale) && 'Day blocks: ' + staleNote(S.blocks.dayDate, S.blocks.dayStale),
        staleNote(S.blocks.nightDate, S.blocks.nightStale) && 'Night blocks: ' + staleNote(S.blocks.nightDate, S.blocks.nightStale)].filter(Boolean);
      if (notes.length) s7.push(para([run('Note: ' + notes.join('; ') + '.', { italics: true, size: 17 })], { before: 20 }));
    }
    n++;

    /* ---------- 8. machines ---------- */
    const s8 = [section(n, 'Machine utilisation', 'block from meeting sheet, effective time from TMS')];
    if (!ok('machines')) s8.push(problem(S.machines));
    else {
      const M = S.machines, mw = [450, 1150, 1500, 1500, 850, 850, 800, 850, 800, W - 8750];
      s8.push(table(mw, [H(['Sl.', 'Day / Night', 'Machine', 'Location', 'Corridor (min)', 'Block availed (min)', 'Block util.', 'Eff. time (min)', 'Eff. %', 'Progress'], mw, 17),
        ...(M.rows.length ? M.rows.map((r, i) => Rw([String(i + 1), r.when, r.machine, r.location, r.corridor, r.availed, r.check ? 'See note' : pct(r.util),
          r.awaited ? 'Awaited' : val(r.effective), pct(r.effPct), r.progress], mw, ['c', 'c', 'l', 'l', 'c', 'c', 'c', 'c', 'c', 'l']))
          : [TR({ children: [cell('No machine worked', W, { span: 10, align: C })] })]),
        TR({ children: [cell('Total', 4600, { span: 4, bold: true, align: RT }), ...['–', val(M.total.availed), pct(M.total.util), val(M.total.effective), pct(M.total.effPct)]
          .map((v, i) => cell(v, mw[4 + i], { bold: true, align: C })), cell('', mw[9])] })]));
      s8.push(para([run('Machine-days lost: ', { bold: true, size: 18 }), run(M.lost.length ? M.lost.map(l => l.machine + ' – ' + l.reason + (l.remark ? ' (' + l.remark + ')' : '')).join('; ') + '.' : 'Nil.', { size: 18 })], { before: 50 }));
      const notes = [];
      if (M.rows.some(r => r.check)) notes.push('“See note”: block availed exceeds corridor margin; excluded from utilisation');
      if (M.rows.some(r => r.awaited)) notes.push('effective time for night blocks awaited from TMS');
      if (notes.length) s8.push(para([run('Note: ', { bold: true, size: 17 }), run(notes.join('; ') + '.', { italics: true, size: 17 })], { before: 20 }));
    }
    n++;

    /* ---------- 9. TMS items ---------- */
    const s9 = [section(n, 'TMS items', ok('tms') ? 'position of ' + dmy(S.tms.curDate) + (S.tms.prevDate ? ' against ' + dmy(S.tms.prevDate) : '') + (S.tms.stale ? ', latest available' : '') : null)];
    if (!ok('tms')) s9.push(problem(S.tms));
    else {
      const rowsT = []; let last = null;
      S.tms.rows.forEach(t => {
        const first = t.sn !== last;
        const label = first ? t.item + (t.sub ? ' – ' + t.sub : '') : '      ' + (t.sub || t.item);
        const c = (typeof t.prev === 'number' && typeof t.cur === 'number') ? Math.round((t.cur - t.prev) * 100) / 100 : null;
        rowsT.push([first ? String(t.sn) : ' ', label, val(t.prev), t.error ? 'n/a' : val(t.cur), c ? (c > 0 ? '+' : '–') + val(Math.abs(c)) : '–']);
        last = t.sn;
      });
      const pd = S.tms.prevDate ? dm(S.tms.prevDate) : 'Prev.', cd = dm(S.tms.curDate);
      const tw = [420, half - 420 - 2040, 680, 680, 680];
      const tmsTable = rs => table(tw, [H(['Sl.', 'Item', pd, cd, 'Diff.'], tw, 16), ...rs.map(r => Rw(r, tw, ['c', 'l', 'c', 'c', 'c'], { size: 16 }))]);
      const split = Math.ceil(rowsT.length / 2) + 1;
      s9.push(pair([tmsTable(rowsT.slice(0, split))], [tmsTable(rowsT.slice(split))]));
    }

    /* ---------- page filling: where the main report ends ---------- */
    // Walk the main report as Word would: rows that cannot split move to the next page (with the table's header row repeated).
    const flow = (els) => {
      let y = 0, pages = 1;
      const add = (h, canSplit) => {
        if (y + h <= PH) { y += h; return; }
        if (!canSplit) { pages++; y = h; return; }
        h -= PH - y; pages++; while (h > PH) { h -= PH; pages++; } y = h;
      };
      els.forEach(el => {
        if (el._rows) {
          const head = el._rows.filter(r => r._head).reduce((a, r) => a + r._h, 0);
          el._rows.forEach(r => { const before = pages; add(r._h, !r._cant); if (pages > before && head && !r._head) y += head; });
        } else add(paraH(el, W), false);
      });
      return { pages, y };
    };

    /* ---------- the two blank pages for Mini Control ----------
       Rule: if the page they would start on is 70% or more empty, they start right there and end on that page
       (as many rows as fit; more rows can be added in Word later). Otherwise they start on a fresh page, sized to fill it. */
    const SAFE = 900;                                        // spare space so Word's own layout never spills a row over
    const FIXED_CAL = 2300;                                  // measured: the fixed parts of a blank page come out this much shorter than estimated
    const MAST_H = mast._h + 3 * lineH(19);                  // letterhead block at the top of a fresh page
    const ww = [3500, 3500, W - 7000], WROW = 1250;
    const iw = [600, 5372, 1700, 1500, W - 9172], IROW = 560;
    const signLine = () => para([run('Mini Control on duty: ____________________      Signature: ____________________      Time: __________', { size: 19 })], { before: 240 });
    const blocks = [
      { title: 'WORKSITES AND SAFETY ARRANGEMENTS', sub: 'To be filled in by Mini Control', rowH: WROW, minRows: 3, extra: paraH(signLine(), W),
        minH: 950, build: (n, h) => [table(ww, [H(['Worksite details', 'Work details', 'Safety arrangement details'], ww, 21),
          ...Array.from({ length: n }, () => TR({ height: { value: h || WROW, rule: HeightRule.EXACT }, cantSplit: true, children: ww.map(w => cell([para('')], w)) }))]), signLine()] },
      { title: 'OTHER REMARKS / ISSUES PENDING', sub: null, rowH: IROW, minRows: 8, extra: 0,
        minH: 480, build: (n, h) => [table(iw, [H(['Sl.', 'Remarks / issues pending', 'Action by', 'Target date', 'Status'], iw, 20),
          ...Array.from({ length: n }, (_, i) => TR({ height: { value: h || IROW, rule: HeightRule.EXACT }, cantSplit: true,
            children: iw.map((w, j) => cell(j === 0 ? String(i + 1) : '', w, { align: C })) }))])] }
    ];
    const titleParas = (b, top) => [para([run(b.title, { bold: true, underline: true, size: 26 })], { align: C, before: top ? 120 : 360 }),
      ...(b.sub ? [para([run(b.sub, { size: 18, italics: true })], { align: C, before: 40 })] : [])];
    const titleH = (b, top) => titleParas(b, top).reduce((a, p) => a + paraH(p, W), 0) + 300 /* table header row */;

    const nightTxt = 'Work of ' + dmy(R.period.day) + ' (day) and night of ' + R.period.night.split('/').map(d => d.slice(8, 10)).join('/') + '.' + D.slice(5, 7) + '.' + D.slice(0, 4);
    const mainKids = [...letterhead('MORNING POSITION', nightTxt), abstract, pair(s1, s2), ...s3,
      s5 ? pair(s4, s5) : pair(s4, [para('')]), ...s6, ...s7, ...s8, ...s9];
    const mainFlow = flow(mainKids);
    const sections = [{ children: [...mainKids] }];
    let cur = sections[0], used = mainFlow.y;
    const layout = [];                                        // what was decided, for the generated-by line and tests
    // Where each blank part goes, decided to avoid a janky look rather than by a fixed percentage:
    //  - on the page in hand whenever a usable table fits in the space left (rows a little shorter if that is what it takes),
    //  - otherwise on a fresh page, sized to fill it. Each part ends on the page it starts on.
    blocks.forEach(b => {
      const avail = PH - used - titleH(b, false) - b.extra - SAFE;
      let n = Math.floor(avail / b.rowH), h = b.rowH;
      if (n < b.minRows) { n = Math.floor(avail / b.minH); h = n > 0 ? Math.min(b.rowH, Math.floor(avail / n)) : 0; }
      if (used > 0 && n >= b.minRows) {
        cur.children.push(...titleParas(b, false), ...b.build(n, h));
        layout.push(b.title + ': same page, ' + n + ' rows' + (h < b.rowH ? ' (rows ' + Math.round(h / 567 * 10) / 10 + ' cm)' : ''));
      } else {
        const full = Math.floor((PH - MAST_H - titleH(b, true) - b.extra - SAFE + FIXED_CAL) / b.rowH);
        cur = { children: [mast, ...titleParas(b, true), para([run('Units: ' + R.units.map(u => 'SSE/P.Way/' + u).join(', '), { size: 19 })], { before: 80, after: 60 }), ...b.build(full)] };
        sections.push(cur);
        layout.push(b.title + ': new page, ' + full + ' rows');
      }
      used = PH;                                              // each part ends on its page
    });

    const footer = new Footer({ children: [
      new Paragraph({ alignment: C, children: [
        run('Page ', { size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16 }),
        run(' of ', { size: 16 }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16 }),
        run('   ·   Generated ' + (R.generatedAt || '') + (R.generatedBy ? ' by ' + R.generatedBy : ''), { size: 14 })] }),
      new Paragraph({ alignment: C, children: [new TextRun({ text: CREDIT, font: FONT, size: 13, color: '808080' })] })] });
    const PAGE = { size: { width: 11906, height: 16838 }, margin: { top: 567, bottom: 680, left: 567, right: 567, footer: 240 } };
    if (opts.onLayout) opts.onLayout({ mainPages: mainFlow.pages, mainFill: mainFlow.y / PH, layout });

    return new Document({
      creator: 'AEN Morning Position', title: 'Morning Position ' + R.aen + ' ' + dmy(D),
      styles: { default: { document: { run: { font: FONT, size: SZ } } } },
      sections: sections.map(sx => ({ properties: { page: PAGE }, footers: { default: footer }, children: sx.children }))
    });
  }
  g.buildMorningDoc = buildMorningDoc;
  if (typeof module !== 'undefined') module.exports = { buildMorningDoc };
})(typeof globalThis !== 'undefined' ? globalThis : this);
