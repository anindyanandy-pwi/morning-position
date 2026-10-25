/*
 * Builds the Morning Position Word file from the report JSON the server returns.
 * Works in the browser (global `docx` from vendor/docx-9.8.1.iife.js) and in Node
 * (tests pass require('docx')). Layout as agreed: official A4, black ruled tables,
 * bold centred grey headers, numbers centred, then the two blank pages for Mini Control.
 */
(function (g) {
  function buildMorningDoc(X, R) {
    const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
      AlignmentType, VerticalAlign, Footer, PageNumber, TableLayoutType, HeightRule, UnderlineType, TabStopType } = X;

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
    const run = (t, o = {}) => new TextRun({ text: String(t), font: FONT, size: o.size || SZ, bold: !!o.bold, italics: !!o.italics,
      underline: o.underline ? { type: UnderlineType.SINGLE } : undefined });
    const para = (t, o = {}) => new Paragraph({ children: (Array.isArray(t) ? t : [t]).map(x => typeof x === 'string' ? run(x, o) : x),
      alignment: o.align || L, spacing: { before: o.before || 0, after: o.after || 0 }, keepNext: !!o.keepNext, tabStops: o.tabs });
    const cell = (content, width, o = {}) => new TableCell({
      width: { size: width, type: WidthType.DXA }, borders: o.borders || ALL, columnSpan: o.span,
      shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
      margins: { top: 20, bottom: 20, left: 60, right: 60 }, verticalAlign: o.valign || VerticalAlign.CENTER,
      children: Array.isArray(content) ? content : [para(content, o)] });
    const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA },
      columnWidths: widths, layout: TableLayoutType.FIXED, rows });
    const H = (labels, widths, size) => new TableRow({ tableHeader: true, cantSplit: true,
      children: labels.map((l, i) => cell(l, widths[i], { fill: GREY, bold: true, align: C, size })) });
    const Rw = (vals, widths, aligns, o = {}) => new TableRow({ cantSplit: true,
      children: vals.map((v, i) => v instanceof TableCell ? v : cell(val(v), widths[i], { align: aligns[i] === 'c' ? C : L, size: o.size, bold: o.bold })) });
    const fit = (parts, total) => { const s = parts.reduce((a, c) => a + c, 0); const out = parts.map(p => Math.floor(p * total / s)); out[out.length - 1] += total - out.reduce((a, c) => a + c, 0); return out; };

    const section = (no, title, note) => para([run(no + '.  ', { bold: true, size: 21 }), run(title, { bold: true, underline: true, size: 21 }),
      ...(note ? [run('  (' + note + ')', { size: 17, italics: true })] : [])], { before: 110, after: 50, keepNext: true });
    const half = (W - 240) / 2;
    const pair = (left, right) => table([half, 240, half], [new TableRow({ children: [
      cell(left, half, { borders: NONE, valign: VerticalAlign.TOP }), cell([para('')], 240, { borders: NONE }),
      cell(right, half, { borders: NONE, valign: VerticalAlign.TOP })] })]);
    const nilLine = (text) => para([run(text, { size: SZ })], { after: 20 });
    const problem = (s) => para([run('Could not be read: ', { bold: true, size: 18 }), run(s.error, { italics: true, size: 18 })], { after: 20 });
    const staleNote = (iso, stale) => stale && iso ? 'data of ' + dmy(iso) + ', latest available' : null;
    const kv = (rows, w1, width) => { const w = [w1, width - w1]; return table(w, rows.map(([k, v]) => Rw([k, v], w, ['l', 'c']))); };

    /* ---------- heading ---------- */
    const letterhead = (title, sub) => [
      para([run('EASTERN RAILWAY', { bold: true, size: 28 })], { align: C }),
      para([run('Office of the ' + R.office + ', Howrah Division', { size: 20 })], { align: C, before: 20 }),
      para([run(title, { bold: true, underline: true, size: 26 })], { align: C, before: 120 }),
      ...(sub ? [para([run(sub, { size: 18, italics: true })], { align: C, before: 20 })] : []),
      para([run('Units: ' + R.units.map(u => 'SSE/P.Way/' + u).join(', '), { size: 19 }), run('\tDate: ' + dmy(D), { size: 19, bold: true })],
        { before: 120, after: 60, tabs: [{ type: TabStopType.RIGHT, position: W }] })
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
          : [new TableRow({ children: [cell('No block availed', W, { span: 6, align: C })] })]),
        new TableRow({ children: [cell('Total', 6750, { span: 4, bold: true, align: RT }), cell(String(tot), 950, { bold: true, align: C }), cell(naText, W - 7700)] })]));
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
          : [new TableRow({ children: [cell('No machine worked', W, { span: 10, align: C })] })]),
        new TableRow({ children: [cell('Total', 4600, { span: 4, bold: true, align: RT }), ...['–', val(M.total.availed), pct(M.total.util), val(M.total.effective), pct(M.total.effPct)]
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

    /* ---------- blank pages ---------- */
    const ww = [3500, 3500, W - 7000];
    const worksite = table(ww, [H(['Worksite details', 'Work details', 'Safety arrangement details'], ww, 21),
      ...Array.from({ length: 10 }, () => new TableRow({ height: { value: 1250, rule: HeightRule.EXACT }, cantSplit: true, children: ww.map(w => cell([para('')], w)) }))]);
    const iw = [600, 5372, 1700, 1500, W - 9172];
    const issues = table(iw, [H(['Sl.', 'Remarks / issues pending', 'Action by', 'Target date', 'Status'], iw, 20),
      ...Array.from({ length: 22 }, (_, i) => new TableRow({ height: { value: 560, rule: HeightRule.EXACT }, cantSplit: true,
        children: iw.map((w, j) => cell(j === 0 ? String(i + 1) : '', w, { align: C })) }))]);

    const footer = new Footer({ children: [new Paragraph({ alignment: C, children: [
      run('Page ', { size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16 }),
      run(' of ', { size: 16 }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16 }),
      run('   ·   Generated ' + (R.generatedAt || '') + (R.generatedBy ? ' by ' + R.generatedBy : ''), { size: 14 })] })] });
    const PAGE = { size: { width: 11906, height: 16838 }, margin: { top: 567, bottom: 567, left: 567, right: 567, footer: 280 } };
    const nightTxt = 'Work of ' + dmy(R.period.day) + ' (day) and night of ' + R.period.night.split('/').map(d => d.slice(8, 10)).join('/') + '.' + D.slice(5, 7) + '.' + D.slice(0, 4);

    return new Document({
      creator: 'AEN Morning Position', title: 'Morning Position ' + R.aen + ' ' + dmy(D),
      styles: { default: { document: { run: { font: FONT, size: SZ } } } },
      sections: [
        { properties: { page: PAGE }, footers: { default: footer }, children: [
          ...letterhead('MORNING POSITION', nightTxt), abstract, pair(s1, s2), ...s3,
          s5 ? pair(s4, s5) : pair(s4, [para('')]), ...s6, ...s7, ...s8, ...s9] },
        { properties: { page: PAGE }, footers: { default: footer }, children: [
          ...letterhead('WORKSITES AND SAFETY ARRANGEMENTS', 'To be filled in by Mini Control'), worksite,
          para([run('Mini Control on duty: ____________________      Signature: ____________________      Time: __________', { size: 19 })], { before: 240 })] },
        { properties: { page: PAGE }, footers: { default: footer }, children: [...letterhead('OTHER REMARKS / ISSUES PENDING'), issues] }
      ]
    });
  }
  g.buildMorningDoc = buildMorningDoc;
  if (typeof module !== 'undefined') module.exports = { buildMorningDoc };
})(typeof globalThis !== 'undefined' ? globalThis : this);
