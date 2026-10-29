/*
 * Builds the Morning Position Word file from the report JSON the server returns.
 * Works in the browser (global `docx` from vendor/docx-9.8.1.iife.js) and in Node
 * (tests pass require('docx')). Layout as agreed: official A4, black ruled tables,
 * bold centred grey headers, numbers centred. The report ends where its last section ends.
 */
(function (g) {
  function buildMorningDoc(X, R, opts) {
    opts = opts || {};
    const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
      AlignmentType, VerticalAlign, Footer, PageNumber, TableLayoutType, HeightRule, UnderlineType, TabStopType, ImageRun, LineRuleType } = X;

    const CREDIT = 'Designed & developed by Anindya Nandy, SSE/P.Way/Planning, Howrah Division';
    const W = 10772, FONT = 'Arial', SZ = 19, GREY = 'D9D9D9';
    const C = AlignmentType.CENTER, L = AlignmentType.LEFT, RT = AlignmentType.RIGHT;
    const b = { style: BorderStyle.SINGLE, size: 6, color: '000000' }, nb = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const ALL = { top: b, bottom: b, left: b, right: b }, NONE = { top: nb, bottom: nb, left: nb, right: nb };
    const S = R.sections, D = R.date;

    // The whole report is composed at a given density (cell padding, space above headings); see the end of the file.
    const compose = DZ => {

      /* ---------- small helpers ---------- */
      const dmy = iso => iso ? iso.slice(8, 10) + '.' + iso.slice(5, 7) + '.' + iso.slice(0, 4) : '';
      const dm = iso => iso ? iso.slice(8, 10) + '.' + iso.slice(5, 7) : '';
      const addDays = (iso, n) => { if (!iso) return iso; const d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
      const pct = v => v === null || v === undefined ? '–' : Math.round(v * 100) + '%';
      const val = v => v === null || v === undefined || v === '' ? '–' : (typeof v === 'number' ? (Number.isInteger(v) ? String(v) : v.toFixed(2)) : String(v));
      const tag = (r, text, size, bold, imgH) => { r._t = String(text); r._s = size || SZ; r._b = !!bold; if (imgH) r._img = imgH; return r; };
      const STRETCH = opts.testStretch || 1;     // tests only: print text larger than the estimate expects, to prove nothing spills
      const run = (t, o = {}) => tag(new TextRun({ text: String(t), font: FONT, size: Math.round((o.size || SZ) * STRETCH), bold: !!o.bold, italics: !!o.italics,
        underline: o.underline ? { type: UnderlineType.SINGLE } : undefined }), t, o.size, o.bold);
      const para = (t, o = {}) => {
        const kids = (Array.isArray(t) ? t : [t]).map(x => typeof x === 'string' ? run(x, o) : x);
        const spacing = { before: o.before || 0, after: o.after || 0 };
        if (o.exact) { spacing.line = o.exact; spacing.lineRule = LineRuleType.EXACT; }
        const p = new Paragraph({ children: kids, alignment: o.align || L, spacing, keepNext: !!o.keepNext, tabStops: o.tabs, pageBreakBefore: !!o.pageBreakBefore });
        p._runs = kids; p._sp = (o.before || 0) + (o.after || 0); p._kn = !!o.keepNext; if (o.exact) p._h = o.exact + p._sp; return p;
      };

      /* ---------- page filling estimate (twips) ----------
         Word lays the pages out, not us, so we estimate: line height and average letter width from the font size.
         Used to choose where the pages end (below). */
      const PH = 16838 - 567 - 800;                        // usable A4 height inside margins (and the footer)
      const lineH = s => Math.round(s * 11.6);              // half-points -> twips, Arial single spacing
      const charW = (s, b) => s * (b ? 5.45 : 5.1);          // average Arial letter width in twips
      // Arial letter widths (per 1000 of the font size), for wrapping text word by word as Word does
      const AW = (() => { const w = {}; const set = (s, v) => [...s].forEach(c => w[c] = v);
        set('abcdeghknopqsuvxyz', 556); set('fijlt', 278); set('r', 333); set('mw', 833);
        set('ABCDEFGHKNOPQRSUVXYZ', 690); set('IJ', 300); set('LT', 611); set('MW', 900);
        set('0123456789', 556); set(' .,:;!|', 278); set('-–()[]/', 333); set('%&@', 900); return w; })();
      const textW = (s, size, bold) => { let u = 0; for (const c of s) u += AW[c] || (c.charCodeAt(0) > 0x900 ? 650 : 560); return u * size / 100 * (bold ? 1.07 : 1); };
      const paraH = (p, width) => {
        if (p && p._h !== undefined) return p._h;
        if (!p || !p._runs) return lineH(SZ);
        let img = 0, s = 0, lines = 1, x = 0;
        const room = Math.max(400, width);
        p._runs.forEach(r => {
          if (r._img) img = Math.max(img, r._img);
          if (!r._t) return;
          s = Math.max(s, r._s);
          r._t.split(/(\s+)/).forEach(tok => {
            if (!tok) return;
            const w = textW(tok, r._s, r._b);
            if (/^\s+$/.test(tok)) { x += w; return; }
            if (x + w <= room) { x += w; return; }
            if (x > 0) { lines++; x = 0; }
            while (w > room && x === 0) { lines += Math.ceil(w / room) - 1; x = w % room; return; }
            x = w;
          });
        });
        return Math.max(img, lines * lineH(s || SZ)) + p._sp;
      };
      const cell = (content, width, o = {}) => { const kids = Array.isArray(content) ? content : [para(content, o)]; const c = new TableCell({
        width: { size: width, type: WidthType.DXA }, borders: o.borders || ALL, columnSpan: o.span,
        shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
        margins: { top: DZ.pad, bottom: DZ.pad, left: 60, right: 60 }, verticalAlign: o.valign || VerticalAlign.CENTER,
        children: kids });
        c._h = 2 * DZ.pad + kids.reduce((a, k) => a + (k._h !== undefined ? k._h : paraH(k, width - 120)), 0);
        c._keep = () => Array.isArray(content) ? c : cell(content, width, Object.assign({}, o, { keepNext: true })); return c; };
      const TR = o => { const r = new TableRow(o); r._h = o.height && o.height.rule === HeightRule.EXACT ? o.height.value : Math.max(...o.children.map(c => c._h || lineH(SZ) + 40));
        r._head = !!o.tableHeader; r._cant = !!o.cantSplit;
        r._kept = () => { const k = TR(Object.assign({}, o, { children: o.children.map(c => c._keep ? c._keep() : c) })); k._keep = true; return k; }; return r; };
      // No table leaves a stray row: its heading and first three rows stay together, and its last two rows stay together
      // (Word's "keep with next" on the rows; flow() below treats such rows as one block).
      const table = (widths, rows, o = {}) => {
        const body = rows.filter(r => !r._head).length; let k = 0;
        if (o.keep !== false) rows = rows.map(r => { if (r._head) return r._kept(); const i = k++; return (i < 2 || i === body - 2) && i < body - 1 ? r._kept() : r; });
        const tb = new Table({ width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA },
        columnWidths: widths, layout: TableLayoutType.FIXED, rows }); tb._rows = rows; tb._h = rows.reduce((a, r) => a + r._h, 0); return tb; };
      const H = (labels, widths, size) => TR({ tableHeader: true, cantSplit: true,
        children: labels.map((l, i) => cell(l, widths[i], { fill: GREY, bold: true, align: C, size })) });
      const Rw = (vals, widths, aligns, o = {}) => TR({ cantSplit: true,
        children: vals.map((v, i) => v instanceof TableCell ? v : cell(val(v), widths[i], { align: aligns[i] === 'c' ? C : L, size: o.size, bold: o.bold })) });
      const fit = (parts, total) => { const s = parts.reduce((a, c) => a + c, 0); const out = parts.map(p => Math.floor(p * total / s)); out[out.length - 1] += total - out.reduce((a, c) => a + c, 0); return out; };

      // note: the date of the data used (always shown); late: said in bold when that date is older than expected
      const section = (no, title, note, late) => para([run(no + '.  ', { bold: true, size: 21 }), run(title, { bold: true, underline: true, size: 21 }),
        ...(note || late ? [run('  (' + (note || ''), { size: 17, italics: true }),
          ...(late ? [run((note ? ', ' : '') + 'older data: ' + late, { size: 17, italics: true, bold: true })] : []), run(')', { size: 17, italics: true })] : [])],
        { before: DZ.sec, after: 50, keepNext: true });
      const nightOf = (from, to) => from && to ? from.slice(8, 10) + '/' + dmy(to) : (to ? dmy(to) : '');
      const notYet = what => what + ' not yet in the sheet';
      const half = (W - 240) / 2;
      const pair = (left, right) => table([half, 240, half], [TR({ children: [
        cell(left, half, { borders: NONE, valign: VerticalAlign.TOP }), cell([para('')], 240, { borders: NONE }),
        cell(right, half, { borders: NONE, valign: VerticalAlign.TOP })] })], { keep: false });
      // (pair's height: the taller of its two sides, each a list of paragraphs and tables)
      const nilLine = (text) => para([run(text, { size: SZ })], { after: 20 });
      const problem = (s) => para([run('Could not be read: ', { bold: true, size: 18 }), run(s.error, { italics: true, size: 18 })], { after: 20 });
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
      const EX = h => ({ value: h, rule: HeightRule.EXACT });
      const mast = table(lw, [
        TR({ height: EX(480), children: [
          cell([Object.assign(new Paragraph({ children: [hrun('प्रातःकालीन स्थिति', { size: 26, bold: true }), new TextRun({ text: '\tMorning Position', font: FONT, size: 30, bold: true }),
            new TextRun({ text: '\tDate: ' + dmy(D), font: FONT, size: 20, bold: true })],
            tabStops: [{ type: TabStopType.CENTER, position: Math.round(W / 2) }, { type: TabStopType.RIGHT, position: W - 120 }] }), { _h: 420 })], W, { span: 3, borders: NONE, valign: VerticalAlign.BOTTOM })] }),
        TR({ height: EX(1250), children: [
          cell([para([hrun('पूर्व रेलवे', { size: 56, bold: true })], { align: L })], lw[0], { borders: { top: rule(12), bottom: nb, left: nb, right: nb } }),
          cell([para([logoRun], { align: C })], lw[1], { borders: { top: rule(12), bottom: nb, left: nb, right: nb } }),
          cell([para([new TextRun({ text: 'EASTERN RAILWAY', font: 'Arial Narrow', size: 44, bold: true })], { align: RT })], lw[2], { borders: { top: rule(12), bottom: nb, left: nb, right: nb } })] }),
        TR({ height: EX(380), children: [
          cell([para([hrun(HINDI_OFFICE[R.aen] + ' का कार्यालय, हावड़ा मंडल', { size: 19 })], { align: C })], W, { span: 3, borders: { top: rule(6), bottom: nb, left: nb, right: nb } })] }),
        TR({ height: EX(360), children: [
          cell([para([run('Office of the ' + R.office + ', Howrah Division', { size: 19, bold: true })], { align: C })], W, { span: 3, borders: { top: nb, bottom: rule(12), left: nb, right: nb } })] })
      ], { keep: false });
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
      // † on a figure that comes from older data (the previous day's entry not yet in the sheet)
      const old = k => ok(k) && !!S[k].stale;
      const abstractCells = [
        ['Caution past TDC', nilOr(tdcN), old('tdcLapse')], ['Mini control exceptions', nilOr(excN), old('miniControl')],
        ['Failures (' + dm(R.period.day) + ')', nilOr(failN), old('major') || old('minor')],
        ['LCs with bad road', ok('badRoad') && S.badRoad.row ? val(S.badRoad.row.total) : '–', old('badRoad')],
        ['Blocks availed', blk ? blk.rows.length + ' (' + blk.rows.reduce((a, r) => a + (r.availed || 0), 0) + ' min)' : '–', !!blk && (blk.dayStale || blk.nightStale)],
        ['Blocks not availed', blk && blk.notAvailedDay !== null ? blk.notAvailedDay + ' (month: ' + blk.notAvailedMonth + ')' : '–', !!blk && !!blk.notAvailedStale],
        ['M/C block utilisation', mc ? pct(mc.total.util) : '–', !!mc && !!mc.nightStale], ['M/C days lost', mc ? String(mc.lost.length) : '–', false]]
        .map(([k, v, o]) => [k, o && v !== '–' ? v + ' †' : v, o]);
      const anyOld = abstractCells.some(c => c[2]);
      const aw = fit(Array(8).fill(1), W);
      const abstract = table(aw, [H(abstractCells.map(c => c[0]), aw, 17), Rw(abstractCells.map(c => c[1]), aw, Array(8).fill('c'), { bold: true })]);
      const abstractNote = anyOld ? [para([run('† older data: the entry for ' + dmy(R.period.day) + ' was not yet in the sheet; the section below gives the date used.', { italics: true, size: 16 })], { before: 30 })] : [];

      // Section numbers are fixed (docs/phase-2-spec.md §3), the same in every AEN's report and in the contents.
      const wantNight = nightOf(R.period.night.split('/')[0], D);
      const sub = (t) => para([run(t, { bold: true, size: 18 })], { before: 30, after: 30, keepNext: true });

      /* ---------- 1. mini control exceptions: AEN items; PWI items (deferred: no source yet) ---------- */
      const s1 = [section(1, 'Mini control exceptions', ok('miniControl') ? 'data of ' + dmy(S.miniControl.dataDate) : null,
        ok('miniControl') && S.miniControl.stale ? notYet(dmy(R.period.day)) : null)];
      const pwiMc = ['PWI mini control items', '—'];
      if (!ok('miniControl')) s1.push(problem(S.miniControl), kv([pwiMc], 3200, half));
      else if (!S.miniControl.rows.length) s1.push(kv([['AEN mini control items', 'Nil'], pwiMc], 3200, half));
      else { const w = [1800, half - 1800]; s1.push(sub('AEN mini control items'), table(w, [H(['Item', 'Details of exception'], w), ...S.miniControl.rows.map(r => Rw([r.item, r.detail], w, ['l', 'l']))]),
        para('', { exact: 60 }), kv([pwiMc], 3200, half)); }

      /* ---------- 2. winter patrolling (GPS) ---------- */
      const gOn = ok('gps') && S.gps.started && S.gps.row;
      const s2 = [section(2, 'Winter patrolling', 'night of ' + (gOn ? nightOf(addDays(S.gps.dataDate, -1), S.gps.dataDate) : wantNight),
        gOn && S.gps.stale ? notYet('night ' + wantNight) : null)];
      if (!ok('gps')) s2.push(problem(S.gps));
      else if (!gOn) s2.push(kv([['Patrol batches (GPS)', 'Season not started']], 3200, half));
      else { const g2 = S.gps.row;
        s2.push(kv([['Patrol batches (GPS)', val(g2.total)], ['Defaulter batches', val(g2.defaulters)],
          ...g2.kinds.map(k => [k.name, val(k.n)]), ...(g2.action ? [['Action taken by AEN', g2.action]] : [])], 3200, half)); }

      /* ---------- 3. failures ---------- */
      const fd = ['major', 'minor'].filter(ok).map(k => S[k]), fLate = fd.some(x => x.stale);
      const s3 = [section(3, 'Failures', fd.length ? 'reported on ' + [...new Set(fd.map(x => dmy(x.dataDate)))].join(' / ') : null, fLate ? notYet(dmy(R.period.day)) : null)];
      const failTable = (label, s, cols) => {
        if (!s.ok) return [problem(s)];
        if (!s.rows.length) return [kv([[label, 'Nil']], 3200, half)];
        const w = fit([1, 1.6, 2.4], half);
        return [para([run(label, { bold: true, size: 18 })], { before: 30 }), table(w, [H(['Station / section', 'Line / km', 'Nature of failure'], w, 17),
          ...s.rows.map(r => Rw(cols(r), w, ['l', 'c', 'l'], { size: 17 }))])];
      };
      s3.push(...failTable('Major asset failure', S.major, r => [r.section, r.line, r.type + (r.duration ? ' (' + r.duration + ')' : '')]));
      s3.push(...failTable('Minor failure', S.minor, r => [r.station, [r.line, r.km].filter(Boolean).join(', '), r.type]));

      /* ---------- 4. cautions ---------- */
      const cN = [ok('tdcLapse') ? 'TDC lapse ' + dmy(S.tdcLapse.dataDate) : null, ok('pqrsTdc') ? 'PQRS TDC fail as listed' : null,
        ok('protection') ? 'protection caution night of ' + nightOf(S.protection.nightOf, S.protection.dataDate) : null].filter(Boolean);
      const cLate = [ok('tdcLapse') && S.tdcLapse.stale ? 'TDC lapse ' + dmy(R.period.day) : null,
        ok('protection') && S.protection.stale ? 'protection caution night of ' + wantNight : null].filter(Boolean);
      const s4 = [section(4, 'Cautions in the section and TDC overdue', cN.join(' · '), cLate.length ? notYet(cLate.join(', ')) : null)];
      let anyCaution = false;
      if (!ok('tdcLapse')) s4.push(problem(S.tdcLapse));
      else if (S.tdcLapse.rows.length) {
        anyCaution = true;
        const cw = [400, 750, 1600, 1050, 2100, 1000, 700, W - 7600];
        s4.push(para([run('Caution TDC lapse', { bold: true, size: 18 })], { after: 30, keepNext: true }), table(cw, [
          H(['Sl.', 'Speed', 'Section / Line', 'Km', 'Nature of work', 'TDC', 'Days past TDC', 'Remarks'], cw),
          ...S.tdcLapse.rows.map((r, i) => Rw([String(i + 1), r.speed, r.section + ', ' + r.line, r.km, r.work, r.tdc ? dmy(r.tdc).slice(0, 6) + r.tdc.slice(2, 4) : '–', val(r.daysPast), r.remarks], cw,
            ['c', 'c', 'l', 'c', 'l', 'c', 'c', 'l']))]));
      }
      if (!ok('pqrsTdc')) s4.push(problem(S.pqrsTdc));
      else if (S.pqrsTdc.rows.length) {
        anyCaution = true;
        const pw = [400, 2000, 1500, 1100, 1100, 1100, 1300, W - 8500];
        s4.push(para([run('PQRS – 30 km/h not relaxed in time', { bold: true, size: 18 })], { before: 60, after: 30, keepNext: true }), table(pw, [
          H(['Sl.', 'Section / Line', 'Location', '30 km/h imposed', 'Relaxed on', '10th day', 'Final removal', 'Reason for delay'], pw, 17),
          ...S.pqrsTdc.rows.map((r, i) => Rw([String(i + 1), r.section + ', ' + r.line, r.location, r.imposed && r.imposed.length === 10 ? dmy(r.imposed) : r.imposed,
            r.relaxed, r.tenthDay, r.removal, r.removalReason || r.relaxReason], pw, ['c', 'l', 'c', 'c', 'c', 'c', 'c', 'l'], { size: 17 }))]));
      }
      if (!ok('protection')) s4.push(problem(S.protection));
      else if (S.protection.rows.length) {
        anyCaution = true;
        const vw = [400, 1500, 2600, 1300, 2400, W - 8200];
        s4.push(para([run('Protection caution for PQRS, BCM & M/R', { bold: true, size: 18 })], { before: 60, after: 30, keepNext: true }), table(vw, [
          H(['Sl.', 'PWI', 'Section', 'Line', 'Type of work', 'Caution on adjacent line'], vw, 17),
          ...S.protection.rows.map((r, i) => Rw([String(i + 1), r.pwi, r.section, r.line, r.work, r.imposed], vw, ['c', 'l', 'l', 'c', 'l', 'c'], { size: 17 }))]));
      }
      if (!anyCaution && ok('tdcLapse') && ok('pqrsTdc') && ok('protection')) s4.push(nilLine('Nil'));

      /* ---------- 5. bad road, 6. NI LC inspection, 7. CCTV ---------- */
      const s5 = [section(5, 'Bad road surface at LC gates', ok('badRoad') && S.badRoad.dataDate ? 'as on ' + dmy(S.badRoad.dataDate) : null,
        ok('badRoad') && S.badRoad.stale ? notYet('as on ' + dmy(R.period.day)) : null)];
      if (!ok('badRoad')) s5.push(problem(S.badRoad));
      else { const bw = fit([1, 1, 1, 1.1], half), r = S.badRoad.row;
        s5.push(table(bw, [H(['Total LCs', 'Open > 15 days', 'Open > 20 days', 'TDC'], bw), Rw([r.total, r.d15, r.d20, r.tdc], bw, ['c', 'c', 'c', 'c'])])); }
      // tabs 42 and 37 list only the AENs with NI LC gates; for the others both sections are one line
      const notHere = 'Not applicable: the sheet lists no NI LC gate under AEN/' + R.aen;
      const s6 = [section(6, 'NI LC inspection', ok('niLc') ? 'as on ' + dmy(S.niLc.dataDate) : null, ok('niLc') && S.niLc.stale ? notYet('as on ' + dmy(R.period.day)) : null)];
      if (!S.niLc) s6.push(nilLine(notHere));
      else if (!ok('niLc')) s6.push(problem(S.niLc));
      else if (!S.niLc.row) s6.push(nilLine('No entry for this AEN'));
      else { const q = S.niLc.row, cum = S.niLc.since ? ' since ' + S.niLc.since : ' so far';
        s6.push(kv([['NI LC gates', val(q.total)], ['Inspected within 12-day frequency', val(q.covered)], ['Gateman counselling done', val(q.counselled)],
          ['Due for schedule inspection', q.due ? val(q.due) + ' – overdue' : 'Nil'],
          ['Inspections: PWI / AEN / (Sr.)DEN', val(q.pwi) + ' / ' + val(q.aen) + ' / ' + val(q.den)],
          ['Total' + cum, val(q.pwiCum) + ' / ' + val(q.aenCum) + ' / ' + val(q.denCum)]], 3200, half)); }
      const s7 = [section(7, 'CCTV recording check', ok('cctv') ? 'data of ' + dmy(S.cctv.dataDate) : null, ok('cctv') && S.cctv.stale ? notYet(dmy(R.period.day)) : null)];
      if (!S.cctv) s7.push(nilLine(notHere));
      else if (!ok('cctv')) s7.push(problem(S.cctv));
      else if (!S.cctv.row) s7.push(nilLine('No entry for this AEN'));
      else { const c = S.cctv.row;
        s7.push(kv([['Checked since start of drive', val(c.checked)], ['Exceptions', val(c.exceptions)], ['Scheduled trains checked (last 3 days)', val(c.fixed3)],
          ['Random trains checked (last 3 days)', val(c.random3)], ['Exception on date', c.today || 'Nil'], ['CCTV not working', c.working || 'Nil'],
          ...(c.zeroRemark ? [['Remarks against zero figure', c.zeroRemark]] : [])], 3200, half)); }

      /* ---------- 11. KPI items ---------- */
      const s11 = [section(11, 'KPI items', ok('kpi') && S.kpi.generatedAt ? 'TMS Control, ' + S.kpi.generatedAt : 'from TMS Control')];
      if (!ok('kpi')) s11.push(problem(S.kpi));
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
        s11.push(table(kw, [H(['Sl.', 'Item', ...U.map(u => 'SSE/P.Way/' + u.label)], kw),
          ...items.map(([t, f], i) => Rw([String(i + 1), t, ...U.map(f)], kw, ['c', 'l', ...U.map(() => 'c')]))]));
      }
      /* ---------- 12. blocks ---------- */
      const bLate = ok('blocks') ? [S.blocks.dayStale ? 'day ' + dmy(R.period.day) : null, S.blocks.nightStale ? 'night ' + wantNight : null].filter(Boolean) : [];
      const s12 = [section(12, 'Blocks availed and output', ok('blocks') ? 'day of ' + dmy(S.blocks.dayDate) + ' and night of ' + nightOf(addDays(S.blocks.nightDate, -1), S.blocks.nightDate)
        : 'day of ' + dmy(R.period.day) + ' and night of ' + wantNight, bLate.length ? notYet(bLate.join(', ')) : null)];
      if (!ok('blocks')) s12.push(problem(S.blocks));
      else {
        const bw = [450, 1300, 2700, 2300, 950, W - 7700], rows = S.blocks.rows;
        const tot = rows.reduce((a, r) => a + (r.availed || 0), 0);
        const naText = 'Blocks not availed by Engg.: ' + (S.blocks.notAvailedDay === null ? '–' : S.blocks.notAvailedDay) + ' on ' + dmy(R.period.day) + '; ' + S.blocks.notAvailedMonth + ' this month';
        s12.push(table(bw, [H(['Sl.', 'Day / Night', 'Location', 'Nature of work', 'Block availed (min)', 'Progress'], bw),
          ...(rows.length ? rows.map((r, i) => Rw([String(i + 1), r.when, r.location, r.work, r.availed, r.progress], bw, ['c', 'c', 'l', 'l', 'c', 'l']))
            : [TR({ children: [cell('No block availed', W, { span: 6, align: C })] })]),
          TR({ children: [cell('Total', 6750, { span: 4, bold: true, align: RT }), cell(String(tot), 950, { bold: true, align: C }), cell(naText, W - 7700)] })]));
        if (S.blocks.notAvailedStale) s12.push(para([run('Note: blocks not availed are of ' + dmy(S.blocks.notAvailedDate) + ' (older data).', { italics: true, bold: true, size: 17 })], { before: 20 }));
      }
      /* ---------- 13. machines ---------- */
      const mLate = ok('machines') && S.machines.nightStale;
      const s13 = [section(13, 'Machine utilisation', (ok('machines') && S.machines.nightDate ? 'night of ' + nightOf(S.machines.nightOf, S.machines.nightDate) + '; ' : '') + 'block from meeting sheet, effective time from TMS',
        mLate ? notYet('night ' + wantNight) : null)];
      if (!ok('machines')) s13.push(problem(S.machines));
      else {
        const M = S.machines, mw = [450, 1150, 1500, 1500, 850, 850, 800, 850, 800, W - 8750];
        s13.push(table(mw, [H(['Sl.', 'Day / Night', 'Machine', 'Location', 'Corridor (min)', 'Block availed (min)', 'Block util.', 'Eff. time (min)', 'Eff. %', 'Progress'], mw, 17),
          ...(M.rows.length ? M.rows.map((r, i) => Rw([String(i + 1), r.when, r.machine, r.location, r.corridor, r.availed, r.check ? 'See note' : pct(r.util),
            r.awaited ? 'Awaited' : val(r.effective), pct(r.effPct), r.progress], mw, ['c', 'c', 'l', 'l', 'c', 'c', 'c', 'c', 'c', 'l']))
            : [TR({ children: [cell('No machine worked', W, { span: 10, align: C })] })]),
          TR({ children: [cell('Total', 4600, { span: 4, bold: true, align: RT }), ...['–', val(M.total.availed), pct(M.total.util), val(M.total.effective), pct(M.total.effPct)]
            .map((v, i) => cell(v, mw[4 + i], { bold: true, align: C })), cell('', mw[9])] })]));
        s13.push(para([run('Machine-days lost: ', { bold: true, size: 18 }), run(M.lost.length ? M.lost.map(l => l.machine + ' – ' + l.reason + (l.remark ? ' (' + l.remark + ')' : '')).join('; ') + '.' : 'Nil.', { size: 18 })], { before: 50 }));
        const notes = [];
        if (M.rows.some(r => r.check)) notes.push('“See note”: block availed exceeds corridor margin; excluded from utilisation');
        if (M.rows.some(r => r.awaited)) notes.push('effective time for night blocks awaited from TMS');
        if (notes.length) s13.push(para([run('Note: ', { bold: true, size: 17 }), run(notes.join('; ') + '.', { italics: true, size: 17 })], { before: 20 }));
      }
      /* ---------- 14. TMS dashboard ---------- */
      // read by the TMS reader early on the report day: the position up to the previous day
      const s14 = [section(14, 'TMS dashboard', ok('tms') ? 'position on the morning of ' + dmy(S.tms.curDate) + (S.tms.prevDate ? ', against the morning of ' + dmy(S.tms.prevDate) : '') : null, ok('tms') && S.tms.stale ? 'the TMS reader has not sent the reading of ' + dmy(D) : null)];
      if (!ok('tms')) s14.push(problem(S.tms));
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
        // one table, the two halves side by side in the same rows: it breaks across pages cleanly with one repeated heading
        const tw = [420, half - 420 - 2040, 680, 680, 680], gw = W - 2 * half;
        const cols = [...tw, gw, ...tw], al = ['c', 'l', 'c', 'c', 'c'];
        const split = Math.ceil(rowsT.length / 2) + 1, Lr = rowsT.slice(0, split), Rr = rowsT.slice(split);
        const gap = () => cell('', gw, { borders: { top: nb, bottom: nb, left: b, right: b } });
        const half5 = (r, hd) => r ? r.map((v, i) => cell(hd ? v : val(v), tw[i], hd ? { fill: GREY, bold: true, align: C, size: 16 } : { align: al[i] === 'c' ? C : L, size: 16 }))
          : tw.map(w => cell('', w, { borders: NONE }));
        s14.push(table(cols, [
          TR({ tableHeader: true, cantSplit: true, children: [...half5(['Sl.', 'Item', pd, cd, 'Diff.'], true), gap(), ...half5(['Sl.', 'Item', pd, cd, 'Diff.'], true)] }),
          ...Lr.map((r, i) => TR({ cantSplit: true, children: [...half5(r), gap(), ...half5(Rr[i])] }))]));
      }

      /* ---------- page filling: where the main report ends ---------- */
      // Walk the main report as Word would: rows that cannot split move to the next page (with the table's header row repeated).
      const flow = (els, y0, p0) => {
        let y = y0 || 0, pages = p0 || 1, carry = 0;          // carry: headings kept with what follows them
        const add = (h, canSplit) => {
          h += carry; carry = 0;
          if (y + h <= PH) { y += h; return; }
          if (!canSplit) { pages++; y = h; return; }
          h -= PH - y; pages++; while (h > PH) { h -= PH; pages++; } y = h;
        };
        els.forEach(el => {
          if (el._rows) {
            const rs = el._rows, head = rs.filter(r => r._head).reduce((a, r) => a + r._h, 0);
            for (let i = 0; i < rs.length;) {
              let j = i, h = rs[i]._h;
              while (rs[j]._keep && j + 1 < rs.length) { j++; h += rs[j]._h; }
              const before = pages; add(h, j === i && !rs[i]._cant); if (pages > before && head && !rs[i]._head) y += head;
              i = j + 1;
            }
          } else if (el._kn) carry += paraH(el, W);
          else add(paraH(el, W), false);
        });
        if (carry) add(0, false);
        return { pages, y };
      };

      /* ---------- where the pages end ----------
         The report ends where its last section ends. Word lays the pages out, not us, and the estimate above can be off
         by a few per cent of what has been laid out since the last fresh page (checked with tests/measure.js). Rules: no
         page may be left near-empty, and no table may leave a stray row on a page of its own. A natural end that stays
         safe if Word runs a little longer or shorter is kept; otherwise one or two sections start on a fresh page, chosen
         so the page they leave is well filled and the last page is safe. A fresh page also resets the estimate's error. */
      const nightTxt = 'Work of ' + dmy(R.period.day) + ' (day) and night of ' + R.period.night.split('/').map(d => d.slice(8, 10)).join('/') + '.' + D.slice(5, 7) + '.' + D.slice(0, 4);
      const groups = [
        [...letterhead('MORNING POSITION', nightTxt), abstract, ...abstractNote, pair([...s1, ...s2], s3)],
        s4, [pair([...s5, ...s6], s7)], s11, s12, s13, s14];
      const ERR = h => 0.06 * h + 300;                       // twips Word may run longer or shorter, for h twips laid out
      const LAST_MIN = 0.30, LEFT_MIN = 0.45, MAX = 0.98;    // last page / page left by a fresh-page start: worst-case fill
      const plan = brk => {
        let st = { y: 0, pages: 1 }, since = 0, score = Infinity; const starts = [], left = [];
        groups.forEach((g, i) => {
          if (brk.includes(i)) {
            const f = st.y / PH, e = ERR(since) / PH;
            score = Math.min(score, f - e - LEFT_MIN, MAX - f - e); left.push(Math.round(f * 100));
            st = { y: 0, pages: st.pages + 1 }; since = 0;
          }
          const y0 = st; st = flow(g, st.y, st.pages); since += (st.pages - y0.pages) * PH + st.y - y0.y;
          starts.push(y0.y + 700 > PH ? y0.pages + 1 : y0.pages);      // a heading low on a page goes over with its section
        });
        const f = st.y / PH, e = ERR(since) / PH;
        return { brk, left, pages: st.pages, end: f, err: e, score: Math.min(score, f - e - LAST_MIN, MAX - f - e), starts,
          emptiest: Math.min(f, ...left.map(x => x / 100)) };          // among safe layouts of equal cost: the fullest pages win
      };
      const cands = [[]];
      for (let i = 1; i < groups.length; i++) { cands.push([i]); for (let j = i + 1; j < groups.length; j++) cands.push([i, j]); }
      const plans = cands.map(plan);
      const footer = new Footer({ children: [
        new Paragraph({ alignment: C, children: [
          run('Page ', { size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16 }),
          run(' of ', { size: 16 }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16 }),
          run('   ·   Generated ' + (R.generatedAt || '') + (R.generatedBy ? ' by ' + R.generatedBy : ''), { size: 14 })] }),
        new Paragraph({ alignment: C, children: [new TextRun({ text: CREDIT, font: FONT, size: 13, color: '808080' })] })] });
      const build = p => groups.flatMap((g, i) => p.brk.includes(i) ? [para('', { exact: 20, pageBreakBefore: true }), ...g] : g);
      return { plans, build, footer };
    };

    /* ---------- choosing the layout ----------
       Normal density first; slightly tighter or looser spacing, or a section on a fresh page, only when the natural end
       is not safe. A fresh page costs more than a small change of spacing, as it leaves white space mid-report. */
    const DENSITY = [{ name: 'normal', pad: 20, sec: 110, cost: 0 }, { name: 'tight', pad: 12, sec: 80, cost: 1 }, { name: 'loose', pad: 30, sec: 140, cost: 1 },
      { name: 'tighter', pad: 5, sec: 60, cost: 2 }, { name: 'looser', pad: 42, sec: 170, cost: 2 }];
    let pick = null, fallback = null;
    (opts.density ? DENSITY.filter(d => d.name === opts.density) : DENSITY).forEach(dz => {
      const c = compose(dz);
      c.plans.forEach(p => {
        const cost = dz.cost + 1.5 * p.brk.length, cand = { c, p, dz, cost };
        if (p.score >= 0 && (!pick || cost < pick.cost || (cost === pick.cost && p.emptiest > pick.p.emptiest))) pick = cand;
        if (!fallback || p.score > fallback.p.score) fallback = cand;
      });
    });
    pick = pick || fallback;
    const best = pick.p, kids = pick.c.build(best), footer = pick.c.footer;
    const layout = ['report ends ' + Math.round(best.end * 100) + '% down page ' + best.pages + ' (±' + Math.round(best.err * 100) + '%)',
      pick.dz.name + ' spacing, ' + (best.brk.length ? 'fresh page before part ' + best.brk.join(', ') + ', leaving ' + best.left.map(x => x + '%').join(', ') : 'no forced page break'),
      best.score >= 0 ? 'safe' : 'CHECK: no safe page ending found'];
    const PAGE = { size: { width: 11906, height: 16838 }, margin: { top: 567, bottom: 800, left: 567, right: 567, footer: 240 } };
    if (opts.onLayout) opts.onLayout({ density: pick.dz.name, pages: best.pages, end: best.end, err: best.err, score: best.score, breaks: best.brk, left: best.left, starts: best.starts, layout });

    return new Document({
      creator: 'AEN Morning Position', title: 'Morning Position ' + R.aen + ' ' + D.split('-').reverse().join('.'),
      styles: { default: { document: { run: { font: FONT, size: SZ } } } },
      sections: [{ properties: { page: PAGE }, footers: { default: footer }, children: kids }]
    });
  }
  g.buildMorningDoc = buildMorningDoc;
  if (typeof module !== 'undefined') module.exports = { buildMorningDoc };
})(typeof globalThis !== 'undefined' ? globalThis : this);
