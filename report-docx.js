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
      AlignmentType, VerticalAlign, Footer, PageNumber, TableLayoutType, HeightRule, UnderlineType, TabStopType, ImageRun, LineRuleType,
      HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType, LeaderType } = X;

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
        p._runs = kids; p._sp = (o.before || 0) + (o.after || 0); p._kn = !!o.keepNext; if (o.sec) p._sec = o.sec; if (o.exact) p._h = o.exact + p._sp; return p;
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
      // a cell's text keeps the line breaks it has in the sheet: one paragraph per line
      const cell = (content, width, o = {}) => { const kids = Array.isArray(content) ? content : String(content).split('\n').map(t => para(t, o)); const c = new TableCell({
        width: { size: width, type: WidthType.DXA }, borders: o.borders || ALL, columnSpan: o.span,
        shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
        margins: { top: DZ.pad, bottom: DZ.pad, left: 60, right: 60 }, verticalAlign: o.valign || VerticalAlign.CENTER,
        children: kids });
        c._h = 2 * DZ.pad + kids.reduce((a, k) => a + (k._h !== undefined ? k._h : paraH(k, width - 120)), 0);
        c._keep = () => Array.isArray(content) ? c : cell(content, width, Object.assign({}, o, { keepNext: true })); return c; };
      const TR = o => { const r = new TableRow(o); r._h = o.height && o.height.rule === HeightRule.EXACT ? o.height.value : Math.max(...o.children.map(c => c._h || lineH(SZ) + 40));
        r._head = !!o.tableHeader; r._cant = !!o.cantSplit;
        r._kept = () => { const k = TR(Object.assign({}, o, { children: o.children.map(c => c._keep ? c._keep() : c) })); k._keep = true; return k; }; return r; };
      // No table leaves a stray row: its heading and first three rows stay together, and so do its last three rows, so a
      // long table that breaks across pages keeps at least three rows on each side
      // (Word's "keep with next" on the rows; flow() below treats such rows as one block).
      const table = (widths, rows, o = {}) => {
        const body = rows.filter(r => !r._head).length; let k = 0;
        if (o.keep !== false) rows = rows.map(r => { if (r._head) return r._kept(); const i = k++; return (i < 2 || i >= body - 3) && i < body - 1 ? r._kept() : r; });
        const tb = new Table({ width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA },
        columnWidths: widths, layout: TableLayoutType.FIXED, rows }); tb._rows = rows; tb._w = widths.reduce((a, c) => a + c, 0); tb._h = rows.reduce((a, r) => a + r._h, 0); return tb; };
      const H = (labels, widths, size) => TR({ tableHeader: true, cantSplit: true,
        children: labels.map((l, i) => cell(l, widths[i], { fill: GREY, bold: true, align: C, size })) });
      const Rw = (vals, widths, aligns, o = {}) => TR({ cantSplit: true,
        children: vals.map((v, i) => v instanceof TableCell ? v : cell(val(v), widths[i], { align: aligns[i] === 'c' ? C : L, size: o.size, bold: o.bold })) });
      const fit = (parts, total) => { const s = parts.reduce((a, c) => a + c, 0); const out = parts.map(p => Math.floor(p * total / s)); out[out.length - 1] += total - out.reduce((a, c) => a + c, 0); return out; };

      // note: the date of the data used (always shown); late: said in bold when that date is older than expected
      const section = (no, title, note, late) => para([run(no + '.  ', { bold: true, size: 21 }), run(title, { bold: true, underline: true, size: 21 }),
        ...(note || late ? [run('  (' + (note || ''), { size: 17, italics: true }),
          ...(late ? [run((note ? ', ' : '') + 'older data: ' + late, { size: 17, italics: true, bold: true })] : []), run(')', { size: 17, italics: true })] : [])],
        { before: DZ.sec, after: 50, keepNext: true, sec: [no] });
      const nightOf = (from, to) => from && to ? from.slice(8, 10) + '/' + dmy(to) : (to ? dmy(to) : '');
      const notYet = what => what + ' not yet in the sheet';
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

      /* ======================= the sections, as data =======================
         Each section is described, not placed: a heading and a list of parts (tables, key–value tables, lines of text).
         The layout engine below decides from the content how wide each section is, which go side by side, and every
         column width. Columns: h heading, a 'c' centred, one: never wraps (numbers, dates, Sl.), free: free text
         (work, progress, remarks, reasons) that may take several lines and gets the width left over. */
      const T = (cols, rows, o = {}) => Object.assign({ t: 'table', cols, rows }, o);
      const KV = rows => ({ t: 'kv', rows });
      const NOTE = (runs, o = {}) => Object.assign({ t: 'note', runs }, o);       // runs: [[text, {bold, italics, size}]]
      const SUB = text => ({ t: 'sub', text });
      const ERRP = s => NOTE([['Could not be read: ', { bold: true, size: 18 }], [s.error, { italics: true, size: 18 }]]);
      const head = (no, title, note, late) => ({ no, title, note, late });
      const SECS = [];
      const add = (h, parts, o = {}) => SECS.push(Object.assign({ head: h, parts, secs: o.secs || [h.no] }, o));

      /* 1. mini control exceptions: AEN items; PWI items (deferred: no source yet) */
      {
        const pwi = ['PWI mini control items', '—'], mc = S.miniControl, parts = [];
        if (!ok('miniControl')) parts.push(ERRP(mc), KV([pwi]));
        else if (!mc.rows.length) parts.push(KV([['AEN mini control items', 'Nil'], pwi]));
        else parts.push(SUB('AEN mini control items'), T([{ h: 'Item' }, { h: 'Details of exception', free: true }], mc.rows.map(r => [r.item, r.detail])), KV([pwi]));
        add(head(1, 'Mini control exceptions', ok('miniControl') ? 'data of ' + dmy(mc.dataDate) : null, ok('miniControl') && mc.stale ? notYet(dmy(R.period.day)) : null), parts);
      }
      /* 2. winter patrolling (GPS) */
      {
        const gOn = ok('gps') && S.gps.started && S.gps.row, parts = [];
        if (!ok('gps')) parts.push(ERRP(S.gps));
        else if (!gOn) parts.push(KV([['Patrol batches (GPS)', 'Season not started']]));
        else { const g2 = S.gps.row; parts.push(KV([['Patrol batches (GPS)', val(g2.total)], ['Defaulter batches', val(g2.defaulters)],
          ...g2.kinds.map(k => [k.name, val(k.n)]), ...(g2.action ? [['Action taken by AEN', g2.action]] : [])])); }
        add(head(2, 'Winter patrolling', 'night of ' + (gOn ? nightOf(addDays(S.gps.dataDate, -1), S.gps.dataDate) : wantNight), gOn && S.gps.stale ? notYet('night ' + wantNight) : null), parts);
      }
      /* 3. failures */
      {
        const fd = ['major', 'minor'].filter(ok).map(k => S[k]), parts = [];
        const fail = (label, s, cols) => {
          if (!s.ok) return [ERRP(s)];
          if (!s.rows.length) return [KV([[label, 'Nil']])];
          return [SUB(label), T([{ h: 'Station / section' }, { h: 'Line / km', a: 'c' }, { h: 'Nature of failure', free: true }], s.rows.map(cols), { size: 17, hsize: 17 })];
        };
        parts.push(...fail('Major asset failure', S.major, r => [r.section, r.line, r.type + (r.duration ? ' (' + r.duration + ')' : '')]),
          ...fail('Minor failure', S.minor, r => [r.station, [r.line, r.km].filter(Boolean).join(', '), r.type]));
        add(head(3, 'Failures', fd.length ? 'reported on ' + [...new Set(fd.map(x => dmy(x.dataDate)))].join(' / ') : null, fd.some(x => x.stale) ? notYet(dmy(R.period.day)) : null), parts);
      }
      /* 4. cautions */
      {
        const cN = [ok('tdcLapse') ? 'TDC lapse ' + dmy(S.tdcLapse.dataDate) : null, ok('pqrsTdc') ? 'PQRS TDC fail as listed' : null,
          ok('protection') ? 'protection caution night of ' + nightOf(S.protection.nightOf, S.protection.dataDate) : null].filter(Boolean);
        const cLate = [ok('tdcLapse') && S.tdcLapse.stale ? 'TDC lapse ' + dmy(R.period.day) : null,
          ok('protection') && S.protection.stale ? 'protection caution night of ' + wantNight : null].filter(Boolean);
        const parts = []; let any = false;
        if (!ok('tdcLapse')) parts.push(ERRP(S.tdcLapse));
        else if (S.tdcLapse.rows.length) { any = true;
          parts.push(SUB('Caution TDC lapse'), T([{ h: 'Sl.', a: 'c', one: true }, { h: 'Speed', a: 'c', one: true }, { h: 'Section / Line' }, { h: 'Km', a: 'c' },
            { h: 'Nature of work', free: true }, { h: 'TDC', a: 'c', one: true }, { h: 'Days past TDC', a: 'c', one: true }, { h: 'Remarks', free: true }],
            S.tdcLapse.rows.map((r, i) => [String(i + 1), r.speed, r.section + ', ' + r.line, r.km, r.work, r.tdc ? dmy(r.tdc).slice(0, 6) + r.tdc.slice(2, 4) : '–', val(r.daysPast), r.remarks]))); }
        if (!ok('pqrsTdc')) parts.push(ERRP(S.pqrsTdc));
        else if (S.pqrsTdc.rows.length) { any = true;
          parts.push(SUB('PQRS – 30 km/h not relaxed in time'), T([{ h: 'Sl.', a: 'c', one: true }, { h: 'Section / Line' }, { h: 'Location', a: 'c' }, { h: '30 km/h imposed', a: 'c', one: true },
            { h: 'Relaxed on', a: 'c' }, { h: '10th day', a: 'c', one: true }, { h: 'Final removal', a: 'c' }, { h: 'Reason for delay', free: true }],
            S.pqrsTdc.rows.map((r, i) => [String(i + 1), r.section + ', ' + r.line, r.location, r.imposed && r.imposed.length === 10 ? dmy(r.imposed) : r.imposed,
              r.relaxed, r.tenthDay, r.removal, r.removalReason || r.relaxReason]), { size: 17, hsize: 17 })); }
        if (!ok('protection')) parts.push(ERRP(S.protection));
        else if (S.protection.rows.length) { any = true;
          parts.push(SUB('Protection caution for PQRS, BCM & M/R'), T([{ h: 'Sl.', a: 'c', one: true }, { h: 'PWI' }, { h: 'Section' }, { h: 'Line', a: 'c' },
            { h: 'Type of work', free: true }, { h: 'Caution on adjacent line', a: 'c' }], S.protection.rows.map((r, i) => [String(i + 1), r.pwi, r.section, r.line, r.work, r.imposed]), { size: 17, hsize: 17 })); }
        if (!any && ok('tdcLapse') && ok('pqrsTdc') && ok('protection')) parts.push(NOTE([['Nil']]));
        add(head(4, 'Cautions in the section and TDC overdue', cN.join(' · '), cLate.length ? notYet(cLate.join(', ')) : null), parts);
      }
      /* 5. bad road */
      {
        const parts = [];
        if (!ok('badRoad')) parts.push(ERRP(S.badRoad));
        else { const r = S.badRoad.row; parts.push(T([{ h: 'Total LCs', a: 'c', one: true }, { h: 'Open > 15 days', a: 'c', one: true }, { h: 'Open > 20 days', a: 'c', one: true }, { h: 'TDC', a: 'c' }],
          [[r.total, r.d15, r.d20, r.tdc]])); }
        add(head(5, 'Bad road surface at LC gates', ok('badRoad') && S.badRoad.dataDate ? 'as on ' + dmy(S.badRoad.dataDate) : null, ok('badRoad') && S.badRoad.stale ? notYet('as on ' + dmy(R.period.day)) : null), parts);
      }
      /* 6. NI LC inspection, 7. CCTV — tabs 42 and 37 list only the AENs with NI LC gates; for the others one line */
      if (!S.niLc && !S.cctv) add({ no: '6–7', line: 'Sections 6 and 7 not applicable: no NI LC gates under AEN/' + R.aen + '.' }, [], { secs: [6, 7] });
      else {
        const p6 = [];
        if (!S.niLc) p6.push(NOTE([['Not applicable to AEN/' + R.aen]]));
        else if (!ok('niLc')) p6.push(ERRP(S.niLc));
        else if (!S.niLc.row) p6.push(NOTE([['No entry for this AEN']]));
        else { const q = S.niLc.row, cum = S.niLc.since ? ' since ' + S.niLc.since : ' so far';
          p6.push(KV([['NI LC gates', val(q.total)], ['Inspected within 12-day frequency', val(q.covered)], ['Gateman counselling done', val(q.counselled)],
            ['Due for schedule inspection', q.due ? val(q.due) + ' – overdue' : 'Nil'], ['Inspections: PWI / AEN / (Sr.)DEN', val(q.pwi) + ' / ' + val(q.aen) + ' / ' + val(q.den)],
            ['Total' + cum, val(q.pwiCum) + ' / ' + val(q.aenCum) + ' / ' + val(q.denCum)]])); }
        add(head(6, 'NI LC inspection', ok('niLc') ? 'as on ' + dmy(S.niLc.dataDate) : null, ok('niLc') && S.niLc.stale ? notYet('as on ' + dmy(R.period.day)) : null), p6);
        const p7 = [];
        if (!S.cctv) p7.push(NOTE([['Not applicable to AEN/' + R.aen]]));
        else if (!ok('cctv')) p7.push(ERRP(S.cctv));
        else if (!S.cctv.row) p7.push(NOTE([['No entry for this AEN']]));
        else { const c = S.cctv.row;
          p7.push(KV([['Checked since start of drive', val(c.checked)], ['Exceptions', val(c.exceptions)], ['Scheduled trains checked (last 3 days)', val(c.fixed3)],
            ['Random trains checked (last 3 days)', val(c.random3)], ['Exception on date', c.today || 'Nil'], ['CCTV not working', c.working || 'Nil'],
            ...(c.zeroRemark ? [['Remarks against zero figure', c.zeroRemark]] : [])])); }
        add(head(7, 'CCTV recording check', ok('cctv') ? 'data of ' + dmy(S.cctv.dataDate) : null, ok('cctv') && S.cctv.stale ? notYet(dmy(R.period.day)) : null), p7);
      }
      /* 9. work planned: the plan for the report day itself, from the AEN's tab of the work plan sheet */
      if (S.workPlan) {
        const wp = S.workPlan, parts = [];
        if (!wp.ok) parts.push(ERRP(wp));
        else if (!wp.entered) parts.push(NOTE([['Work plan for ' + dmy(D) + ' not entered yet', { bold: true }]]));
        else {
          const cols = [{ h: 'Sl.', a: 'c', one: true }, { h: 'PWI unit' }, { h: 'Line', a: 'c' }, { h: 'Location' }, { h: 'Work', free: true }, { h: 'Station / section' },
            { h: 'Agency', free: true }, { h: 'Supervisor', free: true }, { h: 'Men', a: 'c', one: true }, { h: 'Infringing', a: 'c' }];
          const tbl = rs => T(cols, rs.map((r, i) => [String(i + 1), r.unit, r.line, r.location, r.work, r.station, r.agency, r.supervisor || '–', r.manpower || '–', r.infringing || '–']), { size: 17, hsize: 17 });
          [['Track-circuited area', x => x.tc === 'tc'], ['Non track-circuited area', x => x.tc !== 'tc']].forEach(([label, f]) => {
            const rs = wp.rows.filter(f);
            parts.push(...(rs.length ? [SUB(label + ' (' + rs.length + ' worksite' + (rs.length > 1 ? 's' : '') + ')'), tbl(rs)] : [KV([[label, 'Nil']])]));
          });
          const odd = wp.rows.filter(x => x.tc === 'unclear');
          if (odd.length) parts.push(NOTE([['Note: ', { bold: true, size: 17 }], ['"Whether track circuit" not clear for ' + odd.map(x => x.location || x.work).join('; ') + ' (entered as ' + [...new Set(odd.map(x => '"' + x.tcRaw + '"'))].join(', ') + '); listed as non track-circuited.', { italics: true, size: 17 }]], { before: 20 }));
        }
        add(head(9, 'Work planned for the next day', 'plan for ' + dmy(D) + (wp.ok && wp.inferredDate ? ', its date taken from the next day\'s heading' : '')), parts);
      }
      /* 11. KPI items */
      {
        const parts = [];
        if (!ok('kpi')) parts.push(ERRP(S.kpi));
        else {
          const U = S.kpi.units, list = (a, max) => !a || !a.length ? 'Nil' : a.length + ': ' + a.slice(0, max).join(', ') + (a.length > max ? ' …' : '');
          const items = [['KPI marks, last month (out of 82)', u => val(u.lastMonth)], ['KPI marks, this month – forecast (out of 82)', u => val(u.thisMonth)],
            ['KPIs with shortfall', u => u.shortfalls.length ? u.shortfalls.join(', ') : 'Nil'],
            ['Inspections due this month (TMS plan)', u => list(u.due, 6)], ['Inspections planned earlier, not done', u => list(u.lateFromPlan, 6)],
            ['Inspections overdue', u => u.overdueCount ? u.overdueCount + ': ' + u.overdue.slice(0, 6).join(', ') + (u.overdue.length > 6 ? ' …' : '') : 'Nil'],
            ['Foot plate inspection – km not covered', u => val(u.footPlate)], ['Trolley inspection – km not covered', u => val(u.trolley)],
            ['Defects crossing 10 days', u => u.defects ? u.defects + ': ' + u.defectPlaces.slice(0, 4).join(', ') : 'Nil']];
          parts.push(T([{ h: 'Sl.', a: 'c', one: true }, { h: 'Item' }, ...U.map(u => ({ h: 'SSE/P.Way/' + u.label, a: 'c' }))], items.map(([t, f], i) => [String(i + 1), t, ...U.map(f)])));
        }
        add(head(11, 'KPI items', ok('kpi') && S.kpi.generatedAt ? 'TMS Control, ' + S.kpi.generatedAt : 'from TMS Control'), parts);
      }
      /* 12. blocks */
      {
        const bLate = ok('blocks') ? [S.blocks.dayStale ? 'day ' + dmy(R.period.day) : null, S.blocks.nightStale ? 'night ' + wantNight : null].filter(Boolean) : [];
        const parts = [];
        if (!ok('blocks')) parts.push(ERRP(S.blocks));
        else {
          const rows = S.blocks.rows, tot = rows.reduce((a, r) => a + (r.availed || 0), 0);
          const naText = 'Blocks not availed by Engg.: ' + (S.blocks.notAvailedDay === null ? '–' : S.blocks.notAvailedDay) + ' on ' + dmy(R.period.day) + '; ' + S.blocks.notAvailedMonth + ' this month';
          parts.push(T([{ h: 'Sl.', a: 'c', one: true }, { h: 'Day / Night', a: 'c', one: true }, { h: 'Location' }, { h: 'Nature of work', free: true }, { h: 'Block availed (min)', a: 'c', one: true }, { h: 'Progress', free: true }],
            rows.length ? rows.map((r, i) => [String(i + 1), r.when, r.location, r.work, val(r.availed), r.progress]) : [[{ v: 'No block availed', span: 6, a: 'c' }]],
            { foot: [[{ v: 'Total', span: 4, bold: true, a: 'r' }, { v: String(tot), bold: true, a: 'c' }, { v: naText }]] }));
          if (S.blocks.notAvailedStale) parts.push(NOTE([['Note: blocks not availed are of ' + dmy(S.blocks.notAvailedDate) + ' (older data).', { italics: true, bold: true, size: 17 }]], { before: 20 }));
        }
        add(head(12, 'Blocks availed and output', ok('blocks') ? 'day of ' + dmy(S.blocks.dayDate) + ' and night of ' + nightOf(addDays(S.blocks.nightDate, -1), S.blocks.nightDate)
          : 'day of ' + dmy(R.period.day) + ' and night of ' + wantNight, bLate.length ? notYet(bLate.join(', ')) : null), parts);
      }
      /* 13. machines */
      {
        const parts = [];
        if (!ok('machines')) parts.push(ERRP(S.machines));
        else {
          const M = S.machines, one = h => ({ h, a: 'c', one: true });
          parts.push(T([one('Sl.'), one('Day / Night'), { h: 'Machine' }, { h: 'Location' }, one('Corridor (min)'), one('Block availed (min)'), one('Block util.'), one('Eff. time (min)'), one('Eff. %'), { h: 'Progress', free: true }],
            M.rows.length ? M.rows.map((r, i) => [String(i + 1), r.when, r.machine, r.location, val(r.corridor), val(r.availed), r.check ? 'See note' : pct(r.util),
              r.awaited ? 'Awaited' : val(r.effective), pct(r.effPct), r.progress]) : [[{ v: 'No machine worked', span: 10, a: 'c' }]],
            { size: 17, hsize: 17, foot: [[{ v: 'Total', span: 4, bold: true, a: 'r' }, ...['–', val(M.total.availed), pct(M.total.util), val(M.total.effective), pct(M.total.effPct)].map(v => ({ v, bold: true, a: 'c' })), { v: '' }]] }));
          parts.push(NOTE([['Machine-days lost: ', { bold: true, size: 18 }], [M.lost.length ? M.lost.map(l => l.machine + ' – ' + l.reason + (l.remark ? ' (' + l.remark + ')' : '')).join('; ') + '.' : 'Nil.', { size: 18 }]], { before: 50 }));
          const notes = [];
          if (M.rows.some(r => r.check)) notes.push('“See note”: block availed exceeds corridor margin; excluded from utilisation');
          if (M.rows.some(r => r.awaited)) notes.push('effective time for night blocks awaited from TMS');
          if (notes.length) parts.push(NOTE([['Note: ', { bold: true, size: 17 }], [notes.join('; ') + '.', { italics: true, size: 17 }]], { before: 20 }));
        }
        add(head(13, 'Machine utilisation', (ok('machines') && S.machines.nightDate ? 'night of ' + nightOf(S.machines.nightOf, S.machines.nightDate) + '; ' : '') + 'block from meeting sheet, effective time from TMS',
          ok('machines') && S.machines.nightStale ? notYet('night ' + wantNight) : null), parts);
      }
      /* 14. TMS dashboard: read by the TMS reader early on the report day, the position up to the previous day */
      {
        const parts = [];
        if (!ok('tms')) parts.push(ERRP(S.tms));
        else {
          const rowsT = []; let last = null;
          S.tms.rows.forEach(t => {
            const first = t.sn !== last, label = first ? t.item + (t.sub ? ' – ' + t.sub : '') : '      ' + (t.sub || t.item);
            const c = (typeof t.prev === 'number' && typeof t.cur === 'number') ? Math.round((t.cur - t.prev) * 100) / 100 : null;
            rowsT.push([first ? String(t.sn) : ' ', label, val(t.prev), t.error ? 'n/a' : val(t.cur), c ? (c > 0 ? '+' : '–') + val(Math.abs(c)) : '–']); last = t.sn;
          });
          const pd = S.tms.prevDate ? dm(S.tms.prevDate) : 'Prev.', cd = dm(S.tms.curDate);
          parts.push({ t: 'twin', cols: [{ h: 'Sl.', a: 'c', one: true }, { h: 'Item', free: true }, { h: pd, a: 'c', one: true }, { h: cd, a: 'c', one: true }, { h: 'Diff.', a: 'c', one: true }], rows: rowsT, size: 16 });
        }
        add(head(14, 'TMS dashboard', ok('tms') ? 'position on the morning of ' + dmy(S.tms.curDate) + (S.tms.prevDate ? ', against the morning of ' + dmy(S.tms.prevDate) : '') : null,
          ok('tms') && S.tms.stale ? 'the TMS reader has not sent the reading of ' + dmy(D) : null), parts);
      }

      /* ======================= the layout engine =======================
         1. Width class from the content: a section is narrow when its widest natural line (a heading or a cell, unwrapped)
            and the narrowest its columns can be both fit in a half column; anything else is wide.
         2. Rows: one wide section at full width, or two narrow sections side by side in equal columns, each exactly
            (text width − 8 mm) / 2, the table filling its column. Sections keep their order; a narrow section pairs only
            with the next one, when their heights are within 1.6× of each other; the pairing that gives the shortest
            report is taken.
         3. A narrow section that is not paired is laid out at full width, columns spread by rule 4.
         4. Column widths from the content: each column gets the width of its longest word (one-line columns: their
            longest line; others at least what keeps them within three lines), then the rest is shared in proportion to
            (natural − minimum), never beyond the natural width; what is still left goes to the free-text columns.
         The same letter widths measure the text for the widths, the wrapping and the page estimate. */
      // PADX: the cell's left and right margins (60 + 60) and 2 pt to spare, as Word wraps a line that only just fits
      const GUT = 454, COL = Math.floor((W - GUT) / 2), PADX = 160;
      const wrapN = (text, room, size, bold) => {          // lines a piece of text takes in `room` twips, word by word as Word wraps
        let lines = 1, x = 0; room = Math.max(200, room);
        String(text).split(/(\s+)/).forEach(tok => {
          if (!tok) return; const w = textW(tok, size, bold);
          if (/^\s+$/.test(tok)) { x += w; return; }
          if (x + w <= room) { x += w; return; }
          if (x > 0) { lines++; x = 0; }
          if (w > room) { lines += Math.ceil(w / room) - 1; x = w % room; } else x = w;
        });
        return lines;
      };
      const ln = v => String(v === null || v === undefined ? '' : v).split('\n');
      const natW = (v, s, bd) => Math.max(...ln(v).map(l => textW(l.trim(), s, bd))) + PADX;
      const minW = (v, s, bd) => Math.max(0, ...ln(v).flatMap(l => l.split(/\s+/)).map(w => textW(w, s, bd))) + PADX;
      const linesAt = (v, width, s, bd) => ln(v).reduce((a, l) => a + wrapN(l, width - PADX, s, bd), 0);
      const within = (v, n, s, bd) => {                     // the narrowest width that keeps v within n lines
        let lo = minW(v, s, bd), hi = Math.max(lo, natW(v, s, bd));
        if (linesAt(v, lo, s, bd) <= n) return lo;
        while (hi - lo > 20) { const m = (lo + hi) / 2; if (linesAt(v, m, s, bd) <= n) hi = m; else lo = m; }
        return Math.ceil(hi);
      };
      const sum = a => a.reduce((x, y) => x + y, 0);
      const plain = c => c !== null && typeof c === 'object' ? c : { v: c };
      // the widths of a table's columns for a total width (rule 4)
      const fitCols = (cols, rows, total, size, hsize) => {
        const body = rows.map(r => r.map(plain)).filter(r => r.every(c => !c.span || c.span === 1));
        const vals = j => body.map(r => val(r[j] ? r[j].v : ''));
        const hs = hsize || size;
        const mn = cols.map((c, j) => { const v = vals(j), h = c.h || '';
          let m = Math.max(minW(h, hs, true), ...v.map(x => minW(x, size)));
          if (c.one) m = Math.max(m, ...v.map(x => natW(x, size)));
          else if (!c.free) m = Math.max(m, within(h, 3, hs, true), ...v.map(x => within(x, 3, size)));
          return m; });
        const nat = cols.map((c, j) => Math.max(mn[j], (natW(c.h || '', hs, true) - PADX) / 2 + PADX, ...vals(j).map(x => natW(x, size))));
        let w = mn.slice(), rem = total - sum(w);
        if (rem < 0) w = w.map(x => x * total / sum(mn));
        else {
          const want = sum(nat.map((x, j) => x - w[j]));
          if (want <= rem) { w = nat.slice(); rem -= want; } else { w = w.map((x, j) => x + (nat[j] - x) * rem / want); rem = 0; }
          if (rem > 0) { const fr = cols.map((c, j) => c.free ? j : -1).filter(j => j >= 0), to = fr.length ? fr : cols.map((c, j) => j);
            const base = sum(to.map(j => nat[j])); to.forEach(j => { w[j] += rem * nat[j] / base; }); }
        }
        w = w.map(Math.floor); const big = w.indexOf(Math.max(...w)); w[big] += total - sum(w);
        return { w, mn, nat };
      };
      const kvCols = (rows, total) => {                      // key–value: label 65 % / value 35 % when both fit, else rule 4
        const cols = [{ h: '' }, { h: '', a: 'c' }], f = fitCols(cols, rows, total, SZ);
        const a = Math.round(total * 0.65);
        return f.nat[0] <= a && f.nat[1] <= total - a ? [a, total - a] : f.w;
      };
      const ALIGN = { c: C, r: RT, l: L };
      let viol = [];                                           // cells of a non-free column taking more than three lines (per rendering)
      const mkCell = (c, width, o) => cell(val(c.v), width, { align: ALIGN[c.a || o.a || 'l'], size: o.size, bold: c.bold || o.bold, span: c.span, fill: o.fill });
      const mkTable = (part, width, where) => {
        const size = part.size || SZ, hsize = part.hsize || size;
        if (part.t === 'kv') {
          const w = kvCols(part.rows, width);
          part.rows.forEach(([k, v]) => [[k, w[0]], [v, w[1]]].forEach(([x, ww]) => { if (linesAt(x, ww, SZ) > 3) viol.push(where + ': "' + String(x).slice(0, 30) + '"'); }));
          return table(w, part.rows.map(([k, v]) => TR({ cantSplit: true, children: [cell(val(k), w[0], { align: L }), cell(val(v), w[1], { align: C })] })));
        }
        const { w } = fitCols(part.cols, part.rows, width, size, hsize);
        const spanW = (j, n) => sum(w.slice(j, j + (n || 1)));
        const row = (r, o) => { let j = 0; return TR({ cantSplit: true, children: r.map(plain).map(c => { const cw = spanW(j, c.span); const cc = mkCell(c, cw, Object.assign({ a: part.cols[j].a }, o));
          if (!part.cols[j].free && !part.cols[j].one && (!c.span || c.span === 1) && linesAt(val(c.v), cw, size) > 3) viol.push(where + ': "' + val(c.v).slice(0, 30) + '"');
          j += c.span || 1; return cc; }) }); };
        return table(w, [H(part.cols.map(c => c.h), w, hsize), ...part.rows.map(r => row(r, { size })), ...(part.foot || []).map(r => row(r, { size }))]);
      };
      const mkTwin = (part, width) => {                       // TMS dashboard: the list in two halves side by side, one table
        const gw = 240, hw = Math.floor((width - gw) / 2), { w: tw } = fitCols(part.cols, part.rows, hw, part.size);
        const cols = [...tw, width - 2 * hw, ...tw], rowsT = part.rows;
        const split = Math.ceil(rowsT.length / 2) + 1, Lr = rowsT.slice(0, split), Rr = rowsT.slice(split);
        const gap = () => cell('', width - 2 * hw, { borders: { top: nb, bottom: nb, left: b, right: b } });
        const half5 = (r, hd) => r ? r.map((v, i) => cell(hd ? v : val(v), tw[i], hd ? { fill: GREY, bold: true, align: C, size: part.size } : { align: ALIGN[part.cols[i].a || 'l'], size: part.size }))
          : tw.map(x => cell('', x, { borders: NONE }));
        const hd = part.cols.map(c => c.h);
        return table(cols, [TR({ tableHeader: true, cantSplit: true, children: [...half5(hd, true), gap(), ...half5(hd, true)] }),
          ...Lr.map((r, i) => TR({ cantSplit: true, children: [...half5(r), gap(), ...half5(Rr[i])] }))]);
      };
      const headPara = h => h.line
        ? para([run(h.no + '.  ', { bold: true, size: 21 }), run(h.line, { size: 19 })], { before: DZ.sec, after: 50, sec: [6, 7] })
        : section(h.no, h.title, h.note, h.late);
      // a section laid out at a width: its heading and parts; an empty line between two tables keeps Word from joining them
      const render = (sp, width) => {
        const els = [headPara(sp.head)]; let lastTable = false; viol = els._viol = [];
        sp.parts.forEach(p => {
          if (p.t === 'note') { els.push(para(p.runs.map(([t, o]) => run(t, o || {})), { before: p.before || 0, after: 20 })); lastTable = false; return; }
          if (p.t === 'sub') { els.push(para([run(p.text, { bold: true, size: 18 })], { before: 30, after: 30, keepNext: true })); lastTable = false; return; }
          if (lastTable) els.push(para('', { exact: 60 }));
          els.push(p.t === 'twin' ? mkTwin(p, width) : mkTable(p, width, 'section ' + sp.head.no)); lastTable = true;
        });
        return els;
      };
      const heightOf = (els, width) => sum(els.map(e => e._h !== undefined ? e._h : paraH(e, width)));
      // rule 1: narrow or wide, from the content
      const narrowOK = sp => sp.parts.every(p => {
        if (p.t === 'twin') return false;
        if (p.t === 'note' || p.t === 'sub') return true;
        const size = p.size || SZ, hs = p.hsize || size;
        if (p.t === 'kv') { const f = fitCols([{ h: '' }, { h: '' }], p.rows, COL, SZ); return Math.max(...f.nat) <= COL && sum(f.mn) <= COL; }
        const f = fitCols(p.cols, p.rows, COL, size, hs);
        const widest = Math.max(...p.rows.map(r => r.map(plain)).flat().filter(c => !c.span || c.span === 1).map(c => natW(val(c.v), size)), ...p.cols.map(c => (natW(c.h || '', hs, true) - PADX) / 2 + PADX));
        return widest <= COL && sum(f.mn) <= COL;
      });
      SECS.forEach(sp => {
        sp.narrow = narrowOK(sp);
        sp.full = render(sp, W); sp.hFull = heightOf(sp.full, W);
        if (sp.narrow) { sp.half = render(sp, COL); sp.hHalf = heightOf(sp.half, COL); }
      });
      // rule 2: the pairing that gives the shortest report (heights within 1.6x); `no` lists pairs that must not be made
      const pack = no => {
        const n = SECS.length, best = Array(n + 1).fill(null); best[n] = { h: 0, rows: [] };
        for (let i = n - 1; i >= 0; i--) {
          const a = SECS[i], opts2 = [{ h: a.hFull + best[i + 1].h, rows: [{ kind: 'full', s: [a] }, ...best[i + 1].rows] }];
          const bb = SECS[i + 1];
          if (bb && a.narrow && bb.narrow && !no.includes(i)) {
            const hi = Math.max(a.hHalf, bb.hHalf), lo = Math.min(a.hHalf, bb.hHalf);
            if (hi <= 1.6 * lo) opts2.push({ h: hi + 40 + best[i + 2].h, rows: [{ kind: 'pair', s: [a, bb], at: i }, ...best[i + 2].rows] });
          }
          best[i] = opts2.sort((x, y) => x.h - y.h || y.rows.length - x.rows.length)[0];
        }
        return best[0].rows;
      };
      const rowEls = rw => {
        if (rw.kind === 'full') return rw.s[0].full;
        // two columns of exactly COL with the 8 mm gutter: the cells have no padding, so each table fills its column
        const pc = (els, w) => { const c = new TableCell({ width: { size: w, type: WidthType.DXA }, borders: NONE, margins: { top: 0, bottom: 0, left: 0, right: 0 },
          verticalAlign: VerticalAlign.TOP, children: els.length ? els : [new Paragraph('')] }); c._h = heightOf(els, w); return c; };
        const cA = pc(rw.s[0].half, COL), cB = pc(rw.s[1].half, COL), cG = pc([para('')], W - 2 * COL);
        const r = new TableRow({ cantSplit: true, children: [cA, cG, cB] }); r._h = Math.max(cA._h, cB._h); r._cant = true;
        const t = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [COL, W - 2 * COL, COL], layout: TableLayoutType.FIXED, rows: [r] });
        t._rows = [r]; t._h = r._h; t._secs = rw.s.flatMap(x => x.secs);
        return [para('', { exact: 40, keepNext: true }), t];
      };
      const nightTxt = 'Work of ' + dmy(R.period.day) + ' (day) and night of ' + R.period.night.split('/').map(d => d.slice(8, 10)).join('/') + '.' + D.slice(5, 7) + '.' + D.slice(0, 4);
      const head0 = [...letterhead('MORNING POSITION', nightTxt), abstract, ...abstractNote];
      const packed = pack([]);
      // the layouts tried: the best pairing, and each of its pairs undone (rule 3 then lays both at full width)
      const ARR = [{ name: 'best pairing', cost: 0, rows: packed },
        ...packed.filter(r => r.kind === 'pair').map(r => ({ name: 'sections ' + r.s.map(x => x.head.no).join(' and ') + ' not paired', cost: 1.5, rows: pack([r.at]) }))]
        .map(a => Object.assign(a, { g: a.rows.map((rw, i) => i ? rowEls(rw) : [...head0, ...rowEls(rw)]) }));

      /* ---------- page filling: where the main report ends ---------- */
      // Walk the main report as Word would: rows that cannot split move to the next page (with the table's header row repeated).
      // m: every height times m, to see the report as Word may lay it out a little longer or shorter
      const flow = (els, y0, p0, fills, secs, m) => {
        let y = y0 || 0, pages = p0 || 1, carry = 0, over = false, waiting = [];   // carry: headings kept with what follows them
        fills = fills || []; secs = secs || {}; m = m || 1;   // how full each finished page is (twips); page where each section starts
        const add = (h, canSplit, own) => {
          h = h * m + carry; carry = 0;
          const at = y + h <= PH || canSplit ? pages : pages + 1;
          [...waiting, ...(own || [])].forEach(n => { if (!(n in secs)) secs[n] = at; }); waiting = [];
          if (y + h <= PH) { y += h; return; }
          if (!canSplit) { if (h > PH) over = true; fills.push(y); pages++; y = h; return; }
          h -= PH - y; fills.push(PH); pages++; while (h > PH) { h -= PH; fills.push(PH); pages++; } y = h;
        };
        els.forEach(el => {
          if (el._rows) {
            const rs = el._rows, head = rs.filter(r => r._head).reduce((a, r) => a + r._h, 0);
            for (let i = 0; i < rs.length;) {
              let j = i, h = rs[i]._h;
              while (rs[j]._keep && j + 1 < rs.length) { j++; h += rs[j]._h; }
              const before = pages; add(h, j === i && !rs[i]._cant, i === 0 ? el._secs : null); if (pages > before && head && !rs[i]._head) y += head * m;
              i = j + 1;
            }
          } else if (el._kn) { carry += paraH(el, W) * m; if (el._sec) waiting.push(...el._sec); }
          else add(paraH(el, W), false, el._sec);
        });
        if (carry) add(0, false);
        return { pages, y, fills, over };
      };

      /* ---------- where the pages end ----------
         The report ends where its last section ends. Word lays the pages out, not us, and the estimate above can be off
         by a few per cent of what has been laid out since the last fresh page (checked with tests/measure.js). Rules: no
         page may be left near-empty, a page in the middle must not look hollow, no table leaves a stray row. Long tables
         split across pages between rows (header repeated, at least three rows each side); a side-by-side row moves as
         one. In order of preference: spacing a little tighter or looser; a pair laid out as two full-width rows; last,
         a section on a fresh page, only where the page it leaves is at least 70 % full. */
      const ERR = h => 0.045 * h + 300;                      // twips Word may run longer or shorter, for h twips laid out (measured: within 3 %)
      const plan = (arr, brk) => {
        const groups = arr.g;
        let st = { y: 0, pages: 1 }, since = 0; const starts = [], left = [], fills = [], secs = {};
        let worst = Infinity, over = false;
        groups.forEach((g, i) => {
          if (brk.includes(i)) {
            const f = st.y / PH, e = ERR(since) / PH;
            worst = Math.min(worst, f - 0.70, 0.98 - f - e); left.push(Math.round(f * 100));
            fills.push(st.y); st = { y: 0, pages: st.pages + 1 }; since = 0;
          }
          const y0 = st; st = flow(g, st.y, st.pages, fills, secs); over = over || st.over; since += (st.pages - y0.pages) * PH + st.y - y0.y;
          starts.push(y0.y + 700 > PH ? y0.pages + 1 : y0.pages);      // a heading low on a page goes over with its section
        });
        const f = st.y / PH, e = ERR(since) / PH;
        const pagesF = fills.map(x => Math.round(x / PH * 100));
        // the same layout as Word may make it, 4.5 % longer and shorter: a block of rows kept together near the foot of a
        // page may then go over to the next one. In both, the last page stays between 25 % and 96 % full and no page left by
        // a fresh-page start spills over.
        [1.045, 0.955].forEach(mm => {
          let t = { y: 0, pages: 1 };
          groups.forEach((g, i) => { if (brk.includes(i)) { worst = Math.min(worst, 0.99 - t.y / PH, t.y / PH - 0.55); t = { y: 0, pages: t.pages + 1 }; }
            t = flow(g, t.y, t.pages, [], {}, mm); });
          worst = Math.min(worst, t.y / PH - 0.25, 0.96 - t.y / PH);
        });
        // the last page: at least 45 % full as estimated, never near-empty or spilling a sliver in the worst case
        const lastOk = Math.min(f - 0.45, f - e - 0.25, 0.98 - f - e), lastLoose = Math.min(f - 0.35, f - e - 0.25, 0.98 - f - e);
        return { arr, brk, left, pages: st.pages, end: f, err: e, fills: pagesF, starts, secs,
          ok: !over && Math.min(worst, lastOk) >= 0, okLoose: !over && Math.min(worst, lastLoose) >= 0, score: Math.min(worst, lastLoose) - (over ? 9 : 0),
          middle: pagesF.length ? Math.min(...pagesF) : 100 };      // among layouts of equal cost: the fullest middle pages win
      };
      const cands = n => { const c = [[]]; for (let i = 1; i < n; i++) { c.push([i]); for (let j = i + 1; j < n; j++) c.push([i, j]); } return c; };
      const plans = ARR.flatMap(a => cands(a.g.length).map(b => plan(a, b)));
      // what the tests check: every row's tables fill its width, side-by-side heights, cells wrapping beyond three lines
      const rowInfo = rows => rows.map(rw => rw.kind === 'full'
        ? { kind: 'full', secs: rw.s[0].secs, box: W, tables: rw.s[0].full.filter(e => e._w).map(e => e._w), viol: rw.s[0].full._viol }
        : { kind: 'pair', secs: rw.s.flatMap(x => x.secs), box: COL, tables: rw.s.flatMap(x => x.half.filter(e => e._w).map(e => e._w)),
          heights: rw.s.map(x => x.hHalf), viol: rw.s.flatMap(x => x.half._viol) });
      const footer = new Footer({ children: [
        new Paragraph({ alignment: C, children: [
          run('Page ', { size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16 }),
          run(' of ', { size: 16 }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16 }),
          run('   ·   Generated ' + (R.generatedAt || '') + (R.generatedBy ? ' by ' + R.generatedBy : ''), { size: 14 })] }),
        new Paragraph({ alignment: C, children: [new TextRun({ text: CREDIT, font: FONT, size: 13, color: '808080' })] })] });
      const build = p => p.arr.g.flatMap((g, i) => p.brk.includes(i) ? [para('', { exact: 20, pageBreakBefore: true }), ...g] : g);
      return { plans, build, footer, rowInfo };
    };

    /* ---------- choosing the layout ----------
       Cheapest safe layout wins: spacing a little tighter or looser (0.5–2.5), tables split between rows (free), a
       side-by-side pair laid out as two full-width rows instead (1.5), a fresh page (10 each), and a page before the last under 85 % full (1 for every 4 %). Equal cost: the fullest middle pages, then the fuller last page. If nothing
       meets the 45 % last-page rule, a last page down to 35 % is accepted before anything worse. */
    const DENSITY = [{ name: 'normal', pad: 20, sec: 110, cost: 0 },
      { name: 'a little tight', pad: 15, sec: 90, cost: 0.5 }, { name: 'a little loose', pad: 25, sec: 125, cost: 0.5 },
      { name: 'tight', pad: 10, sec: 75, cost: 1 }, { name: 'loose', pad: 32, sec: 145, cost: 1 },
      { name: 'tighter', pad: 5, sec: 60, cost: 2 }, { name: 'looser', pad: 44, sec: 175, cost: 2 }, { name: 'loosest', pad: 58, sec: 210, cost: 2.5 }];
    const all = [];
    (opts.density ? DENSITY.filter(d => d.name === opts.density) : DENSITY).forEach(dz => {
      const c = compose(dz);
      c.plans.forEach(p => all.push({ c, p, dz, cost: dz.cost + p.arr.cost + 10 * p.brk.length + Math.max(0, 85 - p.middle) / 4 }));
    });
    const better = (x, y) => x.cost - y.cost || y.p.middle - x.p.middle || y.p.end - x.p.end;
    const pick = all.filter(x => x.p.ok).sort(better)[0] || all.filter(x => x.p.okLoose).sort(better)[0] || all.sort((x, y) => y.p.score - x.p.score)[0];
    const best = pick.p, kids = pick.c.build(best), footer = pick.c.footer, rowsUsed = pick.c.rowInfo(best.arr.rows);
    const layout = ['report ends ' + Math.round(best.end * 100) + '% down page ' + best.pages + ' (±' + Math.round(best.err * 100) + '%)',
      'pages ' + [...best.fills, Math.round(best.end * 100)].map(x => x + '%').join(' · '),
      pick.dz.name + ' spacing, ' + best.arr.name + ', ' + (best.brk.length ? 'fresh page before part ' + best.brk.join(', ') : 'no forced page break'),
      best.ok ? 'safe' : best.okLoose ? 'safe, last page under 45 %' : 'CHECK: no safe page ending found'];

    /* ---------- the top sheet (page 1) ----------
       docs/top-sheet-final.dc.html, drawn at 794 x 1123 px for A4: 1 px = 15 twips. The frame, the seal, the short rule
       and the CONTENTS rule are pictures rendered once (tools/render-cover-art.js); all text is Word text. Every line has
       an exact height, so the page is laid out the same in any Word, and the gaps are stretched or shrunk together so
       the contents reach the footer, as on the design. */
    const topSheet = secs => {
      const art = opts.cover || {}, px = v => Math.round(v * 15), INK = '141414', SOFT = '3B3B3B';
      const SERIF = 'Georgia', HIN = 'Nirmala UI';
      const CW = px(794 - 2 * 80.2), TOP = px(66.2), FOOT = px(1123 - 1068.8), BOT = FOOT + px(26);
      const AVAIL = 16838 - TOP - BOT;
      const t = (text, o = {}) => new TextRun({ text, font: o.font ? { ascii: o.font, hAnsi: o.font, cs: o.font } : FONT, size: Math.round((o.px || 14) * 1.5),
        sizeComplexScript: Math.round((o.px || 14) * 1.5), bold: !!o.bold, boldComplexScript: !!o.bold, italics: !!o.italic, color: o.color || INK,
        characterSpacing: o.track ? px((o.px || 14) * o.track) : undefined });
      const img = (data, w, h) => data ? new ImageRun({ type: 'png', data, transformation: { width: w, height: h } }) : new TextRun('');
      // the stack under the seal: [gap above (px), line height (px), runs]
      const name = ((R.office || '').match(/[,/]\s*([^,/(]+?)\s*\(/) || [, R.aen])[1].toUpperCase();
      const day = new Date(D + 'T00:00:00Z'), MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      const dateLong = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][day.getUTCDay()] + ', ' + day.getUTCDate() + ' ' + MONTHS[day.getUTCMonth()] + ' ' + day.getUTCFullYear();
      const d1 = R.period.day, dd = iso => iso.slice(8, 10) + '.' + iso.slice(5, 7);
      const stack = [
        [16, 16, [t('EASTERN RAILWAY · HOWRAH DIVISION', { px: 12.7, track: 0.32 })]],
        [2, 21, [t('पूर्व रेलवे · हावड़ा मंडल', { px: 15, font: HIN, color: SOFT })]],
        [12, 64, [t('Morning Position', { px: 57.3, font: SERIF, bold: true })]],
        [0, 30, [t('प्रातःकालीन स्थिति', { px: 20, font: HIN, bold: true })]],
        [14, 12, [img(art.rule, 400, 12)]],
        [14, 22, [t('AEN/' + R.aen + ' · ' + name, { px: 18, bold: true, track: 0.08 })]],
        [3, 18, [t('Units: ' + R.units.map(u => 'SSE/P.Way/' + u).join(', '), { px: 14, color: SOFT })]],
        [8, 30, [t(dateLong, { px: 24, font: SERIF })]],
        [2, 20, [t('the day of ' + dd(d1) + ' and the night of ' + d1.slice(8, 10) + '/' + dd(D), { px: 15, font: SERIF, italic: true, color: SOFT })]],
        [20, 16, [img(art.contents, 634, 16)]]];
      // the contents: the sections in this report, with the page each starts on (the top sheet is page 1)
      const has = n => n in secs, pg = n => String(secs[n] + 1);
      const ITEMS = [[1, 'Mini control exceptions', [['AEN mini control items'], ['PWI mini control items', 'deferred'], ...(has('1c') ? [['IOW mini control items']] : [])]],
        [2, 'Winter patrolling'], [3, S.pir ? 'Failures and PIRs' : 'Failures'], [4, 'Cautions and TDC overdue'], [5, 'Bad road surface at LC gates'],
        [6, 'NI LC inspection'], [7, 'CCTV recording check'], [8, 'Spot attention due'],
        [9, 'Work planned for the next day', [['Track-circuited area'], ['Non track-circuited area']]], [10, 'Inspections due'], [11, 'KPI items'],
        [12, 'Blocks availed and output'], [13, 'Machine utilisation'], [14, 'TMS dashboard'], [15, 'Worksite entry in TMS'],
        [16, 'Worksite safety arrangement', [['PWI unit-wise'], ...(S.iowSafety ? [['IOW unit-wise']] : [])]], [17, 'Weekly position']];
      const saturday = day.getUTCDay() === 6;
      const shown = ITEMS.filter(([n]) => has(n) || (n === 17 && saturday));
      const colW = Math.floor((CW - px(40)) / 2), numT = px(30);
      // each contents line: [gap above (px), line height (px), paragraph maker]
      const line = (n, title, page, grey) => o => new Paragraph({ tabStops: [{ type: TabStopType.LEFT, position: numT }, { type: TabStopType.RIGHT, position: colW, leader: LeaderType.DOT }],
        spacing: o, children: [t(String(n), { px: 14, font: SERIF, bold: true, color: grey ? '6B6B6B' : INK }), t('\t' + title, { px: 14, color: grey ? '6B6B6B' : INK }),
          t('\t', { px: 14, color: grey ? 'B0B0B0' : '888888' }), page === null ? t('Saturdays only', { px: 12.5, italic: true, color: '6B6B6B' }) : t(page, { px: 14 })] });
      const subLine = (title, note) => o => new Paragraph({ indent: { left: numT }, tabStops: [{ type: TabStopType.RIGHT, position: colW }], spacing: o,
        children: [t(title, { px: 12.5, color: '2B2B2B' }), ...(note ? [t('\t' + note, { px: 12.5, italic: true, color: '6B6B6B' })] : [])] });
      const lines = it => { const [n, title, subs] = it;
        return [[8, 17, line(n, title, n === 17 ? null : pg(n), n === 17)], ...(subs || []).map(([st, note], i) => [i ? 3 : 2, 15, subLine(st, note)])]; };
      const colL = shown.filter(i => i[0] <= 8).flatMap(lines), colR = shown.filter(i => i[0] > 8).flatMap(lines);
      [colL, colR].forEach(c => { if (c.length) c[0][0] = 14; });
      // stretch (or shrink) every gap by the same factor so the contents end at the footer, as on the design
      const SEAL = 256, sum = a => a.reduce((x, y) => x + y, 0);
      const fixedPx = SEAL + sum(stack.map(x => x[1])), gapPx = sum(stack.map(x => x[0]));
      const colH = c => [sum(c.map(x => x[1])), sum(c.map(x => x[0]))], [lh, lg] = colH(colL), [rh, rg] = colH(colR);
      const roomPx = AVAIL / 15 - 4;
      let k = 1; for (const tryK of [2.4, 2.2, 2, 1.8, 1.6, 1.5, 1.4, 1.3, 1.2, 1.1, 1, 0.9, 0.8, 0.7, 0.6]) {
        if (fixedPx + gapPx * tryK + Math.max(lh + lg * tryK, rh + rg * tryK) <= roomPx) { k = tryK; break; } k = tryK; }
      const exact = (gap, h) => ({ before: px(gap * k), after: 0, line: px(h), lineRule: LineRuleType.EXACT });
      const frame = art.frame ? new ImageRun({ type: 'png', data: art.frame, transformation: { width: 794, height: 1123 },
        floating: { horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 0 }, verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 0 },
          behindDocument: true, allowOverlap: true, lockAnchor: true, wrap: { type: TextWrappingType.NONE } } }) : null;
      const kidsC = [new Paragraph({ alignment: C, spacing: { before: 0, after: 0 }, children: [...(frame ? [frame] : []), img(art.seal, SEAL, SEAL)] }),
        ...stack.map(([g, h, runs]) => new Paragraph({ alignment: C, spacing: exact(g, h), children: runs }))];
      const col = c => new TableCell({ width: { size: colW, type: WidthType.DXA }, borders: NONE, margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: c.length ? c.map(([g, h, mk]) => mk(exact(g, h))) : [new Paragraph('')] });
      kidsC.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [colW, CW - 2 * colW, colW], layout: TableLayoutType.FIXED,
        rows: [new TableRow({ cantSplit: true, children: [col(colL), new TableCell({ width: { size: CW - 2 * colW, type: WidthType.DXA }, borders: NONE, children: [new Paragraph('')] }), col(colR)] })] }));
      const at = (R.generatedAt || '').slice(11, 16);
      const footerC = new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: CW }], children: [
        t('Prepared by Engg/CNL, Howrah' + (at ? ' · ' + at : ''), { px: 11, color: '555555' }),
        t('\tDesigned & developed by Anindya Nandy · SSE/P.Way/Planning, Howrah Division', { px: 11, color: '555555' })] })] });
      const usedPx = fixedPx + gapPx * k + Math.max(lh + lg * k, rh + rg * k);
      return { children: kidsC, footer: footerC, fill: usedPx / roomPx, k, index: shown.map(([n]) => [n, n === 17 ? null : +pg(n)]),
        page: { size: { width: 11906, height: 16838 }, margin: { top: TOP, bottom: BOT, left: px(80.2), right: px(80.2), footer: FOOT } } };
    };
    const top = topSheet(best.secs);
    const PAGE = { size: { width: 11906, height: 16838 }, margin: { top: 567, bottom: 800, left: 567, right: 567, footer: 240 } };
    if (opts.onLayout) opts.onLayout({ cover: { fill: top.fill, k: top.k, index: top.index }, density: pick.dz.name, arrangement: best.arr.name, pages: best.pages, end: best.end, err: best.err, ok: best.ok, okLoose: best.okLoose,
      score: best.score, breaks: best.brk, left: best.left, fills: best.fills, starts: best.starts, secs: best.secs, rows: rowsUsed, layout });

    return new Document({
      creator: 'AEN Morning Position', title: 'Morning Position ' + R.aen + ' ' + D.split('-').reverse().join('.'),
      styles: { default: { document: { run: { font: FONT, size: SZ } } } },
      sections: [{ properties: { page: top.page }, footers: { default: top.footer }, children: top.children },
        { properties: { page: PAGE }, footers: { default: footer }, children: kids }]
    });
  }
  g.buildMorningDoc = buildMorningDoc;
  if (typeof module !== 'undefined') module.exports = { buildMorningDoc };
})(typeof globalThis !== 'undefined' ? globalThis : this);
