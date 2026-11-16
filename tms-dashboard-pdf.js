/**
 * §32 TMS dashboard PDF. Draws the rows from buildTmsDashboard(D) in the user's sample's format — A4 landscape, two
 * pages, the title centred across the full width, the Sr DEN group heads and AEN names as merged cells over the ten
 * AENs, SN | Item | sub | Total | ten AENs with a D / D-1 pair under each, a green SN column, yellow-fill + red-bold
 * highlighting by rule, a heavier line after each pair, and the header repeated on page 2. Low-level pdf-lib drawing
 * (pinned web/vendor/pdf-lib-1.17.1.min.js), the same library in node for the tests.
 * buildTmsDashboardPdf(PDFLib, R) -> Promise<Uint8Array>.
 */
(function (root) {
  function pdfName(R) {
    var d = String(R.date).split('-');                    // YYYY-MM-DD -> DD.MM.YYYY
    return 'TMS_DASHBOARD_-_HOWRAH_DIVISION_-' + d[2] + '.' + d[1] + '.' + d[0] + '.pdf';
  }

  async function buildTmsDashboardPdf(PDFLib, R) {
    var PDFDocument = PDFLib.PDFDocument, StandardFonts = PDFLib.StandardFonts, rgb = PDFLib.rgb;
    var doc = await PDFDocument.create();
    var font = await doc.embedFont(StandardFonts.Helvetica);
    var bold = await doc.embedFont(StandardFonts.HelveticaBold);

    var W = 841.89, H = 595.28, M = 16;                   // A4 landscape
    var colSN = 20, colItem = 196, colSub = 54, subW = 24, grpW = subW * 2;
    var xSN = M, xItem = xSN + colSN, xSub = xItem + colItem, xTotal = xSub + colSub;
    var groupX = [], gx = xTotal; for (var i = 0; i < 11; i++) { groupX.push(gx); gx += grpW; }
    var rightEdge = gx;
    var GREY = rgb(0.6, 0.6, 0.6), BLACK = rgb(0, 0, 0), GREEN = rgb(0, 0.69, 0.31), YELLOW = rgb(1, 1, 0), RED = rgb(0.77, 0, 0);
    var dShort = R.date.slice(8) + '.' + R.date.slice(5, 7), pShort = R.prevDate.slice(8) + '.' + R.prevDate.slice(5, 7);
    var dmy = R.date.split('-').reverse().join('.'), pdmy = R.prevDate.split('-').reverse().join('.');
    var aenLabel = { 'TR/HWH': 'TRHWH', LLH: 'LLH', BDC: 'BDC', DKAE: 'DKAE', BWN: 'BWN', NDAE: 'NDAE', KWAE: 'KWAE', AZ: 'AZ', BHP: 'BHP', RPH: 'RPH' };

    function wrap(t, width, size) {
      var out = [];
      String(t == null ? '' : t).split('\n').forEach(function (part) {
        var words = part.split(/\s+/).filter(Boolean), cur = '';
        for (var k = 0; k < words.length; k++) { var s = cur ? cur + ' ' + words[k] : words[k];
          if (font.widthOfTextAtSize(s, size) > width && cur) { out.push(cur); cur = words[k]; } else cur = s; }
        out.push(cur);
      });
      return out.length ? out : [''];
    }
    function fitSize(t, width, size) { var s = size; while (s > 4 && font.widthOfTextAtSize(String(t), s) > width) s -= 0.5; return s; }

    var page;
    function fill(x, yBot, w, h, color) { page.drawRectangle({ x: x, y: yBot, width: w, height: h, color: color }); }
    function vseg(x, yTop, yBot, thick) { page.drawLine({ start: { x: x, y: yTop }, end: { x: x, y: yBot }, thickness: thick || 0.4, color: GREY }); }
    function hseg(x1, x2, yy) { page.drawLine({ start: { x: x1, y: yy }, end: { x: x2, y: yy }, thickness: 0.4, color: GREY }); }
    function txt(s, x, w, yy, o) { o = o || {}; s = s == null ? '' : String(s); if (s === '') return;
      var fo = o.font || font, sz = o.size || 6, tw = fo.widthOfTextAtSize(s, sz), tx;
      if (o.align === 'l') tx = x + 2; else if (o.align === 'r') tx = x + w - tw - 2; else tx = x + (w - tw) / 2;
      page.drawText(s, { x: tx, y: yy, size: sz, font: fo, color: o.color || BLACK }); }

    // the group heads and their AEN spans
    var groups = R.groups;   // [[label,[aens]], ...]

    // the heading block (title + group heads + SN./Item/Total + AEN names + dates); drawn on both pages. Returns the y
    // below it. The group-head and AEN-name cells are merged (no internal vertical line); the SN corner cell is yellow.
    var dateTop;                                           // the y of the date-row top (set by header; verticals below it)
    function header(top) {
      var titleH = 18, hA = 14, hB = 13, hC = 11;
      var yTB = top - titleH, yAB = yTB - hA, yBB = yAB - hB, yEnd = yBB - hC;      // band bottoms, top → down
      dateTop = yBB;
      // title across the full width, centred and bold
      txt(R.title, M, rightEdge - M, yTB + titleH / 2 - 4, { size: 11, font: bold, align: 'c' });
      page.drawRectangle({ x: M, y: yTB, width: rightEdge - M, height: titleH, borderColor: GREY, borderWidth: 0.4 });
      // SN column: green through the header below the title, yellow over the group-head band
      fill(xSN, yEnd, colSN, yTB - yEnd, GREEN);
      fill(xSN, yAB, colSN, yTB - yAB, YELLOW);
      txt('SN.', xSN, colSN, (yAB + yBB) / 2 - 2, { size: 6.5, font: bold });
      txt('Item', xItem, colItem, (yAB + yEnd) / 2 - 2, { size: 7, font: bold });
      // group heads over their AEN columns (merged — the rect is the only border, so no internal line)
      var ai = 0;
      groups.forEach(function (g) { var span = g[1].length, gX = groupX[1 + ai], w = span * grpW;
        page.drawRectangle({ x: gX, y: yAB, width: w, height: yTB - yAB, borderColor: GREY, borderWidth: 0.4 });
        txt(g[0], gX, w, (yAB + yTB) / 2 - 3, { size: 6.5, font: bold }); ai += span; });
      // AEN-name band: Total and each AEN name merged over their D/D-1 pair
      [{ x: xTotal, l: 'Total' }].concat(R.aens.map(function (a, j) { return { x: groupX[1 + j], l: aenLabel[a] || a }; })).forEach(function (c) {
        page.drawRectangle({ x: c.x, y: yBB, width: grpW, height: yAB - yBB, borderColor: GREY, borderWidth: 0.4 });
        txt(c.l, c.x, grpW, (yBB + yAB) / 2 - 3, { size: 6.5, font: bold });
      });
      // date band: 13.11 | 12.11 under Total and each AEN, bold
      for (var c = 0; c < 11; c++) {
        txt(dShort, groupX[c], subW, (yEnd + yBB) / 2 - 3, { size: 5.5, font: bold });
        txt(pShort, groupX[c] + subW, subW, (yEnd + yBB) / 2 - 3, { size: 5.5, font: bold });
      }
      // horizontal borders + the left-column verticals (through the whole header below the title)
      hseg(M, rightEdge, yTB); hseg(M, rightEdge, yEnd);
      [xSN, xItem, xSub, xTotal].forEach(function (x) { vseg(x, yTB, yEnd); });
      vseg(M, top, yEnd, 1.1); vseg(rightEdge, top, yEnd, 1.1);
      return yEnd;
    }

    var rows = R.rows, splitAt = rows.findIndex(function (r) { return r.sn === 35; });
    if (splitAt < 0) splitAt = rows.length;
    var pageBlocks = [rows.slice(0, splitAt), rows.slice(splitAt)];

    pageBlocks.forEach(function (pageRows) {
      page = doc.addPage([W, H]);
      var yStart = header(H - M);
      // geometry
      var geo = [], yy = yStart;
      pageRows.forEach(function (r) {
        var itemLines = wrap(r.item, colItem - 4, 6), rowH = Math.max(itemLines.length * 7, 10) + 2;
        geo.push({ r: r, top: yy, bot: yy - rowH, itemLines: itemLines }); yy -= rowH;
      });
      var yEnd = yy;
      var spanEnd = function (i, key) { var j = i; while (j + 1 < geo.length && geo[j + 1].r[key] === '') j++; return j; };

      // 1. fills — green SN column, then the yellow highlight cells
      geo.forEach(function (g) {
        var r = g.r, h = g.top - g.bot;
        fill(xSN, g.bot, colSN, h, GREEN);
        var pf = function (gX, hd, hp) { if (hd) fill(gX, g.bot, subW, h, YELLOW); if (hp) fill(gX + subW, g.bot, subW, h, YELLOW); };
        pf(xTotal, r.hlCurTotal, r.hlPrevTotal);
        r.hlCur.forEach(function (_, j) { pf(groupX[1 + j], r.hlCur[j], r.hlPrev[j]); });
      });

      // 2. grid lines (segmented so a merged SN/Item cell has no internal border; heavier after each pair)
      hseg(M, rightEdge, yStart);
      geo.forEach(function (g, i) {
        var next = geo[i + 1];
        if (!(next && next.r.sn === '')) hseg(xSN, xItem, g.bot);
        if (!(next && next.r.item === '')) hseg(xItem, xSub, g.bot);
        hseg(xSub, rightEdge, g.bot);
      });
      [xSN, xItem, xSub, xTotal].forEach(function (x) { vseg(x, yStart, yEnd); });
      for (var k = 0; k < 11; k++) vseg(groupX[k] + subW, yStart, yEnd);          // D/D-1 within each group (thin)
      for (var m = 0; m < 11; m++) vseg(groupX[m] + grpW, yStart, yEnd, 1.1);      // after each pair (heavier)
      vseg(rightEdge, yStart, yEnd); vseg(M, yStart, yEnd);

      // 3. text
      geo.forEach(function (g, i) {
        var r = g.r, midY = (g.top + g.bot) / 2 - 2, vsz = r.effSn === 7 ? 8 : 6;   // SN 7 (LC OH, 24 months) larger
        if (r.sn !== '') { var se = spanEnd(i, 'sn'), cy = (g.top + geo[se].bot) / 2 - 2; txt(String(r.sn), xSN, colSN, cy, { size: 6, font: bold }); }
        if (r.item !== '') { var ie = spanEnd(i, 'item'), c = (g.top + geo[ie].bot) / 2, startY = c + (g.itemLines.length * 7) / 2 - 6;
          g.itemLines.forEach(function (ln, li) { txt(ln, xItem, colItem, startY - li * 7, { size: 6, align: 'c' }); }); }
        txt(r.sub, xSub, colSub, midY, { size: 5.5, align: 'c' });
        var pair = function (gX, d, p, hd, hp, bold0) {
          txt(d, gX, subW, midY, { size: fitSize(d, subW - 1, vsz), font: (hd || bold0) ? bold : font, color: hd ? RED : BLACK });
          txt(p, gX + subW, subW, midY, { size: fitSize(p, subW - 1, vsz), font: (hp || bold0) ? bold : font, color: hp ? RED : BLACK });
        };
        pair(xTotal, r.curTotal, r.prevTotal, r.hlCurTotal, r.hlPrevTotal, true);   // Total figures bold
        r.cur.forEach(function (cc, j) { pair(groupX[1 + j], cc, r.prev[j], r.hlCur[j], r.hlPrev[j], false); });
      });
    });

    // footer: generated time + any AEN/day with no TMS read (that cell shows "–")
    var notes = [];
    if (R.noCur) notes.push('No TMS items read for ' + dmy + '.'); else if (R.missingCur.length) notes.push('No TMS read on ' + dmy + ' for: ' + R.missingCur.join(', ') + '.');
    if (R.noPrev) notes.push('No TMS items read for ' + pdmy + '.'); else if (R.missingPrev.length) notes.push('No TMS read on ' + pdmy + ' for: ' + R.missingPrev.join(', ') + '.');
    if (R.noSpot) notes.push('Spot attention not read for ' + dmy + '.');
    var foot = 'TMS data only · generated ' + (R.generatedAt || '') + (notes.length ? '  ·  ' + notes.join('  ') : '');
    var last = doc.getPages()[doc.getPageCount() - 1];
    last.drawText(foot, { x: M, y: M - 6, size: 5.5, font: font, color: GREY });

    return doc.save();
  }

  root.TmsDashboardPdf = { buildTmsDashboardPdf: buildTmsDashboardPdf, pdfName: pdfName };
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));

if (typeof module !== 'undefined' && module.exports) module.exports = (typeof window !== 'undefined' ? window : globalThis).TmsDashboardPdf;
