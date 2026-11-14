/* The one place both the live site and the §25 preview build the .docx, so they never disagree (spec §24.1 / §25.8).
 * It fetches the masthead logo and the cover artwork relative to the page (ir-logo-256.png, cover-*.png), builds the
 * report with buildMorningDoc, and makes it Word-2007 safe. Needs the globals docx, JSZip, word2007Safe,
 * buildMorningDoc. Exposes window.AMPDocx. A picture that cannot be loaded is left out; the text still comes. */
(function (g) {
  var logoBytes = null, coverArt = null;   // fetched once
  async function getLogo() {
    if (logoBytes === null) { try { logoBytes = new Uint8Array(await (await fetch('ir-logo-256.png')).arrayBuffer()); } catch (e) { logoBytes = false; } }
    return logoBytes || null;
  }
  async function getCover() {
    if (!coverArt) { coverArt = {}; for (var i = 0, ks = ['frame', 'seal', 'rule', 'contents']; i < ks.length; i++) {
      try { coverArt[ks[i]] = new Uint8Array(await (await fetch('cover-' + ks[i] + '.png')).arrayBuffer()); } catch (e) { coverArt[ks[i]] = null; } } }
    return coverArt;
  }
  async function docxBlob(R, opts) {
    var o = { logo: await getLogo(), cover: await getCover() };
    if (opts) for (var k in opts) o[k] = opts[k];
    return word2007Safe(JSZip, await docx.Packer.toBlob(buildMorningDoc(docx, R, o)), 'blob');
  }
  // on-screen reader blobs (docx-preview does not flow pages itself):
  //  measureBlob: one page, every body block flat and in order, so the reader can measure the real rendered height.
  //  displayBlob(R, spec): split into one section per page at the reader's MEASURED break points (spec.breaks),
  //  so the reader shows discrete A4 pages. No spec → the builder's own Word-model split (used by tests). The
  //  Download always uses docxBlob (no display breaks).
  async function measureBlob(R) {
    var meta = null;
    var blob = await docxBlob(R, { displayPageBreaks: true, displaySpec: { measure: true }, onDisplayMeta: function (m) { meta = m; } });
    return { blob: blob, meta: meta };
  }
  async function displayBlob(R, spec) { return docxBlob(R, { displayPageBreaks: true, displaySpec: spec }); }
  function dmy_(iso) { return iso ? iso.split('-').reverse().join('.') : ''; }
  function docxName(R) {
    if (R && R.kind === 'srden') return 'Morning_Position_SrDEN-' + String(R.code || '').replace(/[^0-9A-Za-z]+/g, '') + '_' + dmy_(R.date) + '.docx';
    return 'Morning_Position_AEN-' + String(R.aen).replace('/', '-') + '_' + dmy_(R.date) + '.docx';
  }
  function saveBlob(blob, filename) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(a.href); }, 30000);
  }
  // On-screen reader only (Word is untouched). docx-preview places a picture anchored to the PAGE (the cover frame)
  // relative to the text area, so it lands one margin right and down, runs off the sheet and sits under the seal
  // (classic root, 13.11). Seat any sheet-sized portrait picture on the sheet itself, behind the text; and re-apply
  // the cover's tracked lettering, which docx-preview drops. Same rules as preview/reader.js seatCoverFrame.
  function seatCover(root) {
    var secs = root.querySelectorAll('section.docx');
    for (var i = 0; i < secs.length; i++) {
      var sec = secs[i], sr = sec.getBoundingClientRect(), ims = sec.querySelectorAll('img');
      for (var j = 0; j < ims.length; j++) {
        var im = ims[j], r = im.getBoundingClientRect();
        if (!(im.naturalHeight > im.naturalWidth && r.width > sr.width * 0.8 && r.height > sr.height * 0.8)) continue;
        if (getComputedStyle(sec).position === 'static') sec.style.position = 'relative';
        sec.style.zIndex = '0';   // its own stacking layer: the frame (-1) sits above the white sheet, below the text
        for (var w = im.parentElement; w && w !== sec; w = w.parentElement) if (getComputedStyle(w).position !== 'static') w.style.position = 'static';
        im.style.position = 'absolute'; im.style.left = '0'; im.style.top = '0';
        im.style.width = '100%'; im.style.height = '100%'; im.style.margin = '0'; im.style.zIndex = '-1';
      }
    }
    var first = root.querySelector('section.docx'); if (!first) return;
    var ps = first.querySelectorAll('p');
    for (var n = 0; n < ps.length; n++) {
      var t = (ps[n].textContent || '').trim();
      if (/^EASTERN RAILWAY\b.*HOWRAH DIVISION$/.test(t)) ps[n].style.letterSpacing = '0.32em';
      else if (/^AEN\/\S+\s+·\s+\S/.test(t) && t === t.toUpperCase()) ps[n].style.letterSpacing = '0.08em';
    }
  }
  g.AMPDocx = { getLogo: getLogo, getCover: getCover, docxBlob: docxBlob, measureBlob: measureBlob, displayBlob: displayBlob, docxName: docxName, saveBlob: saveBlob, seatCover: seatCover };
})(window);
