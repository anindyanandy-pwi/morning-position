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
  function docxName(R) { return 'Morning_Position_AEN-' + String(R.aen).replace('/', '-') + '_' + dmy_(R.date) + '.docx'; }
  function saveBlob(blob, filename) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(a.href); }, 30000);
  }
  g.AMPDocx = { getLogo: getLogo, getCover: getCover, docxBlob: docxBlob, measureBlob: measureBlob, displayBlob: displayBlob, docxName: docxName, saveBlob: saveBlob };
})(window);
