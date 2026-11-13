/* Word 2007 compatibility (user, 05.11.2026): Word 2007 refuses a .docx ("problems with the contents") when a
 * picture's <wp:docPr> carries the title="" attribute, which exists only from Word 2010 on; the docx library writes
 * it on every picture. This rewrites the finished file: the title attribute is dropped and an empty picture name
 * becomes "Picture <id>". Nothing else changes; newer Word reads the result the same. Works in the browser
 * (window.word2007Safe) and in node (module.exports). */
(function (root) {
  function fixXml(s, next) {
    // ids must also be unique across the file (the docx library can repeat them); next() hands out the next one
    return s.replace(/<wp:docPr\b[^>]*>/g, function (tag) {
      var id = String(next());
      return tag.replace(/\s+title="[^"]*"/, '').replace(/\bid="\d+"/, 'id="' + id + '"').replace(/\bname=""/, 'name="Picture ' + id + '"');
    });
  }
  /** data: Blob / ArrayBuffer / Buffer of a .docx; type: 'blob' (browser) or 'nodebuffer'. */
  async function word2007Safe(JSZip, data, type) {
    var zip = await JSZip.loadAsync(data);
    var n = 1000, next = function () { return n++; };
    var names = Object.keys(zip.files).filter(function (n) { return /^word\/[^/]+\.xml$/.test(n); });
    for (var i = 0; i < names.length; i++) {
      var s = await zip.file(names[i]).async('string'), t = fixXml(s, next);
      if (t !== s) zip.file(names[i], t);
    }
    return zip.generateAsync({ type: type || 'blob', compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }
  root.word2007Safe = word2007Safe; 
  if (typeof module !== 'undefined' && module.exports) module.exports = { word2007Safe: word2007Safe, fixXml: fixXml };
})(typeof window !== 'undefined' ? window : globalThis);
