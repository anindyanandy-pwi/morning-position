/* §25.6 interaction and motion — physical, fast, transform/opacity only, no library. prefers-reduced-motion makes
 * everything instant. Exposes window.MX. */
(function () {
  const reduce = () => window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;
  const SPRING = 'linear(0,0.55 7.5%,0.9 15%,1.06 23%,1.02 35%,1,1)';   // slight overshoot, ease-out fallback via @supports

  /* press-in by size: small 0.95, normal 0.97, wide/card 0.985; spring back */
  function press(el, scale) {
    if (!el || el.__press) return; el.__press = true;
    const s = scale || (el.classList.contains('wide') || el.classList.contains('card') ? 0.985 : el.classList.contains('sm') ? 0.95 : 0.97);
    const down = () => { if (reduce()) return; el.style.transition = 'transform 90ms ease-out'; el.style.transform = 'scale(' + s + ')'; vibrate(8); };
    const up = () => { if (reduce()) { el.style.transform = ''; return; } el.style.transition = 'transform 260ms ' + springFn(); el.style.transform = 'scale(1)'; };
    el.addEventListener('pointerdown', down); el.addEventListener('pointerup', up);
    el.addEventListener('pointerleave', up); el.addEventListener('pointercancel', up);
  }
  function springFn() { return CSS && CSS.supports && CSS.supports('transition-timing-function', SPRING) ? SPRING : 'cubic-bezier(.22,1.2,.36,1)'; }
  function pressAll(root) { (root || document).querySelectorAll('[data-press]').forEach(el => press(el)); }

  /* Android haptics at notches/presses/done; iOS has no vibrate (snaps only) */
  function vibrate(ms) { try { if (navigator.vibrate && !reduce()) navigator.vibrate(ms); } catch (e) {} }

  /* a button becomes its own progress; never changes size */
  function busy(btn, label) {
    if (!btn) return () => {};
    const w = btn.getBoundingClientRect().width; btn.style.minWidth = w + 'px';
    const old = btn.textContent, wasDisabled = btn.disabled; btn.disabled = true; btn.classList.add('is-busy');
    if (label) btn.dataset.label = label, btn.textContent = label;
    return (doneText, kind) => {
      btn.classList.remove('is-busy');
      if (doneText) { btn.textContent = doneText; btn.classList.add(kind === 'err' ? 'did-err' : 'did-ok');
        setTimeout(() => { btn.textContent = old; btn.classList.remove('did-ok', 'did-err'); btn.disabled = wasDisabled; btn.style.minWidth = ''; }, 2200); }
      else { btn.textContent = old; btn.disabled = wasDisabled; btn.style.minWidth = ''; }
      vibrate(10);
    };
  }

  /* bottom toast with a thin countdown line; optional Undo (Keep/Undo) */
  let toastWrap = null;
  function toast(text, opts) {
    opts = opts || {};
    if (!toastWrap) { toastWrap = el('div', 'mx-toasts'); document.body.appendChild(toastWrap); }
    const t = el('div', 'mx-toast'), msg = el('span', 'mx-toast-msg', text); t.appendChild(msg);
    let undone = false;
    if (opts.undo) { const u = el('button', 'mx-toast-undo', 'Undo'); u.onclick = () => { undone = true; close(); opts.undo(); }; t.appendChild(u); }
    const bar = el('div', 'mx-toast-bar'); t.appendChild(bar);
    toastWrap.appendChild(t); requestAnimationFrame(() => t.classList.add('in'));
    const life = opts.ms || (opts.undo ? 6000 : 3200);
    if (!reduce()) { bar.style.transition = 'transform ' + life + 'ms linear'; requestAnimationFrame(() => bar.style.transform = 'scaleX(0)'); }
    const timer = setTimeout(close, life);
    function close() { clearTimeout(timer); t.classList.remove('in'); t.classList.add('out'); setTimeout(() => t.remove(), 220); if (!undone && opts.onclose) opts.onclose(); }
    return close;
  }

  /* hold-to-confirm: a ring fills over 0.6 s; resolves true if held, false if released early */
  function hold(btn, onConfirm) {
    if (!btn) return; const ring = el('span', 'mx-hold-ring'); btn.appendChild(ring); btn.classList.add('mx-hold');
    let t0 = 0, raf = 0, done = false;
    const start = e => { e.preventDefault(); done = false; t0 = performance.now(); if (reduce()) { finish(); return; } tick(); vibrate(8); };
    const tick = () => { const p = Math.min(1, (performance.now() - t0) / 600); ring.style.setProperty('--p', p); if (p >= 1) finish(); else raf = requestAnimationFrame(tick); };
    const finish = () => { if (done) return; done = true; cancelAnimationFrame(raf); ring.style.setProperty('--p', 0); vibrate(18); onConfirm(); };
    const cancel = () => { cancelAnimationFrame(raf); if (!done) ring.style.setProperty('--p', 0); };
    btn.addEventListener('pointerdown', start); btn.addEventListener('pointerup', cancel);
    btn.addEventListener('pointerleave', cancel); btn.addEventListener('pointercancel', cancel);
  }

  /* bottom sheet: drag/flick down to close */
  function sheet(node, opts) {
    opts = opts || {};
    const back = el('div', 'mx-sheet-back'), s = el('div', 'mx-sheet');
    const grip = el('div', 'mx-grip'); s.append(grip, node); back.appendChild(s); document.body.appendChild(back);
    document.body.style.overflow = 'hidden';
    requestAnimationFrame(() => { back.classList.add('in'); s.classList.add('in'); });
    let y0 = null, dy = 0;
    const close = () => { back.classList.remove('in'); s.classList.remove('in'); document.body.style.overflow = ''; setTimeout(() => back.remove(), 260); if (opts.onclose) opts.onclose(); };
    back.onclick = e => { if (e.target === back) close(); };
    grip.addEventListener('pointerdown', e => { y0 = e.clientY; s.style.transition = 'none'; grip.setPointerCapture(e.pointerId); });
    grip.addEventListener('pointermove', e => { if (y0 == null) return; dy = Math.max(0, e.clientY - y0); s.style.transform = 'translateY(' + dy + 'px)'; });
    grip.addEventListener('pointerup', () => { if (y0 == null) return; s.style.transition = ''; if (dy > 90) close(); else s.style.transform = ''; y0 = null; dy = 0; });
    return close;
  }

  /* skeleton placeholder with a faint shimmer */
  function skeleton(kind) { const s = el('div', 'mx-skel ' + (kind || '')); return s; }

  function el(tag, cls, text) { const n = document.createElement(tag); if (cls) n.className = cls; if (text != null) n.textContent = text; return n; }

  window.MX = { press, pressAll, vibrate, busy, toast, hold, sheet, skeleton, reduce, springFn };
})();
