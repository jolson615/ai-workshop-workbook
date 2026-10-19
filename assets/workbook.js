/* RESET, run before anything else so every write is seen.
   - Storage.setItem/removeItem are wrapped. A wb:* key written while a session page shows slide N is
     recorded in the registry wb:reg:s<N> (JSON: slide index -> keys). Nav and registry keys are not recorded.
   - "Reset this activity" (footer) removes that slide's registered keys and reloads at the same hash.
   - Shift+R resets the current slide, same as the button. Ignored while focus is in an input,
     textarea or contenteditable.
   - ?fresh: loading a session page with ?fresh (e.g. session-2.html?fresh#2) clears ALL of that
     session's wb:s<N>: keys and its registry, then strips ?fresh from the URL and keeps the hash.
     This is the presenter's "start the whole session clean" link. */
(function () {
  'use strict';
  var LS = null;
  try { LS = window.localStorage; } catch (e) {}
  var proto = window.Storage ? Storage.prototype : null;
  var rawSet = proto && proto.setItem, rawRemove = proto && proto.removeItem;

  function sessionOf() { return document.body ? document.body.getAttribute('data-session') : null; }
  function readReg(s) {
    try {
      var v = LS.getItem('wb:reg:s' + s), o = v ? JSON.parse(v) : null;
      return o && typeof o === 'object' ? o : {};
    } catch (e) { return {}; }
  }
  function writeReg(s, reg) { try { rawSet.call(LS, 'wb:reg:s' + s, JSON.stringify(reg)); } catch (e) {} }
  function isTracked(k) {
    return typeof k === 'string' && k.indexOf('wb:') === 0 && k.indexOf('wb:reg:') !== 0 && k !== 'wb:nav';
  }
  function slideIndex() {
    var all = document.querySelectorAll('.stage > .slide');
    for (var i = 0; i < all.length; i++) if (all[i].classList.contains('on')) return i;
    return -1;
  }
  function track(k) {
    var s = sessionOf(), i;
    if (!s || !isTracked(k)) return;
    i = slideIndex();
    if (i < 0) return;
    var reg = readReg(s), list = reg[i] || [];
    if (list.indexOf(k) > -1) return;
    list.push(k);
    reg[i] = list;
    writeReg(s, reg);
    document.dispatchEvent(new Event('wb:registry'));
  }
  function untrack(k) {
    var s = sessionOf();
    if (!s || !isTracked(k)) return;
    var reg = readReg(s), changed = false;
    Object.keys(reg).forEach(function (i) {
      var list = reg[i], at = list.indexOf(k);
      if (at > -1) { list.splice(at, 1); changed = true; if (!list.length) delete reg[i]; }
    });
    if (changed) { writeReg(s, reg); document.dispatchEvent(new Event('wb:registry')); }
  }
  if (proto && rawSet && rawRemove) {
    proto.setItem = function (k, v) {
      var r = rawSet.apply(this, arguments);
      if (this === LS) track(k);
      return r;
    };
    proto.removeItem = function (k) {
      var r = rawRemove.apply(this, arguments);
      if (this === LS) untrack(k);
      return r;
    };
  }

  /* ?fresh: clear this session's state before any slide script reads it */
  (function freshStart() {
    var s = sessionOf(), sp;
    if (!s || !LS) return;
    try { sp = new URLSearchParams(location.search); } catch (e) { return; }
    if (!sp.has('fresh')) return;
    try {
      var prefix = 'wb:s' + s + ':', regKey = 'wb:reg:s' + s, drop = [], i;
      for (i = 0; i < LS.length; i++) {
        var k = LS.key(i);
        if (k && (k.indexOf(prefix) === 0 || k === regKey)) drop.push(k);
      }
      drop.forEach(function (k) { rawRemove.call(LS, k); });
    } catch (e) {}
    try { history.replaceState(null, '', location.pathname + location.hash); } catch (e) {}
  })();
})();

/* Workbook shared script. Classic script, no modules, no fetch(), so it works from file://.
   Every localStorage call is wrapped in try/catch. Session 1 keys are prefixed wb:s1:, nav uses wb:nav. */
(function () {
  'use strict';

  function $(s, r) { return (r || document).querySelector(s); }
  function $$(s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); }
  function store(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  function load(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function storeJSON(k, v) { store(k, JSON.stringify(v)); }
  function loadJSON(k, d) {
    var v = load(k);
    if (v === null) return d;
    try { return JSON.parse(v); } catch (e) { return d; }
  }
  function clearWb() {
    try {
      var keys = [], i;
      for (i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && k.indexOf('wb:') === 0) keys.push(k);
      }
      keys.forEach(function (k) { localStorage.removeItem(k); });
    } catch (e) {}
  }

  var body = document.body;
  var app = $('#app');
  var stage = $('#stage');

  /* ---------- side nav ---------- */
  function setNav(collapsed) {
    if (!app) return;
    app.classList.toggle('collapsed', collapsed);
    document.documentElement.classList.toggle('nav-c', collapsed);
    var t = $('#navToggle');
    if (t) {
      t.setAttribute('aria-expanded', String(!collapsed));
      t.setAttribute('aria-label', collapsed ? 'Expand navigation' : 'Collapse navigation');
      t.setAttribute('title', collapsed ? 'Expand sidebar ( [ )' : 'Collapse sidebar ( [ )');
    }
    store('wb:nav', collapsed ? '1' : '0');
  }
  if (app) {
    var navT = $('#navToggle');
    if (navT) navT.innerHTML = '<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false"><polyline points="10 3 5 8 10 13"/></svg>';
    var savedNav = load('wb:nav');
    setNav(savedNav !== null ? savedNav === '1' : window.innerWidth < 900);
    /* re-enable transitions only after the saved state has painted */
    setTimeout(function () { document.documentElement.classList.remove('preload'); }, 60);
    $('#navToggle').addEventListener('click', function () { setNav(!app.classList.contains('collapsed')); });
    buildNavReset();
  }

  /* ---------- sidebar: reset all activities (two-step, no window.confirm) ---------- */
  function buildNavReset() {
    var nav = $('.nav', app);
    if (!nav) return;
    var TITLE = 'Reset all activities in every session', ARMED = 'Click again to reset everything';
    var LABEL = 'Reset all activities';
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'nav-reset';
    btn.setAttribute('aria-label', LABEL);
    btn.setAttribute('title', TITLE);
    btn.innerHTML = '<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false"><path d="M13.5 8a5.5 5.5 0 1 1-1.61-3.89"/><polyline points="13.5 2.5 13.5 5.5 10.5 5.5"/></svg><span class="label"></span>';
    btn.querySelector('.label').textContent = LABEL;
    var live = document.createElement('span');
    live.className = 'sr-only';
    live.setAttribute('aria-live', 'polite');
    nav.appendChild(btn);
    nav.appendChild(live);

    var armed = false, timer = null;
    function disarm() {
      armed = false;
      clearTimeout(timer);
      btn.classList.remove('armed');
      btn.setAttribute('aria-label', LABEL);
      btn.setAttribute('title', TITLE);
      btn.querySelector('.label').textContent = LABEL;
      live.textContent = '';
    }
    function resetAllActivities() {
      var keys = [], i;
      try {
        for (i = 0; i < localStorage.length; i++) {
          var k = localStorage.key(i);
          if (k && k.indexOf('wb:') === 0 && k !== 'wb:nav') keys.push(k);
        }
        keys.forEach(function (k) { localStorage.removeItem(k); });
      } catch (e) {}
      location.reload();
    }
    btn.addEventListener('click', function () {
      if (!armed) {
        armed = true;
        btn.classList.add('armed');
        btn.setAttribute('aria-label', ARMED);
        btn.setAttribute('title', ARMED);
        btn.querySelector('.label').textContent = ARMED;
        live.textContent = ARMED;
        timer = setTimeout(disarm, 4000);
        return;
      }
      resetAllActivities();
    });
  }

  /* ---------- slides, steps, pips, hash ---------- */
  var slides = body.hasAttribute('data-session') ? $$('.stage > .slide') : [];
  var pos = 0;
  var revealed = [];
  var resetBtn = null;
  var WIDGET_SEL = '.qs, input, textarea, [data-widget], .pick, .take';
  var ATTN = { up: 'Eyes up front', laptop: 'On your laptop' };

  function stepsOf(i) { return $$('.step', slides[i]); }
  function applySteps(i) {
    var n = revealed[i] || 0;
    var steps = stepsOf(i);
    steps.forEach(function (s, k) { s.classList.toggle('on', k < n); });
    markSteps(i, steps);
  }
  /* emphasis follows attention: the newest revealed step is em-current, earlier ones em-done.
     A slide with data-em="off" marks its own items in its script. */
  function markSteps(i, steps) {
    if (slides[i].getAttribute('data-em') === 'off') return;
    var shown = steps.filter(function (s) { return s.classList.contains('on'); });
    steps.forEach(function (s) { s.classList.remove('em-current', 'em-done'); });
    shown.forEach(function (s, k) { s.classList.add(k === shown.length - 1 ? 'em-current' : 'em-done'); });
  }

  function show(p, dir) {
    if (!slides.length) return;
    pos = Math.max(0, Math.min(slides.length - 1, p));
    slides.forEach(function (el, i) {
      var on = i === pos;
      el.classList.toggle('on', on);
      el.inert = !on;
    });
    var n = stepsOf(pos).length;
    if (n) {
      /* slides marked data-steps-start open with no steps shown (session 2 blanks, bars) */
      var start = slides[pos].hasAttribute('data-steps-start') ? 0 : 1;
      revealed[pos] = dir < 0 ? n : Math.min(start, n);
      applySteps(pos);
    }
    stage.scrollTop = 0;
    var mode = slides[pos].getAttribute('data-attn') === 'laptop' ? 'laptop' : 'up';
    var attn = $('#attn');
    if (attn) {
      attn.setAttribute('data-mode', mode);
      attn.textContent = ATTN[mode];
    }
    var pBtn = $('#prev'), nBtn = $('#next');
    if (pBtn) pBtn.disabled = pos === 0;
    if (nBtn) nBtn.disabled = pos === slides.length - 1;
    $$('.pip').forEach(function (b, i) {
      if (i === pos) b.setAttribute('aria-current', 'step');
      else b.removeAttribute('aria-current');
    });
    updateReset();
  }

  function go(p, dir) {
    show(p, dir);
    var h = '#' + (pos + 1);
    if (location.hash !== h) location.hash = h;
  }

  function next() {
    var n = slides.length ? stepsOf(pos).length : 0;
    if (n && (revealed[pos] || 0) < n) {
      revealed[pos] = (revealed[pos] || 0) + 1;
      applySteps(pos);
      return;
    }
    go(pos + 1, 1);
  }
  function prev() { go(pos - 1, -1); }

  function buildPips() {
    var box = $('#pips');
    if (!box) return;
    slides.forEach(function (el, i) {
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'pip';
      b.setAttribute('aria-label', 'Slide ' + (i + 1) + ' of ' + slides.length);
      b.addEventListener('click', function () { go(i, i < pos ? -1 : 1); });
      box.appendChild(b);
    });
  }

  /* ---------- reset this activity (see the RESET notes at the top of this file) ---------- */
  function regFor(i) {
    var s = body.getAttribute('data-session');
    var reg = s ? loadJSON('wb:reg:s' + s, {}) : {};
    if (!reg || typeof reg !== 'object') reg = {};
    return { session: s, reg: reg, keys: Array.isArray(reg[i]) ? reg[i] : [] };
  }
  function canReset(i) {
    return !!slides[i] && (regFor(i).keys.length > 0 || !!slides[i].querySelector(WIDGET_SEL));
  }
  function resetSlide(i) {
    var info = regFor(i);
    /* removeItem is wrapped, so each removed key also leaves the registry */
    info.keys.slice().forEach(function (k) { try { localStorage.removeItem(k); } catch (e) {} });
    var fresh = regFor(i);
    delete fresh.reg[i];
    if (fresh.session) storeJSON('wb:reg:s' + fresh.session, fresh.reg);
    location.reload();
  }
  function updateReset() {
    if (resetBtn) resetBtn.hidden = !canReset(pos);
  }
  /* the pips move into the footer, between Back and Next; the reset sits quietly to their left */
  function buildFooter() {
    var foot = $('.foot'), pips = $('#pips'), nb = $('#next');
    if (!foot || !pips || !nb) return;
    var mid = document.createElement('div');
    mid.className = 'mid';
    foot.insertBefore(mid, nb);
    mid.appendChild(pips);
    resetBtn = document.createElement('button');
    resetBtn.type = 'button';
    resetBtn.className = 'reset';
    resetBtn.id = 'reset';
    resetBtn.hidden = true;
    resetBtn.textContent = 'Reset this activity';
    resetBtn.addEventListener('click', function () { resetSlide(pos); });
    mid.appendChild(resetBtn);
  }

  if (slides.length) {
    buildPips();
    buildFooter();
    document.addEventListener('wb:registry', updateReset);
    var pb = $('#prev'), nb = $('#next');
    if (pb) pb.addEventListener('click', prev);
    if (nb) nb.addEventListener('click', next);
    /* No click-anywhere-to-advance: it made text hard to select. Steps and slides advance only via
       Next/Back, the pips, the nav, or the arrow keys. */
    window.addEventListener('hashchange', function () {
      var n = parseInt(location.hash.slice(1), 10);
      if (!n) return;
      var target = Math.max(0, Math.min(slides.length - 1, n - 1));
      if (target === pos) return;
      show(target, target < pos ? -1 : 1);
    });
    var start = parseInt(location.hash.slice(1), 10);
    show(start ? start - 1 : 0, 0);
    try { history.replaceState(null, '', '#' + (pos + 1)); } catch (e) {}
  }

  document.addEventListener('keydown', function (e) {
    var t = e.target;
    if (t && (t.isContentEditable || (t.closest && t.closest('textarea,input,[contenteditable]')))) return;
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    if (e.shiftKey && (e.key === 'R' || e.key === 'r')) {
      if (slides.length && canReset(pos)) resetSlide(pos);
      return;
    }
    if (e.key === '[') { if (app) setNav(!app.classList.contains('collapsed')); return; }
    if (!slides.length) return;
    if (e.key === 'ArrowRight') { e.preventDefault(); next(); }
    else if (e.key === 'ArrowLeft') { e.preventDefault(); prev(); }
  });

  /* ---------- session 1: build an agent ---------- */
  function initBuild() {
    var root = $('[data-widget="build"]');
    if (!root) return;
    var key = 'wb:s1:build';
    var picked = loadJSON(key, []);
    if (!Array.isArray(picked)) picked = [];
    var togs = $$('.tog', root);
    var parts = $$('.part', root);
    var count = $('.count', root);
    var ask = $('.ask', root);
    function render() {
      togs.forEach(function (b, i) {
        var on = picked.indexOf(i) > -1;
        b.setAttribute('aria-pressed', String(on));
        b.classList.toggle('em-current', on);
      });
      parts.forEach(function (g) {
        g.classList.toggle('on', picked.indexOf(Number(g.getAttribute('data-part'))) > -1);
      });
      var n = picked.length;
      count.textContent = n === 0 ? 'Pick the parts you’d need.' :
        n + (n === 1 ? ' part: ' : ' parts: ') + (n <= 2 ? 'easier to build.' : n <= 5 ? 'harder to build.' : 'much harder to build.');
      var fill = $('.bm-fill', root);
      if (fill) fill.style.width = Math.round(100 * n / togs.length) + '%';
      ask.hidden = picked.length === 0;
    }
    togs.forEach(function (b, i) {
      b.addEventListener('click', function () {
        var at = picked.indexOf(i);
        if (at > -1) picked.splice(at, 1); else picked.push(i);
        picked.sort(function (x, y) { return x - y; });
        storeJSON(key, picked);
        render();
      });
    });
    render();
  }

  /* ---------- session 1: permissions grid (No -> With approval -> Yes) ---------- */
  function initPerm() {
    var root = $('[data-widget="perm"]');
    if (!root) return;
    var key = 'wb:s1:perms';
    var LABEL = ['No', 'With approval', 'Yes'];
    var state = loadJSON(key, null);
    if (!Array.isArray(state) || state.length !== 10) state = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
    var cells = $$('.cell', root);
    var segs = $$('.seg', root);
    function render() {
      cells.forEach(function (c) {
        var i = Number(c.getAttribute('data-i'));
        var s = state[i];
        c.textContent = LABEL[s];
        c.setAttribute('data-s', String(s));
        c.setAttribute('aria-label', c.getAttribute('data-name') + ': ' + LABEL[s]);
      });
      segs.forEach(function (g, i) { g.setAttribute('data-s', String(state[i])); });
    }
    cells.forEach(function (c) {
      c.addEventListener('click', function () {
        var i = Number(c.getAttribute('data-i'));
        state[i] = (state[i] + 1) % 3;
        storeJSON(key, state);
        render();
      });
    });
    render();
  }

  /* ---------- merge a sorted board into one "Tools" column ---------- */
  function mergeBoard(board) {
    var info = $('.zone-info', board), act = $('.zone-act', board);
    $$('.card', act).forEach(function (c) { $('.drop', info).appendChild(c); });
    $('.zlabel', info).textContent = 'Tools';
    $('.zlabel', info).classList.add('em-current');
    $$('.card', info).forEach(function (c) { c.classList.add('em-done'); });
    board.classList.add('merged');
  }

  /* ---------- session 1: lunch, sorted (merge on click) ---------- */
  function initMerge() {
    var board = $('#s1-merge');
    if (!board) return;
    var key = 'wb:s1:merge';
    var btn = $('#s1-merge-btn');
    var line = $('#s1-merge-line');
    function apply() {
      mergeBoard(board);
      btn.hidden = true;
      line.classList.add('on', 'em-current');
    }
    if (load(key) === '1') apply();
    btn.addEventListener('click', function () { apply(); store(key, '1'); });
  }

  /* ---------- shared pattern (.qs): one question at a time ----------
     Markup: .qs[data-qs] > ol.qs-done + .qs-item[data-answer][data-short] (.qs-text, .pick buttons, .qs-reveal)
     + .qs-count. Answering collapses the question into a summary row and shows the next one at once.
     A match gets a green check (with a small pop); a miss gets a grey x. The newest summary also shows
     its reveal text as a muted line. data-qs-mode="record" shows only the choice: no check, no x.
     Optional data-on-done="<id>" shows that element once every item is answered.
     Storage key: wb:s<session>:q:<data-qs>. */
  function confettiHTML() {
    var s = '', k;
    for (k = 0; k < 5; k++) s += '<i class="cf" aria-hidden="true"></i>';
    return s;
  }
  var CHECK_SVG = '<svg viewBox="0 0 24 24" aria-hidden="true"><path class="tick" d="M5 12.6l4.3 4.3L19 7.2"/></svg>' + confettiHTML();
  var CROSS_SVG = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6.5 6.5l11 11M17.5 6.5l-11 11"/></svg>';

  function initQs(root) {
    var sess = body.getAttribute('data-session') || '0';
    var key = 'wb:s' + sess + ':q:' + root.getAttribute('data-qs');
    var record = root.getAttribute('data-qs-mode') === 'record';
    var items = $$('.qs-item', root);
    var done = $('.qs-done', root);
    var count = $('.qs-count', root);
    var onDone = root.getAttribute('data-on-done');
    var doneEl = onDone ? document.getElementById(onDone) : null;
    var st = { cur: 0, picks: items.map(function () { return null; }) };
    var saved = loadJSON(key, null);
    if (saved && Array.isArray(saved.picks) && saved.picks.length === items.length) {
      st.picks = saved.picks.map(function (p, i) {
        return $('.pick[data-pick="' + p + '"]', items[i]) ? p : null;
      });
      st.cur = Math.max(0, Math.min(items.length, Number(saved.cur) || 0));
    }
    var fresh = -1; /* the item answered in this visit only; it alone plays the check animation */
    function hit(i) { return st.picks[i] === items[i].getAttribute('data-answer'); }
    function answerLabel(i) {
      var b = $('.pick[data-pick="' + items[i].getAttribute('data-answer') + '"]', items[i]);
      return b ? b.textContent : '';
    }
    function label(i) {
      var b = $('.pick[data-pick="' + st.picks[i] + '"]', items[i]);
      return b ? b.textContent : '';
    }
    /* a pick marked data-unsure ("Not sure") is neither right nor wrong: no check, no x, no "lean" line */
    function unsure(i) {
      var b = $('.pick[data-pick="' + st.picks[i] + '"]', items[i]);
      return !!(b && b.hasAttribute('data-unsure'));
    }
    function reveal(i) {
      var rv = $('.qs-reveal', items[i]);
      if (!rv) return '';
      var c = rv.cloneNode(true);
      if (hit(i) || unsure(i)) $$('.lean', c).forEach(function (n) { n.parentNode.removeChild(n); });
      return c.innerHTML;
    }
    function paint(i) {
      var it = items[i], pick = st.picks[i];
      it.classList.toggle('cur', i === st.cur);
      var q = $('.qs-text', it);
      if (q) q.classList.toggle('em-current', i === st.cur);
      $$('.pick', it).forEach(function (b) {
        b.setAttribute('aria-pressed', String(b.getAttribute('data-pick') === pick));
        b.disabled = pick !== null;
      });
    }
    function render() {
      items.forEach(function (it, i) { paint(i); });
      done.innerHTML = '';
      var latest = st.cur - 1;
      items.slice(0, st.cur).forEach(function (it, i) {
        var ok = !record && hit(i);
        var li = document.createElement('li');
        li.className = 'qs-sum em-done' +
          (i === latest ? ' latest' : '') +
          (i === fresh ? ' fresh' : '') +
          (record || unsure(i) ? '' : (ok ? ' ok' : ' no'));
        var qq = document.createElement('span'); qq.className = 'qs-q'; qq.textContent = it.getAttribute('data-short');
        var you = document.createElement('span'); you.className = 'qs-you'; you.textContent = label(i);
        /* when someone differs from us, show our answer too */
        var ours = (!record && !ok && !unsure(i)) ? answerLabel(i) : '';
        if (ours) { var o = document.createElement('span'); o.className = 'qs-ours'; o.textContent = 'Our answer: ' + ours; you.appendChild(document.createElement('br')); you.appendChild(o); }
        li.appendChild(qq); li.appendChild(you);
        if (!record && !unsure(i)) {
          var mk = document.createElement('span');
          mk.className = 'qs-mark';
          mk.setAttribute('role', 'img');
          mk.setAttribute('aria-label', ok ? 'Matched' : 'Differed');
          mk.innerHTML = ok ? CHECK_SVG : CROSS_SVG;
          li.appendChild(mk);
        }
        if (!record && i === latest && reveal(i)) {
          var note = document.createElement('div');
          note.className = 'qs-note';
          note.innerHTML = reveal(i);
          li.appendChild(note);
        }
        done.appendChild(li);
      });
      done.hidden = st.cur === 0;
      var answered = st.picks.filter(function (p) { return p !== null; }).length;
      count.textContent = st.cur >= items.length ? 'All ' + items.length + ' answered.' : answered + ' of ' + items.length + ' answered.';
      if (doneEl) doneEl.hidden = answered < items.length;
    }
    items.forEach(function (it, i) {
      $$('.pick', it).forEach(function (b) {
        b.addEventListener('click', function () {
          if (i !== st.cur || st.picks[i] !== null) return;
          st.picks[i] = b.getAttribute('data-pick');
          st.cur = Math.min(items.length, st.cur + 1);
          storeJSON(key, st);
          fresh = i;
          render();
          fresh = -1;
          var nf = st.cur < items.length ? $('.pick', items[st.cur]) : null;
          if (nf) nf.focus({ preventScroll: true });
        });
      });
    });
    render();
  }

  /* ---------- session 4: support desk. Team and role pickers, then the Rules or Tickets view ---------- */
  function initSupport() {
    var root = $('[data-widget="support"]');
    if (!root) return;
    var TEAM = 'wb:s4:team', ROLE = 'wb:s4:role';
    var team = load(TEAM), role = load(ROLE);
    if (team !== 'flow' && team !== 'guide') team = null;
    if (role !== 'rules' && role !== 'tickets') role = null;
    var editing = !(team && role);
    var pickers = $('.s4-pickers', root), view = $('.s4-view', root);
    var flow = $('.s4-flow', root), brief = $('.s4-brief', root), tickets = $('.s4-tickets', root);
    function render() {
      $$('[data-team]', root).forEach(function (b) {
        b.setAttribute('aria-pressed', String(b.getAttribute('data-team') === team));
      });
      $$('[data-role]', root).forEach(function (b) {
        b.setAttribute('aria-pressed', String(b.getAttribute('data-role') === role));
      });
      var showView = !!(team && role) && !editing;
      pickers.hidden = showView;
      view.hidden = !showView;
      flow.hidden = !(role === 'rules' && team === 'flow');
      brief.hidden = !(role === 'rules' && team === 'guide');
      tickets.hidden = role !== 'tickets';
    }
    $$('[data-team]', root).forEach(function (b) {
      b.addEventListener('click', function () {
        team = b.getAttribute('data-team');
        store(TEAM, team);
        if (team && role) editing = false;
        render();
      });
    });
    $$('[data-role]', root).forEach(function (b) {
      b.addEventListener('click', function () {
        role = b.getAttribute('data-role');
        store(ROLE, role);
        if (team && role) editing = false;
        render();
      });
    });
    var sw = $('.s4-switch', root);
    if (sw) sw.addEventListener('click', function () { editing = true; render(); });
    render();
  }

  /* ---------- session 4: debrief. "You chose" appears only if this browser played the Tickets role ---------- */
  function initDebrief() {
    var sec = $('[data-widget="debrief"]');
    if (!sec) return;
    var table = $('.s4-table', sec);
    var WORD = { refund: 'Refund', replace: 'Replace', deny: 'Deny', cancel: 'Cancel subscription', person: 'Send to a person' };
    function refresh() {
      var saved = loadJSON('wb:s4:q:tickets', null);
      var picks = saved && Array.isArray(saved.picks) ? saved.picks : [];
      table.classList.toggle('played', picks.some(function (p) { return p !== null; }));
      $$('td.s4-you', sec).forEach(function (cell, i) {
        cell.textContent = picks[i] ? (WORD[picks[i]] || '') : '';
      });
    }
    refresh();
    if (window.MutationObserver) {
      new MutationObserver(function () {
        if (sec.classList.contains('on')) refresh();
      }).observe(sec, { attributes: true, attributeFilter: ['class'] });
    }
  }

  /* ---------- home: clear saved answers ---------- */
  function initHome() {
    var clr = $('#clear');
    if (!clr) return;
    var label = clr.textContent;
    clr.addEventListener('click', function () {
      if (!window.confirm('Clear my saved answers?')) return;
      clearWb();
      clr.textContent = 'Cleared';
      setTimeout(function () { clr.textContent = label; }, 1600);
    });
  }

  /* Key takeaways checklist: people tick what feels clear; saved per session. */
  function initTake(root) {
    var sess = body.getAttribute('data-session') || '0';
    var key = 'wb:s' + sess + ':take';
    var saved = loadJSON(key, []);
    if (!Array.isArray(saved)) saved = [];
    var boxes = $$('input[type="checkbox"]', root);
    boxes.forEach(function (b, i) {
      b.checked = saved.indexOf(i) > -1;
      b.closest('li').classList.toggle('is-checked', b.checked);
      b.addEventListener('change', function () {
        var li = b.closest('li');
        li.classList.toggle('is-checked', b.checked);
        if (b.checked) { li.classList.remove('pop'); void li.offsetWidth; li.classList.add('pop'); }
        storeJSON(key, boxes.map(function (x, k) { return x.checked ? k : -1; }).filter(function (k) { return k > -1; }));
      });
    });
  }

  if (!app) document.documentElement.classList.remove('preload');
  buildInit();
  function buildInit() {
    initBuild();
    $$('.qs').forEach(initQs);
    initPerm();
    initMerge();
    initHome();
    initSupport();
    initDebrief();
    $$('.take').forEach(initTake);
  }
})();
