// v15

// ---- Helpers ----
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

function parseHHMM(s) {
  if (!s || !/^\d{2}:\d{2}/.test(s)) return null;
  const [h, m] = s.split(':').map(n => parseInt(n, 10));
  if (isNaN(h) || isNaN(m)) return null;
  return { h, m };
}

function minutesBetween(a, b) {
  if (!a || !b) return 0;
  const t1 = a.h * 60 + a.m;
  const t2 = b.h * 60 + b.m;
  return Math.max(0, t2 - t1);
}

function overlapMinutes(aBeg, aEnd, bBeg, bEnd) {
  const a1 = aBeg.h * 60 + aBeg.m;
  const a2 = aEnd.h * 60 + aEnd.m;
  const b1 = bBeg.h * 60 + bBeg.m;
  const b2 = bEnd.h * 60 + bEnd.m;
  const start = Math.max(a1, b1);
  const end = Math.min(a2, b2);
  return Math.max(0, end - start);
}

function hoursWithBreaks(beg, end, pauseMin) {
  if (!beg || !end) return 0;
  let total = minutesBetween(beg, end);
  if (total <= 0) return 0;

  // 1h szünet: 9:00-9:15 és 12:00-12:45 ablakok
  if (pauseMin >= 60) {
    const m1 = overlapMinutes(beg, end, {h:9,m:0}, {h:9,m:15});
    const m2 = overlapMinutes(beg, end, {h:12,m:0}, {h:12,m:45});
    total = Math.max(0, total - (m1 + m2));
  } else {
    // fél óra
    total = Math.max(0, total - Math.min(total, 30));
  }
  return total / 60;
}

function toHHMM({h,m}) {
  const hh = String(h).padStart(2, '0');
  const mm = String(m).padStart(2, '0');
  return `${hh}:${mm}`;
}

function roundTo15(value) {
  // value: "HH:MM" -> kerekítés 15 perchez
  const t = parseHHMM(value);
  if (!t) return value;
  const total = t.h * 60 + t.m;
  const rounded = Math.round(total / 15) * 15;
  const h = Math.floor(rounded / 60);
  const m = rounded % 60;
  return toHHMM({h, m});
}

// ---- 1000 char counter ----
(function () {
  const ta = document.getElementById('beschreibung');
  if (!ta) return;
  const cnt = document.getElementById('besch-count');
  const update = () => {
    const len = ta.value.length;
    cnt.textContent = `${len} / 1000`;
  };
  ta.addEventListener('input', update);
  update();
})();

// ---- Break kapcsoló ----
(function () {
  const ck = document.getElementById('break_half');
  const hidden = document.getElementById('break_minutes');
  if (!ck || !hidden) return;
  const sync = () => {
    hidden.value = ck.checked ? '30' : '60';
    recalcAll();
  };
  ck.addEventListener('change', sync);
  sync();
})();

// ---- TIME: 15 perc kényszerítése ----
function enforceStep15(input) {
  // állítsuk be a stepet és kerekítsünk elhagyáskor
  input.setAttribute('step', '900');
  input.addEventListener('blur', () => {
    if (!input.value) return;
    const v = roundTo15(input.value);
    input.value = v;
    recalcAll();
  });
}

$$('.t-beginn, .t-ende').forEach(enforceStep15);

// ---- Autocomplete a workers.csv-ből ----
let WORKERS = [];     // {vorname, nachname, ausweis}
let MAP_AUSWEIS = new Map(); // ausweis -> worker
let MAP_BY_FULLNAME = new Map(); // "vorname|nachname" -> worker
let VORNAMEN_SET = new Set();
let NACHNAMEN_SET = new Set();
let AUSWEISE_SET = new Set();

function normalizeHeader(h) {
  h = (h || '').toLowerCase().trim();
  // lehetséges változatokból kitaláljuk
  if (['vorname','first','keresztnev','firstname','given_name'].includes(h)) return 'vorname';
  if (['nachname','last','vezeteknev','lastname','family_name','surname'].includes(h)) return 'nachname';
  if (['ausweis','kennzeichen','id','badge','igazolvany','card','ausweis_nr','ausweis-nr.'].includes(h)) return 'ausweis';
  return h;
}

function loadWorkers() {
  return fetch('/static/workers.csv', { cache: 'no-store' })
    .then(r => r.text())
    .then(txt => {
      const lines = txt.split(/\r?\n/).filter(l => l.trim().length);
      if (lines.length === 0) return;

      const headers = lines[0].split(',').map(s => s.trim());
      const idx = { vorname: -1, nachname: -1, ausweis: -1 };
      headers.forEach((h, i) => {
        const n = normalizeHeader(h);
        if (n === 'vorname') idx.vorname = i;
        if (n === 'nachname') idx.nachname = i;
        if (n === 'ausweis') idx.ausweis = i;
      });

      // ha nincs fejléc vagy nem ismerjük, próbáljuk default sorrenddel
      const hasHeader = idx.vorname !== -1 || idx.nachname !== -1 || idx.ausweis !== -1;
      const start = hasHeader ? 1 : 0;

      WORKERS = [];
      MAP_AUSWEIS.clear();
      MAP_BY_FULLNAME.clear();
      VORNAMEN_SET.clear();
      NACHNAMEN_SET.clear();
      AUSWEISE_SET.clear();

      for (let i = start; i < lines.length; i++) {
        const parts = lines[i].split(',').map(s => s.trim());
        let vor = idx.vorname !== -1 ? parts[idx.vorname] : parts[0] || '';
        let nach = idx.nachname !== -1 ? parts[idx.nachname] : parts[1] || '';
        let aus = idx.ausweis !== -1 ? parts[idx.ausweis] : parts[2] || '';

        if (!(vor || nach || aus)) continue;
        const w = { vorname: vor, nachname: nach, ausweis: aus };
        WORKERS.push(w);

        if (aus) MAP_AUSWEIS.set(aus, w);
        const key = `${vor}|${nach}`.toLowerCase();
        if (vor && nach) MAP_BY_FULLNAME.set(key, w);

        if (vor) VORNAMEN_SET.add(vor);
        if (nach) NACHNAMEN_SET.add(nach);
        if (aus) AUSWEISE_SET.add(aus);
      }

      // feltöltjük a datalist-eket
      const dlV = document.getElementById('dl_vornamen');
      const dlN = document.getElementById('dl_nachnamen');
      const dlA = document.getElementById('dl_ausweise');
      if (dlV) dlV.innerHTML = Array.from(VORNAMEN_SET).sort().map(v => `<option value="${v}"></option>`).join('');
      if (dlN) dlN.innerHTML = Array.from(NACHNAMEN_SET).sort().map(v => `<option value="${v}"></option>`).join('');
      if (dlA) dlA.innerHTML = Array.from(AUSWEISE_SET).sort().map(v => `<option value="${v}"></option>`).join('');

      bindAllRows();
    })
    .catch(() => {
      // ha bármi gond, a form többi része működjön tovább
    });
}

function bindRow(row) {
  const i = row.dataset.index;
  const fVor = row.querySelector(`[name=vorname${i}]`);
  const fNach = row.querySelector(`[name=nachname${i}]`);
  const fAus = row.querySelector(`[name=ausweis${i}]`);
  const fBeg = row.querySelector(`[name=beginn${i}]`);
  const fEnde = row.querySelector(`[name=ende${i}]`);
  const fStd = row.querySelector(`.stunden-display`);

  // amikor pontos értéket kapunk valamelyik mezőben, próbálunk összekapcsolni
  function tryLinkByAusweis() {
    const w = MAP_AUSWEIS.get((fAus.value || '').trim());
    if (w) {
      if (fVor.value !== w.vorname) fVor.value = w.vorname || '';
      if (fNach.value !== w.nachname) fNach.value = w.nachname || '';
    }
  }

  function tryLinkByNames() {
    const key = `${(fVor.value||'').trim()}|${(fNach.value||'').trim()}`.toLowerCase();
    const w = MAP_BY_FULLNAME.get(key);
    if (w) {
      if (fAus.value !== w.ausweis) fAus.value = w.ausweis || '';
    }
  }

  function onChangeNames() {
    // ha mindkettő megvan, töltsük ki az ausweist
    if ((fVor.value||'').trim() && (fNach.value||'').trim()) {
      tryLinkByNames();
    }
  }

  function onChangeAusweis() {
    if ((fAus.value||'').trim()) {
      tryLinkByAusweis();
    }
  }

  fVor.addEventListener('change', onChangeNames);
  fNach.addEventListener('change', onChangeNames);
  fAus.addEventListener('change', onChangeAusweis);

  // 15 perces kényszer + óraszám számítás
  [fBeg, fEnde].forEach(inp => {
    if (!inp) return;
    inp.setAttribute('step', '900');
    inp.addEventListener('blur', () => {
      if (!inp.value) return;
      inp.value = roundTo15(inp.value);
      recalcAll();
    });
    inp.addEventListener('input', () => {
      // natív pickernél is frissüljön
      recalcAll();
    });
  });

  function recalcRow() {
    const pauseMin = parseInt(document.getElementById('break_minutes').value || '60', 10);
    const beg = parseHHMM(fBeg.value);
    const en  = parseHHMM(fEnde.value);
    const h = hoursWithBreaks(beg, en, pauseMin);
    fStd.value = h ? h.toFixed(2) : '';
    return h || 0;
  }

  row.__recalcRow = recalcRow; // eltároljuk későbbre
}

function bindAllRows() {
  $$('.worker').forEach(bindRow);
  recalcAll();
}

function recalcAll() {
  let sum = 0;
  $$('.worker').forEach(row => {
    if (typeof row.__recalcRow === 'function') {
      sum += row.__recalcRow();
    }
  });
  const out = document.getElementById('gesamtstunden_auto');
  if (out) out.value = sum ? sum.toFixed(2) : '';
}

// induláskor: betöltjük a CSV-t és kötjük az eseményeket
document.addEventListener('DOMContentLoaded', () => {
  // ha a CSV nem érhető el, akkor is kösd a sorokat (CSV nélkül is működik a kerekítés/számítás)
  bindAllRows();
  loadWorkers(); // ez felül fogja tölteni a datalisteket és a bindRow linkelést is életre kelti
});
