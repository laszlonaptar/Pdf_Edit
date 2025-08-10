// ---- Utils ----
function q(sel, root = document) { return root.querySelector(sel); }
function qa(sel, root = document) { return Array.from(root.querySelectorAll(sel)); }

function parseHHMM(str) {
  if (!str) return null;
  const [h, m] = str.split(':').map(Number);
  if (Number.isNaN(h) || Number.isNaN(m)) return null;
  return { h, m };
}
function toMinutes(t) { return t.h * 60 + t.m; }
function overlap(a1, a2, b1, b2) {
  const s = Math.max(a1, b1);
  const e = Math.min(a2, b2);
  return Math.max(0, e - s);
}
function hoursWithBreaks(begStr, endStr) {
  const b = parseHHMM(begStr);
  const e = parseHHMM(endStr);
  if (!b || !e) return 0;
  let B = toMinutes(b), E = toMinutes(e);
  if (E <= B) return 0;
  let total = E - B;
  // breaks: 09:00–09:15 (15m), 12:00–12:45 (45m)
  total -= overlap(B, E, 9*60, 9*60 + 15);
  total -= overlap(B, E, 12*60, 12*60 + 45);
  return Math.max(0, total) / 60;
}
function fmtHours(h) {
  return (Math.round(h * 100) / 100).toFixed(2).replace('.', ',');
}

// ---- Beschreibung counter ----
(function() {
  const ta = q('#beschreibung');
  const cnt = q('#besch-count');
  const update = () => { cnt.textContent = `${ta.value.length} / 1000`; };
  ta.addEventListener('input', update);
  update();
})();

// ---- Workers ----
const workerList = q('#worker-list');
const addBtn = q('#add-worker');
let maxWorkers = 5;

function updateHoursForWorker(fs) {
  const idx = fs.dataset.index;
  const beg = q(`input[name="beginn${idx}"]`, fs).value;
  const end = q(`input[name="ende${idx}"]`, fs).value;
  const hours = hoursWithBreaks(beg, end);
  q('.stunden-display', fs).value = hours ? fmtHours(hours) : '';
}
function updateTotal() {
  let sum = 0;
  qa('.worker').forEach(fs => {
    const idx = fs.dataset.index;
    const beg = q(`input[name="beginn${idx}"]`, fs).value;
    const end = q(`input[name="ende${idx}"]`, fs).value;
    sum += hoursWithBreaks(beg, end);
  });
  q('#gesamtstunden_auto').value = sum ? fmtHours(sum) : '';
}
function attachWorkerEvents(fs) {
  const idx = fs.dataset.index;
  ['beginn', 'ende'].forEach(k => {
    q(`input[name="${k}${idx}"]`, fs).addEventListener('input', () => {
      updateHoursForWorker(fs);
      updateTotal();
    });
  });
}

attachWorkerEvents(q('.worker'));

addBtn.addEventListener('click', () => {
  const current = qa('.worker').length;
  if (current >= maxWorkers) return;

  const next = current + 1;
  const tpl = qa('.worker')[0].cloneNode(true);
  tpl.dataset.index = String(next);
  q('legend', tpl).textContent = `Mitarbeiter ${next}`;

  // Reset values + rename names
  [
    'vorname', 'nachname', 'ausweis',
    'beginn', 'ende'
  ].forEach(key => {
    const inp = q(`input[name^="${key}"]`, tpl);
    inp.name = `${key}${next}`;
    inp.value = '';
  });

  q('.stunden-display', tpl).value = '';
  workerList.appendChild(tpl);
  attachWorkerEvents(tpl);
});

// ---- Default date today ----
(function setDefaultDate() {
  const el = q('#datum');
  if (!el.value) {
    const d = new Date();
    const mm = String(d.getMonth()+1).padStart(2,'0');
    const dd = String(d.getDate()).padStart(2,'0');
    el.value = `${d.getFullYear()}-${mm}-${dd}`;
  }
})();
