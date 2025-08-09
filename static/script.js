const workersDiv = document.getElementById('workers');
const addBtn = document.getElementById('addWorkerBtn');
const totalEl = document.getElementById('totalHours');

const BREAKS = [
  { start: "09:00", end: "09:15" },
  { start: "12:00", end: "12:45" },
];

let workerCount = 0;
const MAX_WORKERS = 5;

function addWorker(prefill = {}) {
  if (workerCount >= MAX_WORKERS) return;
  workerCount++;

  const wrap = document.createElement('div');
  wrap.className = 'row';
  wrap.dataset.index = workerCount;

  wrap.innerHTML = `
    <input class="wide" name="mitarbeiter_vorname${workerCount === 1 ? '' : workerCount}" placeholder="Vezetéknév" value="${prefill.vorname || ''}" required>
    <input class="wide" name="mitarbeiter_nachname${workerCount === 1 ? '' : workerCount}" placeholder="Keresztnév" value="${prefill.nachname || ''}" required>
    <input class="small" name="ausweis${workerCount === 1 ? '' : workerCount}" placeholder="Igazolvány" value="${prefill.ausweis || ''}" required>
    <input class="small" type="time" name="beginn${workerCount === 1 ? '' : workerCount}" value="${prefill.beginn || ''}" required>
    <input class="small" type="time" name="ende${workerCount === 1 ? '' : workerCount}" value="${prefill.ende || ''}" required>
  `;

  workersDiv.appendChild(wrap);
  attachTimeHandlers(wrap);
  recomputeTotal();
}

function parseMinutes(hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  return h*60 + m;
}

function overlapMinutes(aStart, aEnd, bStart, bEnd) {
  const L = Math.max(aStart, bStart);
  const R = Math.min(aEnd, bEnd);
  return Math.max(0, R - L);
}

function computeOne(begin, end) {
  if (!begin || !end) return 0;
  const s = parseMinutes(begin);
  const e = parseMinutes(end);
  if (e <= s) return 0;
  let mins = e - s;
  for (const br of BREAKS) {
    mins -= overlapMinutes(s, e, parseMinutes(br.start), parseMinutes(br.end));
  }
  return Math.max(0, mins) / 60.0;
}

function attachTimeHandlers(row) {
  row.addEventListener('change', (e) => {
    if (e.target.name.startsWith('beginn') || e.target.name.startsWith('ende')) {
      recomputeTotal();
    }
  });
}

function recomputeTotal() {
  let total = 0;
  const rows = Array.from(workersDiv.children);
  for (const row of rows) {
    const beg = row.querySelector('input[name^="beginn"]').value;
    const end = row.querySelector('input[name^="ende"]').value;
    total += computeOne(beg, end);
  }
  totalEl.textContent = total.toFixed(2);
}

addBtn.addEventListener('click', () => addWorker());

// Start with one required worker
addWorker();
