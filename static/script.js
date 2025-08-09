const workersDiv = document.getElementById("workers");
const addBtn = document.getElementById("addWorker");
const totalEl = document.getElementById("totalHours");

let workerCount = 0;
const MAX = 5;

function addWorker() {
  if (workerCount >= MAX) return;

  workerCount += 1;
  const i = workerCount;

  const wrap = document.createElement("div");
  wrap.className = "worker";

  wrap.innerHTML = `
    <label>Vorname
      <input type="text" name="vorname${i}" required />
    </label>
    <label>Nachname
      <input type="text" name="nachname${i}" required />
    </label>
    <label>Ausweis-Nr.
      <input type="text" name="ausweis${i}" required />
    </label>
    <label>Beginn
      <input type="time" name="beginn${i}" required />
    </label>
    <label>Ende
      <input type="time" name="ende${i}" required />
    </label>
  `;

  // óraszám újraszámolás változásra
  wrap.querySelectorAll('input[type="time"]').forEach(inp => {
    inp.addEventListener("change", recompute);
    inp.addEventListener("input", recompute);
  });

  workersDiv.appendChild(wrap);
  recompute();
}

addBtn.addEventListener("click", addWorker);

// első dolgozó alapból
addWorker();

function toMinutes(t) {
  if (!t) return null;
  const [h, m] = t.split(":").map(Number);
  return h * 60 + m;
}

function overlap(a1, a2, b1, b2) {
  // metszet hossza (perc)
  const s = Math.max(a1, b1);
  const e = Math.min(a2, b2);
  return Math.max(0, e - s);
}

function recompute() {
  let totalMin = 0;

  const blocks = [
    [9*60+0, 9*60+15],   // 09:00–09:15
    [12*60+0, 12*60+45], // 12:00–12:45
  ];

  for (let i = 1; i <= workerCount; i++) {
    const b = document.querySelector(`input[name=beginn${i}]`)?.value || "";
    const e = document.querySelector(`input[name=ende${i}]`)?.value || "";
    const start = toMinutes(b);
    const end = toMinutes(e);
    if (start == null || end == null || end <= start) continue;

    let minutes = end - start;
    // szünetek levonása (csak amennyit tényleg metszenek)
    for (const [s, t] of blocks) {
      minutes -= overlap(start, end, s, t);
    }
    totalMin += Math.max(0, minutes);
  }

  const hours = (totalMin / 60).toFixed(2);
  totalEl.textContent = hours;
}
