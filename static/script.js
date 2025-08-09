// Max 5 Mitarbeiter
const MAX_WORKERS = 5;

const addBtn = document.getElementById("addWorkerBtn");
const moreWorkers = document.getElementById("moreWorkers");
const totalHoursEl = document.getElementById("totalHours");
const form = document.getElementById("reportForm");

function timeToMinutes(t) {
  if (!t) return null;
  const [h, m] = t.split(":").map(Number);
  return h * 60 + m;
}

function overlapMinutes(aStart, aEnd, bStart, bEnd) {
  const start = Math.max(aStart, bStart);
  const end = Math.min(aEnd, bEnd);
  return Math.max(0, end - start);
}

function calcTotalHours() {
  // összesített idő minden dolgozóra (munkaidő tartományok unióját egyszerűsítve: összeadjuk az egyes dolgozókat)
  let totalMinutes = 0;

  // Pausenzeiten percben
  const pauseMorningStart = 9 * 60;
  const pauseMorningEnd   = 9 * 60 + 15;   // 09:15
  const pauseLunchStart   = 12 * 60;
  const pauseLunchEnd     = 12 * 60 + 45;  // 12:45

  const workerBlocks = document.querySelectorAll(".worker");
  workerBlocks.forEach(block => {
    const i = block.dataset.index;
    const beg = document.querySelector(`input[name="beginn${i}"]`)?.value || "";
    const end = document.querySelector(`input[name="ende${i}"]`)?.value || "";

    const bMin = timeToMinutes(beg);
    const eMin = timeToMinutes(end);
    if (bMin == null || eMin == null || eMin <= bMin) return;

    let work = eMin - bMin;

    // Pausenabzug csak akkor, ha a tartomány metszi a szünetet
    work -= overlapMinutes(bMin, eMin, pauseMorningStart, pauseMorningEnd);
    work -= overlapMinutes(bMin, eMin, pauseLunchStart,   pauseLunchEnd);

    totalMinutes += Math.max(0, work);
  });

  totalHoursEl.value = (totalMinutes / 60).toFixed(2).replace(".", ",");
}

function currentWorkerCount() {
  return document.querySelectorAll(".worker").length;
}

function createWorker(index) {
  const div = document.createElement("div");
  div.className = "worker";
  div.dataset.index = index;
  div.innerHTML = `
    <h3>Mitarbeiter ${index}</h3>
    <div class="grid">
      <label>Vorname
        <input type="text" name="vorname${index}" required />
      </label>
      <label>Nachname
        <input type="text" name="nachname${index}" required />
      </label>
      <label>Ausweis-Nr.
        <input type="text" name="ausweis${index}" required />
      </label>
      <label>Beginn
        <input type="time" name="beginn${index}" required />
      </label>
      <label>Ende
        <input type="time" name="ende${index}" required />
      </label>
    </div>
  `;
  return div;
}

// Dinamikus hozzáadás
addBtn?.addEventListener("click", () => {
  const count = currentWorkerCount();
  if (count >= MAX_WORKERS) {
    alert(`Max. ${MAX_WORKERS} Mitarbeiter.`);
    return;
  }
  const next = count + 1;
  moreWorkers.appendChild(createWorker(next));
});

// Realtime összóra
form.addEventListener("input", (e) => {
  if (e.target.type === "time" || e.target.name?.startsWith("beginn") || e.target.name?.startsWith("ende")) {
    calcTotalHours();
  }
});

// első betöltéskor is próbáljuk kiszámolni (ha a böngésző visszatölt értékeket)
window.addEventListener("DOMContentLoaded", calcTotalHours);
