document.addEventListener("DOMContentLoaded", () => {
  const addBtn = document.getElementById("add-worker");
  const workerList = document.getElementById("worker-list");
  const totalOut = document.getElementById("gesamtstunden_auto");
  const MAX_WORKERS = 5;

  // --- helpers ---
  function toTime(s) {
    if (!s || !/^\d{2}:\d{2}$/.test(s)) return null;
    const [hh, mm] = s.split(":").map(Number);
    if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
    return { hh, mm };
  }
  function minutes(t) { return t.hh * 60 + t.mm; }

  // overlap of [a1,a2) with [b1,b2) in minutes
  function overlap(a1, a2, b1, b2) {
    const s = Math.max(a1, b1);
    const e = Math.min(a2, b2);
    return Math.max(0, e - s);
  }

  // 09:00–09:15 (15’) és 12:00–12:45 (45’)
  const BREAKS = [
    { start: 9 * 60 + 0,  end: 9 * 60 + 15 },
    { start: 12 * 60 + 0, end: 12 * 60 + 45 },
  ];

  function hoursWithBreaks(begStr, endStr) {
    const bt = toTime(begStr);
    const et = toTime(endStr);
    if (!bt || !et) return 0;
    const b = minutes(bt), e = minutes(et);
    if (e <= b) return 0;
    let total = e - b;
    for (const br of BREAKS) total -= overlap(b, e, br.start, br.end);
    return Math.max(0, total) / 60;
  }

  function formatHours(h) {
    // két tizedesre kerekítve, ponttal (backenddel konzisztens)
    return (Math.round(h * 100) / 100).toFixed(2);
  }

  function digitsOnly(input) {
    input.addEventListener("input", () => {
      input.value = input.value.replace(/\D/g, "");
    });
  }

  function recalcWorker(workerEl) {
    const beg = workerEl.querySelector('input[name^="beginn"]')?.value || "";
    const end = workerEl.querySelector('input[name^="ende"]')?.value || "";
    const out = workerEl.querySelector(".stunden-display");
    const h = hoursWithBreaks(beg, end);
    if (out) out.value = h ? formatHours(h) : "";
    return h;
  }

  function recalcAll() {
    let sum = 0;
    workerList.querySelectorAll(".worker").forEach(w => {
      sum += recalcWorker(w);
    });
    if (totalOut) totalOut.value = sum ? formatHours(sum) : "";
  }

  function wireWorker(workerEl) {
    // Ausweis csak szám
    const ausweis = workerEl.querySelector('input[name^="ausweis"]');
    if (ausweis) digitsOnly(ausweis);

    // idő változásra újraszámolás
    ["beginn", "ende"].forEach(prefix => {
      const inp = workerEl.querySelector(`input[name^="${prefix}"]`);
      if (inp) {
        // 15 perces léptetés UX (nem kötelező, de kényelmes)
        inp.step = 60; // percenként – a natív wheel így is 1 perces lehet, de jó így hagyni
        inp.addEventListener("input", recalcAll);
        inp.addEventListener("change", recalcAll);
      }
    });
  }

  // már meglévő első dolgozó bekötése
  wireWorker(workerList.querySelector(".worker"));
  // és első összámítás
  recalcAll();

  // új dolgozó hozzáadása
  addBtn?.addEventListener("click", () => {
    const current = workerList.querySelectorAll(".worker").length;
    if (current >= MAX_WORKERS) return;

    const idx = current + 1;
    const tpl = document.createElement("fieldset");
    tpl.className = "worker";
    tpl.dataset.index = String(idx);
    tpl.innerHTML = `
      <legend>Mitarbeiter ${idx}</legend>
      <div class="grid grid-3">
        <div class="field">
          <label>Vorname</label>
          <input name="vorname${idx}" type="text" />
        </div>
        <div class="field">
          <label>Nachname</label>
          <input name="nachname${idx}" type="text" />
        </div>
        <div class="field">
          <label>Ausweis-Nr. / Kennzeichen</label>
          <input name="ausweis${idx}" type="text" />
        </div>
      </div>
      <div class="grid grid-3">
        <div class="field">
          <label>Beginn</label>
          <input name="beginn${idx}" type="time" step="60" />
        </div>
        <div class="field">
          <label>Ende</label>
          <input name="ende${idx}" type="time" step="60" />
        </div>
        <div class="field">
          <label>Stunden (auto)</label>
          <input class="stunden-display" type="text" value="" readonly />
        </div>
      </div>
    `;
    workerList.appendChild(tpl);

    // ha az első dolgozónál van idő, előtöltjük (kért feature #1 előkészítés)
    const firstBeg = document.querySelector('input[name="beginn1"]')?.value || "";
    const firstEnd = document.querySelector('input[name="ende1"]')?.value || "";
    if (firstBeg) tpl.querySelector(`input[name="beginn${idx}"]`).value = firstBeg;
    if (firstEnd) tpl.querySelector(`input[name="ende${idx}"]`).value = firstEnd;

    wireWorker(tpl);
    recalcAll();
  });

  // ha az első dolgozó ideje változik, frissítjük az összórát
  const b1 = document.querySelector('input[name="beginn1"]');
  const e1 = document.querySelector('input[name="ende1"]');
  [b1, e1].forEach(inp => {
    if (inp) {
      inp.addEventListener("input", recalcAll);
      inp.addEventListener("change", recalcAll);
    }
  });
});
