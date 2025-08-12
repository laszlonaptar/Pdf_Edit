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

  function overlap(a1, a2, b1, b2) {
    const s = Math.max(a1, b1);
    const e = Math.min(a2, b2);
    return Math.max(0, e - s);
  }

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
    return (Math.round(h * 100) / 100).toFixed(2);
  }

  function digitsOnly(input) {
    input.addEventListener("input", () => {
      input.value = input.value.replace(/\D/g, "");
    });
  }

  function enforceNumericKeyboard(input) {
    input.setAttribute("inputmode", "numeric");
    input.setAttribute("pattern", "[0-9]*");
    digitsOnly(input);
  }

  // --- 15 perces "snap" + datalist ---
  function snapToQuarter(inp) {
    const v = inp.value;
    if (!/^\d{2}:\d{2}$/.test(v)) return;
    let [h, m] = v.split(":").map(Number);
    let q = Math.round(m / 15) * 15;
    if (q === 60) { h = (h + 1) % 24; q = 0; }
    const newVal = String(h).padStart(2, "0") + ":" + String(q).padStart(2, "0");
    if (newVal !== v) inp.value = newVal;
  }

  function buildQuarterDatalistFor(inp) {
    if (inp.hasAttribute("list")) return; // már van
    const id = `times_${inp.name}`;
    const dl = document.createElement("datalist");
    dl.id = id;
    // 00:00..23:45 15 perces lépésekkel (96 opció)
    for (let h = 0; h < 24; h++) {
      for (let m = 0; m < 60; m += 15) {
        const val = String(h).padStart(2, "0") + ":" + String(m).padStart(2, "0");
        const opt = document.createElement("option");
        opt.value = val;
        dl.appendChild(opt);
      }
    }
    document.body.appendChild(dl);
    inp.setAttribute("list", id);
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

  // ---- SYNC LOGIKA ----
  function markSynced(inp, isSynced) {
    if (!inp) return;
    if (isSynced) inp.dataset.synced = "1";
    else delete inp.dataset.synced;
  }
  function isSynced(inp) {
    return !!(inp && inp.dataset.synced === "1");
  }
  function setupManualEditUnsync(inp) {
    if (!inp) return;
    const unsync = () => markSynced(inp, false);
    inp.addEventListener("input", (e) => { if (e.isTrusted) unsync(); });
    inp.addEventListener("change", (e) => { if (e.isTrusted) unsync(); });
  }
  function syncFromFirst() {
    const firstBeg = document.querySelector('input[name="beginn1"]')?.value || "";
    const firstEnd = document.querySelector('input[name="ende1"]')?.value || "";
    if (!firstBeg && !firstEnd) return;
    const workers = Array.from(workerList.querySelectorAll(".worker"));
    for (let i = 1; i < workers.length; i++) {
      const fs = workers[i];
      const beg = fs.querySelector('input[name^="beginn"]');
      const end = fs.querySelector('input[name^="ende"]');
      if (beg && (isSynced(beg) || !beg.value)) { beg.value = firstBeg; markSynced(beg, true); snapToQuarter(beg); }
      if (end && (isSynced(end) || !end.value)) { end.value = firstEnd; markSynced(end, true); snapToQuarter(end); }
    }
    recalcAll();
  }

  function wireWorker(workerEl) {
    // Ausweis csak szám
    const ausweis = workerEl.querySelector('input[name^="ausweis"]');
    if (ausweis) enforceNumericKeyboard(ausweis);

    // időmezők beállítása: 15 perc + snap + datalist
    ["beginn", "ende"].forEach(prefix => {
      const inp = workerEl.querySelector(`input[name^="${prefix}"]`);
      if (inp) {
        inp.step = 900;            // 15 * 60
        inp.min = "00:00";
        inp.max = "23:45";
        buildQuarterDatalistFor(inp);

        // élőben is igyekszünk 15-re kerekíteni
        const snapAndRecalc = () => { snapToQuarter(inp); recalcAll(); };
        inp.addEventListener("input", snapAndRecalc);
        inp.addEventListener("change", snapAndRecalc);
        inp.addEventListener("blur",   snapAndRecalc);
      }
    });

    // sync státusz
    const idx = workerEl.getAttribute("data-index");
    const b = workerEl.querySelector(`input[name="beginn${idx}"]`);
    const e = workerEl.querySelector(`input[name="ende${idx}"]`);

    if (idx === "1") {
      if (b) markSynced(b, false);
      if (e) markSynced(e, false);
      b?.addEventListener("input", syncFromFirst);
      e?.addEventListener("input", syncFromFirst);
      b?.addEventListener("change", syncFromFirst);
      e?.addEventListener("change", syncFromFirst);
    } else {
      setupManualEditUnsync(b);
      setupManualEditUnsync(e);
    }
  }

  // első dolgozó bekötése + számítás
  wireWorker(workerList.querySelector(".worker"));
  recalcAll();

  // új dolgozó
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
          <input name="beginn${idx}" type="time" />
        </div>
        <div class="field">
          <label>Ende</label>
          <input name="ende${idx}" type="time" />
        </div>
        <div class="field">
          <label>Stunden (auto)</label>
          <input class="stunden-display" type="text" value="" readonly />
        </div>
      </div>
    `;
    workerList.appendChild(tpl);

    // előtöltés az elsőből + synced státusz
    const firstBeg = document.querySelector('input[name="beginn1"]')?.value || "";
    const firstEnd = document.querySelector('input[name="ende1"]')?.value || "";
    const begNew = tpl.querySelector(`input[name="beginn${idx}"]`);
    const endNew = tpl.querySelector(`input[name="ende${idx}"]`);
    if (firstBeg && begNew) { begNew.value = firstBeg; markSynced(begNew, true); }
    if (firstEnd && endNew) { endNew.value = firstEnd; markSynced(endNew, true); }

    wireWorker(tpl);
    // biztos, ami biztos: snap az új mezőkre is
    begNew && snapToQuarter(begNew);
    endNew && snapToQuarter(endNew);
    recalcAll();
  });

  // első dolgozó ideje változik -> összeg frissítése
  const b1 = document.querySelector('input[name="beginn1"]');
  const e1 = document.querySelector('input[name="ende1"]');
  [b1, e1].forEach(inp => {
    if (inp) {
      inp.addEventListener("input", recalcAll);
      inp.addEventListener("change", recalcAll);
    }
  });
});

// --- Beschreibung karakterszámláló ---
(function () {
  const besch = document.getElementById('beschreibung');
  const out   = document.getElementById('besch-count');
  if (!besch || !out) return;

  const max = parseInt(besch.getAttribute('maxlength') || '1000', 10);

  function updateBeschCount() {
    const len = besch.value.length || 0;
    out.textContent = `${len} / ${max}`;
  }

  updateBeschCount();
  besch.addEventListener('input', updateBeschCount);
  besch.addEventListener('change', updateBeschCount);
})();

// ==== Zusätzliche Validierung beim Absenden ====
(function () {
  const form = document.getElementById("ln-form");
  if (!form) return;

  function trim(v) { return (v || "").toString().trim(); }

  function readWorker(fs) {
    const idx = fs.getAttribute("data-index") || "";
    const q = (sel) => fs.querySelector(sel);
    return {
      idx,
      vorname: trim((q(`input[name="vorname${idx}"]`) || {}).value),
      nachname: trim((q(`input[name="nachname${idx}"]`) || {}).value),
      ausweis: trim((q(`input[name="ausweis${idx}"]`) || {}).value),
      beginn: trim((q(`input[name="beginn${idx}"]`) || {}).value),
      ende: trim((q(`input[name="ende${idx}"]`) || {}).value),
    };
  }

  form.addEventListener("submit", function (e) {
    const errors = [];
    const datum = trim((document.getElementById("datum") || {}).value);
    const bau   = trim((document.getElementById("bau") || {}).value);
    const bf    = trim((document.getElementById("basf_beauftragter") || {}).value);
    const besch = trim((document.getElementById("beschreibung") || {}).value);

    if (!datum) errors.push("Bitte das Datum der Leistungsausführung angeben.");
    if (!bau) errors.push("Bitte Bau und Ausführungsort ausfüllen.");
    if (!bf) errors.push("Bitte den BASF-Beauftragten (Org.-Code) ausfüllen.");
    if (!besch) errors.push("Bitte die Beschreibung der ausgeführten Arbeiten ausfüllen.");

    const list = document.getElementById("worker-list");
    const sets = Array.from(list.querySelectorAll(".worker"));
    let validWorkers = 0;

    sets.forEach((fs) => {
      const w = readWorker(fs);
      const anyFilled = !!(w.vorname || w.nachname || w.ausweis || w.beginn || w.ende);
      const allCore   = !!(w.vorname && w.nachname && w.ausweis && w.beginn && w.ende);
      if (allCore) validWorkers += 1;
      else if (anyFilled) {
        const missing = [];
        if (!w.vorname) missing.push("Vorname");
        if (!w.nachname) missing.push("Nachname");
        if (!w.ausweis) missing.push("Ausweis-Nr.");
        if (!w.beginn) missing.push("Beginn");
        if (!w.ende) missing.push("Ende");
        errors.push(`Bitte Mitarbeiter ${w.idx}: ${missing.join(", ")} ausfüllen.`);
      }
    });

    if (validWorkers === 0) {
      errors.push("Bitte mindestens einen Mitarbeiter vollständig angeben (Vorname, Nachname, Ausweis-Nr., Beginn, Ende).");
    }

    if (errors.length) {
      e.preventDefault();
      alert(errors.join("\n"));
      return false;
    }
    return true;
  });
})();
