document.addEventListener("DOMContentLoaded", () => {
  const addBtn = document.getElementById("add-worker");
  const workerList = document.getElementById("worker-list");
  const totalOut = document.getElementById("gesamtstunden_auto");
  const MAX_WORKERS = 5;

  // ===== AUTOCOMPLETE: dolgozói lista betöltése CSV-ből =====
  let WORKERS = [];
  let byAusweis = new Map();
  let byFullName = new Map(); // "nachname|vorname" (lowercase, trimmed)

  function norm(s) { return (s || "").toString().trim(); }
  function keyName(nach, vor) { return `${norm(nach).toLowerCase()}|${norm(vor).toLowerCase()}`; }

  async function loadWorkers() {
    try {
      const resp = await fetch("/static/workers.csv", { cache: "no-store" });
      if (!resp.ok) return;
      let text = await resp.text();
      parseCSV(text);
    } catch (_) {}
  }

  // --- CSV parser: automatikus ; / , felismerés + idézőjelek kezelése
  function parseCSV(text) {
    if (text && text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const linesRaw = text.split(/\r?\n/).filter(l => l.trim().length);
    if (!linesRaw.length) return;

    const detectDelim = (s) => {
      const sc = (s.match(/;/g) || []).length;
      const cc = (s.match(/,/g) || []).length;
      return sc > cc ? ";" : ",";
    };
    const delim = detectDelim(linesRaw[0]);

    const splitCSV = (line) => {
      const re = new RegExp(`${delim}(?=(?:[^"]*"[^"]*")*[^"]*$)`, "g");
      return line.split(re).map(x => {
        x = x.trim();
        if (x.startsWith('"') && x.endsWith('"')) x = x.slice(1, -1);
        return x.replace(/""/g, '"');
      });
    };

    const header = splitCSV(linesRaw[0]).map(h => h.toLowerCase().trim());
    const idxNach = header.findIndex(h => h === "nachname" || h === "name"); // „Name” is ok
    const idxVor  = header.findIndex(h => h === "vorname");
    const idxAus  = header.findIndex(h => /(ausweis|kennzeichen)/.test(h));
    if (idxNach < 0 || idxVor < 0 || idxAus < 0) return;

    WORKERS = [];
    byAusweis.clear();
    byFullName.clear();

    for (let i = 1; i < linesRaw.length; i++) {
      const cols = splitCSV(linesRaw[i]);
      if (cols.length <= Math.max(idxNach, idxVor, idxAus)) continue;
      const w = {
        nachname: norm(cols[idxNach]),
        vorname:  norm(cols[idxVor]),
        ausweis:  norm(cols[idxAus]),
      };
      if (!w.nachname && !w.vorname && !w.ausweis) continue;
      WORKERS.push(w);
      if (w.ausweis) byAusweis.set(w.ausweis, w);
      byFullName.set(keyName(w.nachname, w.vorname), w);
    }

    // már kirakott inputok automatikus frissítése
    refreshAllAutocompletes();
  }

  // ===== Egyedi lenyíló autocomplete (body-hoz csatolva, nem vágódik le) =====
  function makeAutocomplete(input, getOptions, onPick) {
    const dd = document.createElement("div");
    dd.style.position = "absolute";
    dd.style.zIndex = "99999";
    dd.style.background = "white";
    dd.style.border = "1px solid #ddd";
    dd.style.borderTop = "none";
    dd.style.maxHeight = "220px";
    dd.style.overflowY = "auto";
    dd.style.display = "none";
    dd.style.boxShadow = "0 6px 14px rgba(0,0,0,0.08)";
    dd.style.borderRadius = "0 0 .5rem .5rem";
    dd.style.fontSize = "14px";
    dd.setAttribute("role", "listbox");
    document.body.appendChild(dd);

    function hide(){ dd.style.display = "none"; }
    function show(){ dd.style.display = dd.children.length ? "block" : "none"; }

    function position() {
      const r = input.getBoundingClientRect();
      dd.style.left = `${window.scrollX + r.left}px`;
      dd.style.top  = `${window.scrollY + r.bottom}px`;
      dd.style.width = `${r.width}px`;
    }

    function render(list) {
      dd.innerHTML = "";
      list.slice(0, 12).forEach(item => {
        const opt = document.createElement("div");
        opt.textContent = item.label ?? item.value;
        opt.dataset.value = item.value ?? item.label ?? "";
        opt.style.padding = ".5rem .75rem";
        opt.style.cursor  = "pointer";
        opt.addEventListener("mousedown", (e) => {
          e.preventDefault();
          input.value = opt.dataset.value;
          hide();
          onPick?.(opt.dataset.value, item);
          input.dispatchEvent(new Event("change", { bubbles: true }));
        });
        opt.addEventListener("mouseover", () => opt.style.background = "#f5f5f5");
        opt.addEventListener("mouseout",  () => opt.style.background = "white");
        dd.appendChild(opt);
      });
      position();
      show();
    }

    function filterOptions() {
      const q = input.value.toLowerCase().trim();
      const raw = getOptions();
      if (!q) { hide(); return; }
      const list = raw
        .filter(v => (v.label ?? v).toLowerCase().includes(q))
        .map(v => (typeof v === "string" ? { value: v } : v));
      render(list);
    }

    input.addEventListener("input", filterOptions);
    input.addEventListener("focus", () => { position(); filterOptions(); });
    input.addEventListener("blur",  () => setTimeout(hide, 120));

    // görgetés/átméretezés esetén maradjon az input alatt
    window.addEventListener("scroll", position, true);
    window.addEventListener("resize", position);
  }

  function refreshAllAutocompletes() {
    document.querySelectorAll("#worker-list .worker").forEach(setupAutocomplete);
  }

  function setupAutocomplete(fs) {
    const idx = fs.getAttribute("data-index");
    const inpNach = fs.querySelector(`input[name="nachname${idx}"]`);
    const inpVor  = fs.querySelector(`input[name="vorname${idx}"]`);
    const inpAus  = fs.querySelector(`input[name="ausweis${idx}"]`);
    if (!inpNach || !inpVor || !inpAus) return;

    // Vorname
    makeAutocomplete(
      inpVor,
      () => [...new Set(WORKERS.map(w => w.vorname).filter(Boolean))],
      (value) => {
        const w = byFullName.get(keyName(inpNach.value, value));
        if (w) inpAus.value = w.ausweis;
      }
    );

    // Nachname
    makeAutocomplete(
      inpNach,
      () => [...new Set(WORKERS.map(w => w.nachname).filter(Boolean))],
      (value) => {
        const w = byFullName.get(keyName(value, inpVor.value));
        if (w) inpAus.value = w.ausweis;
      }
    );

    // Ausweis
    makeAutocomplete(
      inpAus,
      () => WORKERS.map(w => ({ value: w.ausweis, label: `${w.ausweis} – ${w.nachname} ${w.vorname}` })),
      (value) => {
        const w = byAusweis.get(value);
        if (w) { inpNach.value = w.nachname; inpVor.value = w.vorname; }
      }
    );

    // kézi módosításokból következtetés
    inpAus.addEventListener("change", () => {
      const w = byAusweis.get(norm(inpAus.value));
      if (w) { inpNach.value = w.nachname; inpVor.value = w.vorname; }
    });
    const tryNames = () => {
      const w = byFullName.get(keyName(inpNach.value, inpVor.value));
      if (w) inpAus.value = w.ausweis;
    };
    inpNach.addEventListener("change", tryNames);
    inpVor.addEventListener("change",  tryNames);
  }
  // ===========================================================

  // --- helpers (óra-számítás, stb.) ---
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
  function formatHours(h) { return (Math.round(h * 100) / 100).toFixed(2); }
  function digitsOnly(input) { input.addEventListener("input", () => { input.value = input.value.replace(/\D/g, ""); }); }
  function enforceNumericKeyboard(input) {
    input.setAttribute("inputmode", "numeric");
    input.setAttribute("pattern", "[0-9]*");
    digitsOnly(input);
  }

  function snapToQuarter(inp) {
    const v = inp.value;
    if (!/^\d{2}:\d{2}$/.test(v)) return;
    let [h, m] = v.split(":").map(Number);
    let q = Math.round(m / 15) * 15;
    if (q === 60) { h = (h + 1) % 24; q = 0; }
    const nv = String(h).padStart(2, "0") + ":" + String(q).padStart(2, "0");
    if (nv !== v) inp.value = nv;
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
    workerList.querySelectorAll(".worker").forEach(w => { sum += recalcWorker(w); });
    if (totalOut) totalOut.value = sum ? formatHours(sum) : "";
  }

  // ---- SYNC: első dolgozó ideje többiekhez (amíg nem írják át) ----
  function markSynced(inp, isSynced) { if (!inp) return; isSynced ? inp.dataset.synced = "1" : delete inp.dataset.synced; }
  function isSynced(inp) { return !!(inp && inp.dataset.synced === "1"); }
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
    const idx = workerEl.getAttribute("data-index");

    // autocomplete a 3 mezőre
    setupAutocomplete(workerEl);

    // Ausweis csak szám + numerikus billentyűzet
    const ausweis = workerEl.querySelector(`input[name="ausweis${idx}"]`);
    if (ausweis) enforceNumericKeyboard(ausweis);

    // idő mezők: snap + recalculations
    ["beginn", "ende"].forEach(prefix => {
      const inp = workerEl.querySelector(`input[name^="${prefix}"]`);
      if (inp) {
        inp.addEventListener("change", () => { snapToQuarter(inp); recalcAll(); });
        inp.addEventListener("blur",   () => { snapToQuarter(inp); recalcAll(); });
        inp.addEventListener("input",  recalcAll);
      }
    });

    const b = workerEl.querySelector(`input[name="beginn${idx}"]`);
    const e = workerEl.querySelector(`input[name="ende${idx}"]`);

    if (idx === "1") {
      markSynced(b, false); markSynced(e, false);
      b?.addEventListener("input",  syncFromFirst);
      e?.addEventListener("input",  syncFromFirst);
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

    // elsőből előtöltés + synced státusz
    const firstBeg = document.querySelector('input[name="beginn1"]')?.value || "";
    const firstEnd = document.querySelector('input[name="ende1"]')?.value || "";
    const begNew = tpl.querySelector(`input[name="beginn${idx}"]`);
    const endNew = tpl.querySelector(`input[name="ende${idx}"]`);
    if (firstBeg && begNew) { begNew.value = firstBeg; markSynced(begNew, true); }
    if (firstEnd && endNew) { endNew.value = firstEnd; markSynced(endNew, true); }

    wireWorker(tpl);
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

  // dolgozói CSV betöltés a végén
  loadWorkers();
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
