document.addEventListener("DOMContentLoaded", () => {
  const addBtn = document.getElementById("add-worker");
  const workerList = document.getElementById("worker-list");
  const totalOut = document.getElementById("gesamtstunden_auto");
  const breakHalf = document.getElementById("break_half");
  const breakHidden = document.getElementById("break_minutes");
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

  // --- CSV parser ---
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
    const idxNach = header.findIndex(h => h === "nachname" || h === "name");
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

  // ===== Egyedi lenyíló autocomplete =====
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
      () => WORKERS.map(w => ({
        value: w.vorname,
        label: `${w.vorname} — ${w.nachname} [${w.ausweis}]`,
        payload: w
      })),
      (value, item) => {
        if (item && item.payload) {
          const w = item.payload;
          inpVor.value  = w.vorname;
          inpNach.value = w.nachname;
          inpAus.value  = w.ausweis;
          return;
        }
        const w2 = byFullName.get(keyName(inpNach.value, value));
        if (w2) inpAus.value = w2.ausweis;
      }
    );

    // Nachname
    makeAutocomplete(
      inpNach,
      () => WORKERS.map(w => ({
        value: w.nachname,
        label: `${w.nachname} — ${w.vorname} [${w.ausweis}]`,
        payload: w
      })),
      (value, item) => {
        if (item && item.payload) {
          const w = item.payload;
          inpNach.value = w.nachname;
          inpVor.value  = w.vorname;
          inpAus.value  = w.ausweis;
          return;
        }
        const w2 = byFullName.get(keyName(value, inpVor.value));
        if (w2) inpAus.value = w2.ausweis;
      }
    );

    // Ausweis
    makeAutocomplete(
      inpAus,
      () => WORKERS.map(w => ({ value: w.ausweis, label: `${w.ausweis} — ${w.nachname} ${w.vorname}` })),
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

  // --- TIME PICKER ---
  function enhanceTimePicker(inp) {
    if (!inp || inp.dataset.enhanced === "1") return;
    inp.dataset.enhanced = "1";
    inp.type = "hidden";

    const box = document.createElement("div");
    box.style.display = "flex";
    box.style.gap = ".5rem";
    box.style.alignItems = "center";

    const selH = document.createElement("select");
    selH.style.minWidth = "5.2rem";
    const optH0 = new Option("--", "");
    selH.add(optH0);
    for (let h = 0; h < 24; h++) {
      const v = String(h).padStart(2, "0");
      selH.add(new Option(v, v));
    }

    const selM = document.createElement("select");
    selM.style.minWidth = "5.2rem";
    const optM0 = new Option("--", "");
    selM.add(optM0);
    ["00","15","30","45"].forEach(m => selM.add(new Option(m, m)));

    inp.insertAdjacentElement("afterend", box);
    box.appendChild(selH);
    box.appendChild(selM);

    function dispatchBoth() {
      inp.dispatchEvent(new Event("input", { bubbles: true }));
      inp.dispatchEvent(new Event("change", { bubbles: true }));
    }

    function composeFromSelects() {
      const h = selH.value;
      const m = selM.value;
      const prev = inp.value;
      inp.value = (h && m) ? `${h}:${m}` : "";
      if (inp.value !== prev) dispatchBoth();
    }

    inp._setFromValue = (v) => {
      const mm = /^\d{2}:\d{2}$/.test(v) ? v.split(":") : ["",""];
      selH.value = mm[0] || "";
      const m = mm[1] || "";
      const allowed = ["00","15","30","45"];
      selM.value = allowed.includes(m) ? m : (m ? String(Math.round(parseInt(m,10)/15)*15).padStart(2,"0") : "");
      composeFromSelects();
    };

    inp._setFromValue(inp.value || "");
    selH.addEventListener("change", composeFromSelects);
    selM.addEventListener("change", composeFromSelects);
  }

  // --- helpers (óra-számítás, stb.) ---
  function toTime(s) {
    if (!s || !/^\d{2}:\d{2}$/.test(s)) return null;
    const [hh, mm] = s.split(":").map(Number);
    if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
    return { hh, mm };
  }
  function minutes(t) { return t.hh * 60 + t.mm; }

  // átfedés percekben
  function overlap(a1, a2, b1, b2) {
    const s = Math.max(a1, b1);
    const e = Math.min(a2, b2);
    return Math.max(0, e - s);
    }

  function hoursWithBreaks(begStr, endStr) {
    const bt = toTime(begStr);
    const et = toTime(endStr);
    if (!bt || !et) return 0;
    const b = minutes(bt), e = minutes(et);
    if (e <= b) return 0;

    let total = e - b;

    // „fél óra szünet” pipa → fix 30 perc levonás
    const bm = parseInt((breakHidden?.value || "60"), 10);
    const isHalfHour = !isNaN(bm) && bm === 30;

    if (isHalfHour) {
      total = Math.max(0, total - 30);
    } else {
      // Alapértelmezés: CSAK a tényleges átfedés a két fix szüneti sávval
      const br1s = 9 * 60 + 0,  br1e = 9 * 60 + 15;  // 09:00–09:15
      const br2s = 12 * 60 + 0, br2e = 12 * 60 + 45; // 12:00–12:45
      const minus = overlap(b, e, br1s, br1e) + overlap(b, e, br2s, br2e);
      total = Math.max(0, total - minus);
    }

    return Math.max(0, total) / 60;
  }

  function formatHours(h) { return (Math.round(h * 100) / 100).toFixed(2); }
  function digitsOnly(input) { input.addEventListener("input", () => { input.value = input.value.replace(/\D/g, ""); }); }
  function enforceNumericKeyboard(input) {
    input.setAttribute("inputmode", "numeric");
    input.setAttribute("pattern", "[0-9]*");
    digitsOnly(input);
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

  // Globális szünet checkbox kezelése
  function updateBreakAndRecalc() {
    if (!breakHidden) return;
    breakHidden.value = breakHalf?.checked ? "30" : "60";
    recalcAll();
  }
  breakHalf?.addEventListener("change", updateBreakAndRecalc);
  updateBreakAndRecalc();

  // ---- SYNC: első dolgozó ideje többiekhez ----
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
      if (beg && (isSynced(beg) || !beg.value)) {
        beg.value = firstBeg;
        if (beg._setFromValue) beg._setFromValue(firstBeg);
        markSynced(beg, true);
      }
      if (end && (isSynced(end) || !end.value)) {
        end.value = firstEnd;
        if (end._setFromValue) end._setFromValue(firstEnd);
        markSynced(end, true);
      }
    }
    recalcAll();
  }

  function wireWorker(workerEl) {
    const idx = workerEl.getAttribute("data-index");

    // autocomplete
    setupAutocomplete(workerEl);

    // Ausweis numerikus
    const ausweis = workerEl.querySelector(`input[name="ausweis${idx}"]`);
    if (ausweis) enforceNumericKeyboard(ausweis);

    // idő mezők
    ["beginn", "ende"].forEach(prefix => {
      const inp = workerEl.querySelector(`input[name^="${prefix}"]`);
      if (inp) {
        enhanceTimePicker(inp);
        inp.addEventListener("change", recalcAll);
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

  // új dolgozó (Vorhaltung mezővel)
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
      <div class="grid">
        <div class="field">
          <label>Vorhaltung / beauftragtes Gerät / Fahrzeug</label>
          <input name="vorhaltung${idx}" type="text" />
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

    const firstBeg = document.querySelector('input[name="beginn1"]')?.value || "";
    const firstEnd = document.querySelector('input[name="ende1"]')?.value || "";
    const begNew = tpl.querySelector(`input[name="beginn${idx}"]`);
    const endNew = tpl.querySelector(`input[name="ende${idx}"]`);
    if (firstBeg && begNew) { begNew.value = firstBeg; }
    if (firstEnd && endNew) { endNew.value = firstEnd; }

    wireWorker(tpl);
    if (begNew && begNew._setFromValue) begNew._setFromValue(begNew.value || "");
    if (endNew && endNew._setFromValue) endNew._setFromValue(endNew.value || "");

    if (begNew && begNew.value) markSynced(begNew, true);
    if (endNew && endNew.value) markSynced(endNew, true);

    recalcAll();
  });

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

// --- Vezérelt letöltés (fetch + blob) — a kattintott gomb szerint ---
(function () {
  const form = document.getElementById("ln-form");
  if (!form) return;

  function setBusy(btn, busy) {
    if (!btn) return;
    if (busy) {
      btn.dataset._label = btn.textContent || btn.value || "Generieren";
      btn.disabled = true;
      btn.textContent = "Wird generiert...";
    } else {
      btn.disabled = false;
      const lbl = btn.dataset._label;
      if (lbl) btn.textContent = lbl;
    }
  }

  function filenameFromDisposition(h) {
    if (!h) return null;
    const m = /filename\*?=(?:UTF-8''|")?([^\";]+)"/i.exec(h) || /filename=([^;]+)/i.exec(h);
    if (!m) return null;
    try { return decodeURIComponent(m[1].replace(/"/g, "").trim()); }
    catch { return m[1].replace(/"/g, "").trim(); }
  }

  async function fetchAndHandle(fd, url, openInNewTab, fallbackName) {
    const res = await fetch(url, { method: "POST", body: fd });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const blob = await res.blob();

    const dispo = res.headers.get("content-disposition") || "";
    const name  = filenameFromDisposition(dispo) || fallbackName;

    const ct = (res.headers.get("content-type") || "").toLowerCase();
    const objURL = URL.createObjectURL(blob);

    // ha PDF és új fülre kérjük → megnyitás új fülön
    if (openInNewTab && ct.includes("pdf")) {
      window.open(objURL, "_blank", "noopener");
      setTimeout(() => URL.revokeObjectURL(objURL), 4000);
      return;
    }

    // különben letöltés
    const a = document.createElement("a");
    a.href = objURL;
    a.download = name || "download";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(objURL), 2000);
  }

  form.addEventListener("submit", async (e) => {
    if (e.defaultPrevented) return;
    e.preventDefault();

    // melyik gomb indította?
    const btn = e.submitter || form.querySelector('button[type="submit"]');
    const endpoint =
      (btn && btn.getAttribute("formaction")) ||
      form.getAttribute("action") ||
      "/generate_excel";

    const openInNewTab = (btn && btn.getAttribute("formtarget") === "_blank");
    const isPdf = endpoint.endsWith("/generate_pdf");

    const fallbackName = isPdf ? "leistungsnachweis_preview.pdf" : "leistungsnachweis.xlsx";

    const fd = new FormData(form);
    setBusy(btn, true);
    try {
      await fetchAndHandle(fd, endpoint, openInNewTab, fallbackName);
    } catch (err) {
      console.error(err);
      alert("A fájl generálása most nem sikerült. Kérlek próbáld meg újra.");
    } finally {
      setBusy(btn, false);
    }
  }, false);
})();
