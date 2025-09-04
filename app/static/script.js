/* static/script.js – consolidated

Funkciók:
- HR→DE fordítás gomb (/api/translate)
- Karakterszámláló a #beschreibung alatt (JS injektál, ha hiányzik)
- /api/workers JSON betöltése (CSV fallback), autocomplete
- 15 perces időválasztó (óra+perc) rejtett <input type="time"> fölé
- Óraszámítás soronként és összesen, félórás szünet opcióval (30/60)
- 1. sor időinek átvétele többi sorra, amíg kézzel nem írják felül
- Nyelvválasztó (ha nincs: JS injektálja), ?lang=de|hr
*/

(function () {
  // ---------- DOM helpers ----------
  const $id = (id) => document.getElementById(id);
  const norm = (s) => (s ?? "").toString().trim();
  const onReady = (fn) =>
    document.readyState === "loading"
      ? document.addEventListener("DOMContentLoaded", fn)
      : fn();

  // ---------- Stabil referenciák több HTML-variánshoz ----------
  function pickForm() {
    return (
      $id("ln-form") || $id("rapportForm") || $id("frm") || document.querySelector("form")
    );
  }
  function pickWorkersBox() {
    return $id("worker-list") || $id("workers") || document.querySelector("#workers, #worker-list");
  }
  function pickAddBtn() {
    return $id("add-worker") || $id("addWorker") || document.querySelector("#add-worker, #addWorker");
  }

  // ---------- Nyelv (egyszerű i18n) ----------
  const Q = new URLSearchParams(location.search);
  const LANG = (Q.get("lang") || window.APP_LANG || "de").toLowerCase();

  const T = {
    de: {
      workerLegend: (i) => `Mitarbeiter ${i}`,
      vorname: "Vorname",
      nachname: "Nachname",
      ausweis: "Ausweis-Nr. / Kennzeichen",
      vorhaltung: "Vorhaltung / beauftragtes Gerät / Fahrzeug",
      beginn: "Beginn",
      ende: "Ende",
      stunden: "Stunden (auto)",
      addMore: "+ Weiterer Mitarbeiter",
      total: "Gesamtstunden (berechnet)",
      translate: "Aus kroatisch übersetzen",
      translating: "Übersetzen…",
      done: "Fertig ✓",
      noText: "Kein Text zum Übersetzen.",
      fail: "Übersetzung fehlgeschlagen.",
      neterr: "Netzwerkfehler bei der Übersetzung.",
      break30: "Halbe Stunde Pause",
    },
    hr: {
      workerLegend: (i) => `Radnik ${i}`,
      vorname: "Ime",
      nachname: "Prezime",
      ausweis: "Broj iskaznice / oznaka",
      vorhaltung: "Najam / angažirana oprema / vozilo",
      beginn: "Početak",
      ende: "Kraj",
      stunden: "Sati (auto)",
      addMore: "+ Dodaj radnika",
      total: "Ukupno sati (izračun)",
      translate: "Prevedi s hrvatskog",
      translating: "Prevodi…",
      done: "Gotovo ✓",
      noText: "Nema teksta za prijevod.",
      fail: "Prevod neuspješan.",
      neterr: "Mrežna greška pri prijevodu.",
      break30: "Polusatna pauza",
    },
  }[LANG] || T?.de;

  // ---------- Karakterszámláló + Fordítás gomb ----------
  function attachBeschFeatures() {
    const ta = $id("beschreibung");
    if (!ta) return;

    // számláló injektálás, ha hiányzik
    let cnt = $id("besch-count");
    if (!cnt) {
      cnt = document.createElement("div");
      cnt.id = "besch-count";
      cnt.className = "small muted";
      ta.insertAdjacentElement("afterend", cnt);
    }
    const upd = () => (cnt.textContent = `${(ta.value || "").length} / 1000`);
    ta.addEventListener("input", upd);
    upd();

    // fordítás gomb – ha van erre dedikált gomb, azt használjuk; ha nincs, injektálunk egyet
    let trBtn = document.querySelector("#translate-btn");
    if (!trBtn) {
      trBtn = document.createElement("button");
      trBtn.type = "button";
      trBtn.id = "translate-btn";
      trBtn.className = "chip";
      trBtn.textContent = T.translate;
      cnt.insertAdjacentElement("afterend", trBtn);
    }
    let trStatus = document.querySelector("#translate-status");
    if (!trStatus) {
      trStatus = document.createElement("span");
      trStatus.id = "translate-status";
      trStatus.className = "small muted";
      trBtn.insertAdjacentElement("afterend", trStatus);
    }

    trBtn.addEventListener("click", async () => {
      const txt = (ta.value || "").trim();
      if (!txt) {
        trStatus.textContent = T.noText;
        return;
      }
      trBtn.disabled = true;
      trStatus.textContent = T.translating;
      try {
        const res = await fetch("/api/translate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text: txt, source: "hr", target: "de" }),
        });
        const data = await res.json().catch(() => ({}));
        if (res.ok && data && typeof data.translated === "string") {
          ta.value = data.translated;
          ta.dispatchEvent(new Event("input", { bubbles: true }));
          trStatus.textContent = T.done;
        } else trStatus.textContent = T.fail;
      } catch {
        trStatus.textContent = T.neterr;
      } finally {
        trBtn.disabled = false;
        setTimeout(() => (trStatus.textContent = ""), 4000);
      }
    });
  }

  // ---------- Nyelvválasztó injektálása (ha hiányzik) ----------
  function ensureLangPicker(form) {
    let sel = $id("lang-select");
    if (sel) return;
    sel = document.createElement("select");
    sel.id = "lang-select";
    sel.innerHTML = `
      <option value="de"${LANG === "de" ? " selected" : ""}>Deutsch</option>
      <option value="hr"${LANG === "hr" ? " selected" : ""}>Hrvatski</option>
    `;
    // form tetejére tesszük
    form.insertAdjacentElement("afterbegin", sel);
    sel.addEventListener("change", () => {
      const url = new URL(location.href);
      url.searchParams.set("lang", sel.value);
      location.href = url.toString();
    });
  }

  // ---------- Workers forrás (JSON preferált, CSV fallback) ----------
  let WORKERS = []; // {vorname, nachname, ausweis}
  const byAusweis = new Map();
  const byFullName = new Map();
  const keyName = (nach, vor) =>
    `${norm(nach).toLowerCase()}|${norm(vor).toLowerCase()}`;

  async function loadWorkers() {
    WORKERS = [];
    byAusweis.clear();
    byFullName.clear();

    // 1) JSON
    try {
      const r = await fetch("/api/workers", { cache: "no-store" });
      if (r.ok && (r.headers.get("content-type") || "").includes("json")) {
        const arr = await r.json();
        (arr || []).forEach((w) => {
          const obj = {
            vorname: norm(w.first_name || w.vorname),
            nachname: norm(w.last_name || w.nachname),
            ausweis: norm(w.badge || w.ausweis),
          };
          if (!obj.vorname && !obj.nachname && !obj.ausweis) return;
          WORKERS.push(obj);
          if (obj.ausweis) byAusweis.set(obj.ausweis, obj);
          byFullName.set(keyName(obj.nachname, obj.vorname), obj);
        });
        return;
      }
    } catch {/* pass */}

    // 2) CSV fallback
    try {
      const resp = await fetch("/api/workers.csv", { cache: "no-store" });
      if (!resp.ok) return;
      let text = await resp.text();
      if (text && text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
      const lines = text.split(/\r?\n/).filter((l) => l.trim().length);
      if (!lines.length) return;

      const delim = (lines[0].match(/;/g) || []).length >= (lines[0].match(/,/g) || []).length ? ";" : ",";
      const splitCSV = (line) =>
        line
          .split(new RegExp(`${delim}(?=(?:[^"]*"[^"]*")*[^"]*$)`, "g"))
          .map((x) => {
            x = x.trim();
            if (x.startsWith('"') && x.endsWith('"')) x = x.slice(1, -1);
            return x.replace(/""/g, '"');
          });

      const header = splitCSV(lines[0]).map((h) => h.toLowerCase().trim());
      const idxNach = header.findIndex((h) => h === "nachname" || h === "name");
      const idxVor = header.findIndex((h) => h === "vorname");
      const idxAus = header.findIndex((h) => /(ausweis|kennzeichen)/.test(h));
      if (idxNach < 0 || idxVor < 0 || idxAus < 0) return;

      for (let i = 1; i < lines.length; i++) {
        const cols = splitCSV(lines[i]);
        const obj = {
          nachname: norm(cols[idxNach]),
          vorname: norm(cols[idxVor]),
          ausweis: norm(cols[idxAus]),
        };
        if (!obj.vorname && !obj.nachname && !obj.ausweis) continue;
        WORKERS.push(obj);
        if (obj.ausweis) byAusweis.set(obj.ausweis, obj);
        byFullName.set(keyName(obj.nachname, obj.vorname), obj);
      }
    } catch {/* pass */}
  }

  // ---------- Autocomplete ----------
  function makeAutocomplete(input, listProvider, onPick) {
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
    document.body.appendChild(dd);

    function hide() { dd.style.display = "none"; }
    function show() { dd.style.display = dd.children.length ? "block" : "none"; }
    function position() {
      const r = input.getBoundingClientRect();
      dd.style.left = `${window.scrollX + r.left}px`;
      dd.style.top = `${window.scrollY + r.bottom}px`;
      dd.style.width = `${r.width}px`;
    }

    function render(list) {
      dd.innerHTML = "";
      (list || []).slice(0, 12).forEach((item) => {
        const opt = document.createElement("div");
        opt.textContent = item.label ?? item.value ?? "";
        opt.dataset.value = item.value ?? item.label ?? "";
        opt.style.padding = ".5rem .75rem";
        opt.style.cursor = "pointer";
        opt.addEventListener("mousedown", (e) => {
          e.preventDefault();
          input.value = opt.dataset.value;
          hide();
          onPick?.(opt.dataset.value, item);
          input.dispatchEvent(new Event("change", { bubbles: true }));
          input.dispatchEvent(new Event("input", { bubbles: true }));
        });
        opt.addEventListener("mouseover", () => (opt.style.background = "#f5f5f5"));
        opt.addEventListener("mouseout", () => (opt.style.background = "white"));
        dd.appendChild(opt);
      });
      position(); show();
    }

    function filterOptions() {
      const q = (input.value || "").toLowerCase().trim();
      if (!q) return hide();
      const base = listProvider() || [];
      const list = base
        .filter((v) => (v.label ?? v.value ?? v).toLowerCase().includes(q))
        .map((v) => (typeof v === "string" ? { value: v } : v));
      render(list);
    }

    input.addEventListener("input", filterOptions);
    input.addEventListener("focus", () => { position(); filterOptions(); });
    input.addEventListener("blur", () => setTimeout(hide, 120));
    window.addEventListener("scroll", position, true);
    window.addEventListener("resize", position);
  }

  function setupAutocomplete(fs) {
    const idx = fs.getAttribute("data-index");
    const inpNach = fs.querySelector(`input[name="nachname${idx}"]`);
    const inpVor  = fs.querySelector(`input[name="vorname${idx}"]`);
    const inpAus  = fs.querySelector(`input[name="ausweis${idx}"]`);
    if (!inpNach || !inpVor || !inpAus) return;

    makeAutocomplete(
      inpVor,
      () => WORKERS.map((w) => ({
        value: w.vorname,
        label: `${w.vorname} — ${w.nachname} [${w.ausweis}]`,
        payload: w,
      })),
      (_v, item) => {
        const w = item?.payload;
        if (w) { inpVor.value = w.vorname; inpNach.value = w.nachname; inpAus.value = w.ausweis; return; }
        const w2 = byFullName.get(keyName(inpNach.value, _v));
        if (w2) inpAus.value = w2.ausweis;
      }
    );

    makeAutocomplete(
      inpNach,
      () => WORKERS.map((w) => ({
        value: w.nachname,
        label: `${w.nachname} — ${w.vorname} [${w.ausweis}]`,
        payload: w,
      })),
      (_v, item) => {
        const w = item?.payload;
        if (w) { inpNach.value = w.nachname; inpVor.value = w.vorname; inpAus.value = w.ausweis; return; }
        const w2 = byFullName.get(keyName(_v, inpVor.value));
        if (w2) inpAus.value = w2.ausweis;
      }
    );

    makeAutocomplete(
      inpAus,
      () => WORKERS.map((w) => ({
        value: w.ausweis,
        label: `${w.ausweis} — ${w.nachname} ${w.vorname}`,
      })),
      (val) => {
        const w = byAusweis.get(norm(val));
        if (w) { inpNach.value = w.nachname; inpVor.value = w.vorname; }
      }
    );

    // keresztkitöltés
    inpAus.addEventListener("change", () => {
      const w = byAusweis.get(norm(inpAus.value));
      if (w) { inpNach.value = w.nachname; inpVor.value = w.vorname; }
    });
    const tryNames = () => {
      const w = byFullName.get(keyName(inpNach.value, inpVor.value));
      if (w) inpAus.value = w.ausweis;
    };
    inpNach.addEventListener("change", tryNames);
    inpVor.addEventListener("change", tryNames);
  }

  // ---------- Időválasztó + óraszámítás ----------
  function toTime(s) {
    if (!s || !/^\d{2}:\d{2}$/.test(s)) return null;
    const [hh, mm] = s.split(":").map(Number);
    if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
    return { hh, mm };
  }
  const minutes = (t) => t.hh * 60 + t.mm;
  const overlap = (a1, a2, b1, b2) => Math.max(0, Math.min(a2, b2) - Math.max(a1, b1));

  function hoursWithBreaks(begStr, endStr, breakMin) {
    const bt = toTime(begStr); const et = toTime(endStr);
    if (!bt || !et) return 0;
    const b = minutes(bt), e = minutes(et);
    if (e <= b) return 0;
    let total = e - b;
    if (breakMin === 30) {
      total = Math.max(0, total - 30);
    } else {
      const minus = overlap(b,e,9*60,9*60+15) + overlap(b,e,12*60,12*60+45);
      total = Math.max(0, total - minus);
    }
    return Math.max(0, total) / 60;
  }
  const fmtH = (h) => (Math.round(h * 100) / 100).toFixed(2);

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
    selH.add(new Option("--", ""));
    for (let h = 0; h < 24; h++) selH.add(new Option(String(h).padStart(2, "0"), String(h).padStart(2, "0")));

    const selM = document.createElement("select");
    selM.style.minWidth = "5.2rem";
    selM.add(new Option("--", ""));
    ["00", "15", "30", "45"].forEach((m) => selM.add(new Option(m, m)));

    inp.insertAdjacentElement("afterend", box);
    box.appendChild(selH); box.appendChild(selM);

    function dispatchBoth() {
      inp.dispatchEvent(new Event("input", { bubbles: true }));
      inp.dispatchEvent(new Event("change", { bubbles: true }));
      if (typeof recalcAll === "function") try { recalcAll(); } catch {}
    }
    function compose() {
      const h = selH.value, m = selM.value;
      const prev = inp.value;
      inp.value = h && m ? `${h}:${m}` : "";
      if (inp.value !== prev) dispatchBoth();
    }
    inp._setFromValue = (v) => {
      const mm = /^\d{2}:\d{2}$/.test(v) ? v.split(":") : ["", ""];
      selH.value = mm[0] || "";
      const m = mm[1] || "";
      const allowed = ["00","15","30","45"];
      selM.value = allowed.includes(m) ? m : (m ? String(Math.round(parseInt(m, 10)/15)*15).padStart(2,"0") : "");
      compose();
    };

    inp._setFromValue(inp.value || "");
    selH.addEventListener("change", compose);
    selM.addEventListener("change", compose);
  }

  // Szinkron az első sorból – manuális edit bontja a szinkront
  function markSynced(inp, yes) { if (!inp) return; if (yes) inp.dataset.synced = "1"; else delete inp.dataset.synced; }
  function isSynced(inp) { return !!(inp && inp.dataset.synced === "1"); }
  function setupManualUnsync(inp) {
    if (!inp) return;
    const uns = () => markSynced(inp, false);
    inp.addEventListener("input", (e) => { if (e.isTrusted) uns(); });
    inp.addEventListener("change", (e) => { if (e.isTrusted) uns(); });
  }

  // Félórás szünet – ha nincs a HTML-ben, injektálunk
  function ensureBreakToggle(form) {
    let chk = $id("break_half");
    let hidden = $id("break_minutes");
    if (!chk) {
      const wrap = document.createElement("label");
      wrap.className = "chip";
      chk = document.createElement("input");
      chk.type = "checkbox";
      chk.id = "break_half";
      wrap.appendChild(chk);
      wrap.appendChild(document.createTextNode(" " + T.break30));
      form.insertAdjacentElement("afterbegin", wrap);
    }
    if (!hidden) {
      hidden = document.createElement("input");
      hidden.type = "hidden";
      hidden.id = "break_minutes";
      hidden.name = "break_minutes";
      form.appendChild(hidden);
    }
    const sync = () => { hidden.value = chk.checked ? "30" : "60"; recalcAll(); };
    chk.addEventListener("change", sync); sync();
    return { chk, hidden };
  }

  // --------- Sor felépítése ---------
  function buildWorkerFieldset(i) {
    const fs = document.createElement("fieldset");
    fs.className = "worker";
    fs.setAttribute("data-index", String(i));
    fs.innerHTML = `
      <legend>${T.workerLegend(i)}</legend>
      <div class="grid-3">
        <div class="field"><label>${T.vorname}</label><input name="vorname${i}" type="text" /></div>
        <div class="field"><label>${T.nachname}</label><input name="nachname${i}" type="text" /></div>
        <div class="field"><label>${T.ausweis}</label><input name="ausweis${i}" type="text" /></div>
      </div>
      <div class="grid"><div class="field"><label>${T.vorhaltung}</label><input name="vorhaltung${i}" type="text" /></div></div>
      <div class="grid-3">
        <div class="field"><label>${T.beginn}</label><input name="beginn${i}" type="time" /></div>
        <div class="field"><label>${T.ende}</label><input name="ende${i}" type="time" /></div>
        <div class="field"><label>${T.stunden}</label><input class="stunden-display" type="text" value="" readonly /></div>
      </div>
    `;
    return fs;
  }

  // --------- Kalkulációk ---------
  let workerList, form, breakHidden;
  function recalcWorker(workerEl) {
    const beg = workerEl.querySelector('input[name^="beginn"]')?.value || "";
    const end = workerEl.querySelector('input[name^="ende"]')?.value || "";
    const out = workerEl.querySelector(".stunden-display");
    const h = hoursWithBreaks(beg, end, parseInt(breakHidden?.value || "60", 10));
    if (out) out.value = h ? fmtH(h) : "";
    return h;
  }
  function recalcAll() {
    if (!workerList) return;
    let sum = 0;
    workerList.querySelectorAll(".worker").forEach((w) => (sum += recalcWorker(w)));
    const totalOut = $id("gesamtstunden_auto") || $id("total_hours");
    if (totalOut) totalOut.value = sum ? fmtH(sum) : "";
  }

  function syncFromFirst() {
    const first = workerList.querySelector('.worker[data-index="1"]');
    if (!first) return;
    const firstBeg = first.querySelector('input[name="beginn1"]')?.value || "";
    const firstEnd = first.querySelector('input[name="ende1"]')?.value || "";
    if (!firstBeg && !firstEnd) return;
    const workers = Array.from(workerList.querySelectorAll(".worker"));
    for (let i = 1; i < workers.length; i++) {
      const fs = workers[i];
      const beg = fs.querySelector('input[name^="beginn"]');
      const end = fs.querySelector('input[name^="ende"]');
      if (beg && (isSynced(beg) || !beg.value)) { beg.value = firstBeg; beg._setFromValue?.(firstBeg); markSynced(beg, true); }
      if (end && (isSynced(end) || !end.value)) { end.value = firstEnd; end._setFromValue?.(firstEnd); markSynced(end, true); }
    }
    recalcAll();
  }

  // --------- Indítás ---------
  onReady(async function init() {
    form = pickForm();
    workerList = pickWorkersBox();
    const addBtn = pickAddBtn();
    if (!form || !workerList || !addBtn) {
      console.warn("Form/workerList/addBtn hiányzik – ellenőrizd a markupot.");
      return;
    }

    ensureLangPicker(form);
    const { hidden } = ensureBreakToggle(form);
    breakHidden = hidden;

    attachBeschFeatures();
    await loadWorkers();

    // legalább 1 sor
    if (!workerList.querySelector(".worker")) {
      const fs1 = buildWorkerFieldset(1);
      workerList.appendChild(fs1);
    }

    // meglévők bekötése
    workerList.querySelectorAll(".worker").forEach((fs) => {
      fs.querySelectorAll('input[type="time"]').forEach((inp) => {
        enhanceTimePicker(inp);
        setupManualUnsync(inp);
        inp.addEventListener("input", recalcAll);
        inp.addEventListener("change", recalcAll);
      });
      setupAutocomplete(fs);
    });

    // + gomb
    const MAX_WORKERS = 5;
    addBtn.addEventListener("click", () => {
      const count = workerList.querySelectorAll(".worker").length;
      if (count >= MAX_WORKERS) return;
      const next = count + 1;
      const fs = buildWorkerFieldset(next);
      workerList.appendChild(fs);
      fs.querySelectorAll('input[type="time"]').forEach((inp) => {
        enhanceTimePicker(inp);
        setupManualUnsync(inp);
        inp.addEventListener("input", recalcAll);
        inp.addEventListener("change", recalcAll);
      });
      setupAutocomplete(fs);
      syncFromFirst();
      recalcAll();
    });

    // ha az elsőben változtatunk, szinkronizáljuk lefelé
    const firstBeg = form.querySelector('input[name="beginn1"]');
    const firstEnd = form.querySelector('input[name="ende1"]');
    [firstBeg, firstEnd].forEach((inp) => {
      if (!inp) return;
      inp.addEventListener("change", syncFromFirst);
      inp.addEventListener("input", syncFromFirst);
    });

    recalcAll();
  });
})();
