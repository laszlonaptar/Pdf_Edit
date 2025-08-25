/* static/script.js
   - Autocomplete dolgozók CSV-ből (Vorname/Nachname/Ausweis szinkron kitöltés)
   - Mobilbarát 15 perces időválasztó (00/15/30/45)
   - Óraszámítás (fix sávos szünet: 09:00–09:15 és 12:00–12:45, vagy 0,5 h)
   - 1. dolgozó idejének szinkron másolása az új dolgozókra (amíg kézzel felül nem írod)
   - Validáció elküldés előtt
   - Excel letöltés: fetch + blob
   - (Opcionális) PDF Vorschau: új lap + IFRAME
   - i18n: nyelvváltás (de/hr) a window.I18N szótárral, #lang-select alapján
*/

document.addEventListener("DOMContentLoaded", () => {
  /* ===================== I18N ===================== */
  const I18N = (window.I18N || {});
  const DEFAULT_LANG = "de";

  function getLang() {
    // 1) URL ?lang=...  2) localStorage  3) <html lang>  4) default
    const u = new URL(window.location.href);
    const ql = (u.searchParams.get("lang") || "").trim();
    const ls = (localStorage.getItem("app_lang") || "").trim();
    const hl = (document.documentElement.getAttribute("lang") || "").trim();
    return (ql || ls || hl || DEFAULT_LANG);
  }
  function t(key, lang) {
    const L = lang || currentLang;
    try {
      if (I18N[L] && key in I18N[L]) return I18N[L][key];
      if (I18N[DEFAULT_LANG] && key in I18N[DEFAULT_LANG]) return I18N[DEFAULT_LANG][key];
    } catch (_) {}
    return null;
  }
  function translateNode(el, lang) {
    if (!el) return;
    const key = el.getAttribute("data-i18n");
    if (key) {
      const txt = t(key, lang);
      if (txt != null) el.textContent = txt;
    }
    const phKey = el.getAttribute && el.getAttribute("data-i18n-placeholder");
    if (phKey && el.placeholder !== undefined) {
      const ph = t(phKey, lang);
      if (ph != null) el.placeholder = ph;
    }
    const valKey = el.getAttribute && el.getAttribute("data-i18n-value");
    if (valKey && el.value !== undefined) {
      const v = t(valKey, lang);
      if (v != null) el.value = v;
    }
  }
  function applyTranslations(lang) {
    // oldal cím
    const ttl = t("title", lang);
    if (ttl) document.title = ttl;

    // minden jelölt elem
    document.querySelectorAll("[data-i18n], [data-i18n-placeholder], [data-i18n-value]").forEach(el => {
      translateNode(el, lang);
    });

    // „Wird generiert...” szöveg fordítása, ha van
    const gen = t("generating", lang);
    if (gen) window.__GEN_TEXT = gen; // a setBusy használja, ha be van állítva
  }

  let currentLang = getLang();
  // szinkronizáljuk a <select id="lang-select"> értékét (ha van)
  const langSel = document.getElementById("lang-select");
  if (langSel) {
    if ([...langSel.options].some(o => o.value === currentLang)) {
      langSel.value = currentLang;
    }
    langSel.addEventListener("change", () => {
      currentLang = langSel.value || DEFAULT_LANG;
      localStorage.setItem("app_lang", currentLang);
      // <html lang> frissítése
      document.documentElement.setAttribute("lang", currentLang);
      applyTranslations(currentLang);
    });
  }
  // kezdeti beállítás
  document.documentElement.setAttribute("lang", currentLang);
  applyTranslations(currentLang);

  /* ============== A LAP TÖBBI FUNKCIÓJA ============== */

  const addBtn = document.getElementById("add-worker");
  const workerList = document.getElementById("worker-list");
  const totalOut = document.getElementById("gesamtstunden_auto");
  const breakHalf = document.getElementById("break_half");
  const breakHidden = document.getElementById("break_minutes");
  const form = document.getElementById("ln-form");
  const MAX_WORKERS = 5;

  // ===== AUTOCOMPLETE =====
  let WORKERS = [];
  const byAusweis = new Map();
  const byFullName = new Map(); // "nachname|vorname" -> worker

  const norm = (s) => (s || "").toString().trim();
  const keyName = (nach, vor) =>
    `${norm(nach).toLowerCase()}|${norm(vor).toLowerCase()}`;

  async function loadWorkers() {
    try {
      const resp = await fetch("/static/workers.csv", { cache: "no-store" });
      if (!resp.ok) return;
      let text = await resp.text();
      parseCSV(text);
    } catch (_) {}
  }

  function parseCSV(text) {
    if (text && text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const lines = text.split(/\r?\n/).filter((l) => l.trim().length);
    if (!lines.length) return;

    const detectDelim = (s) => {
      const sc = (s.match(/;/g) || []).length;
      const cc = (s.match(/,/g) || []).length;
      return sc > cc ? ";" : ",";
    };
    const delim = detectDelim(lines[0]);

    const splitCSV = (line) => {
      const re = new RegExp(`${delim}(?=(?:[^"]*"[^"]*")*[^"]*$)`, "g");
      return line.split(re).map((x) => {
        x = x.trim();
        if (x.startsWith('"') && x.endsWith('"')) x = x.slice(1, -1);
        return x.replace(/""/g, '"');
      });
    };

    const header = splitCSV(lines[0]).map((h) => h.toLowerCase().trim());
    const idxNach = header.findIndex((h) => h === "nachname" || h === "name");
    const idxVor = header.findIndex((h) => h === "vorname");
    const idxAus = header.findIndex((h) => /(ausweis|kennzeichen)/.test(h));
    if (idxNach < 0 || idxVor < 0 || idxAus < 0) return;

    WORKERS = [];
    byAusweis.clear();
    byFullName.clear();

    for (let i = 1; i < lines.length; i++) {
      const cols = splitCSV(lines[i]);
      if (cols.length <= Math.max(idxNach, idxVor, idxAus)) continue;
      const w = {
        nachname: norm(cols[idxNach]),
        vorname: norm(cols[idxVor]),
        ausweis: norm(cols[idxAus]),
      };
      if (!w.nachname && !w.vorname && !w.ausweis) continue;
      WORKERS.push(w);
      if (w.ausweis) byAusweis.set(w.ausweis, w);
      byFullName.set(keyName(w.nachname, w.vorname), w);
    }

    refreshAllAutocompletes();
  }

  // ===== Egyedi lenyíló =====
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

    function hide() {
      dd.style.display = "none";
    }
    function show() {
      dd.style.display = dd.children.length ? "block" : "none";
    }
    function position() {
      const r = input.getBoundingClientRect();
      dd.style.left = `${window.scrollX + r.left}px`;
      dd.style.top = `${window.scrollY + r.bottom}px`;
      dd.style.width = `${r.width}px`;
    }

    function render(list) {
      dd.innerHTML = "";
      list.slice(0, 12).forEach((item) => {
        const opt = document.createElement("div");
        opt.textContent = item.label ?? item.value;
        opt.dataset.value = item.value ?? item.label ?? "";
        opt.style.padding = ".5rem .75rem";
        opt.style.cursor = "pointer";
        opt.addEventListener("mousedown", (e) => {
          e.preventDefault();
          input.value = opt.dataset.value;
          hide();
          onPick?.(opt.dataset.value, item);
          input.dispatchEvent(new Event("change", { bubbles: true }));
        });
        opt.addEventListener("mouseover", () => (opt.style.background = "#f5f5f5"));
        opt.addEventListener("mouseout", () => (opt.style.background = "white"));
        dd.appendChild(opt);
      });
      position();
      show();
    }

    function filterOptions() {
      const q = input.value.toLowerCase().trim();
      const raw = getOptions();
      if (!q) {
        hide();
        return;
      }
      const list = raw
        .filter((v) => (v.label ?? v).toLowerCase().includes(q))
        .map((v) => (typeof v === "string" ? { value: v } : v));
      render(list);
    }

    input.addEventListener("input", filterOptions);
    input.addEventListener("focus", () => {
      position();
      filterOptions();
    });
    input.addEventListener("blur", () => setTimeout(hide, 120));
    window.addEventListener("scroll", position, true);
    window.addEventListener("resize", position);
  }

  function refreshAllAutocompletes() {
    document
      .querySelectorAll("#worker-list .worker")
      .forEach(setupAutocomplete);
  }

  function setupAutocomplete(fs) {
    const idx = fs.getAttribute("data-index");
    const inpNach = fs.querySelector(`input[name="nachname${idx}"]`);
    const inpVor = fs.querySelector(`input[name="vorname${idx}"]`);
    const inpAus = fs.querySelector(`input[name="ausweis${idx}"]`);
    if (!inpNach || !inpVor || !inpAus) return;

    // Vorname
    makeAutocomplete(
      inpVor,
      () =>
        WORKERS.map((w) => ({
          value: w.vorname,
          label: `${w.vorname} — ${w.nachname} [${w.ausweis}]`,
          payload: w,
        })),
      (_value, item) => {
        if (item && item.payload) {
          const w = item.payload;
          inpVor.value = w.vorname;
          inpNach.value = w.nachname;
          inpAus.value = w.ausweis;
          return;
        }
        const w2 = byFullName.get(keyName(inpNach.value, _value));
        if (w2) inpAus.value = w2.ausweis;
      }
    );

    // Nachname
    makeAutocomplete(
      inpNach,
      () =>
        WORKERS.map((w) => ({
          value: w.nachname,
          label: `${w.nachname} — ${w.vorname} [${w.ausweis}]`,
          payload: w,
        })),
      (_value, item) => {
        if (item && item.payload) {
          const w = item.payload;
          inpNach.value = w.nachname;
          inpVor.value = w.vorname;
          inpAus.value = w.ausweis;
          return;
        }
        const w2 = byFullName.get(keyName(_value, inpVor.value));
        if (w2) inpAus.value = w2.ausweis;
      }
    );

    // Ausweis
    makeAutocomplete(
      inpAus,
      () =>
        WORKERS.map((w) => ({
          value: w.ausweis,
          label: `${w.ausweis} — ${w.nachname} ${w.vorname}`,
        })),
      (value) => {
        const w = byAusweis.get(value);
        if (w) {
          inpNach.value = w.nachname;
          inpVor.value = w.vorname;
        }
      }
    );

    // kézi módosítás: próbáljuk következtetni a harmadik mezőt
    inpAus.addEventListener("change", () => {
      const w = byAusweis.get(norm(inpAus.value));
      if (w) {
        inpNach.value = w.nachname;
        inpVor.value = w.vorname;
      }
    });
    const tryNames = () => {
      const w = byFullName.get(keyName(inpNach.value, inpVor.value));
      if (w) inpAus.value = w.ausweis;
    };
    inpNach.addEventListener("change", tryNames);
    inpVor.addEventListener("change", tryNames);
  }

  // ===== 15 perces TIME PICKER =====
  function enhanceTimePicker(inp) {
    if (!inp || inp.dataset.enhanced === "1") return;
    inp.dataset.enhanced = "1";
    // iOS-on a natív time input sokszor kényelmetlen – elrejtjük, és két selectet adunk
    inp.type = "hidden";

    const box = document.createElement("div");
    box.style.display = "flex";
    box.style.gap = ".5rem";
    box.style.alignItems = "center";

    const selH = document.createElement("select");
    selH.style.minWidth = "5.2rem";
    selH.add(new Option("--", ""));
    for (let h = 0; h < 24; h++) {
      const v = String(h).padStart(2, "0");
      selH.add(new Option(v, v));
    }

    const selM = document.createElement("select");
    selM.style.minWidth = "5.2rem";
    selM.add(new Option("--", ""));
    ["00", "15", "30", "45"].forEach((m) => selM.add(new Option(m, m)));

    inp.insertAdjacentElement("afterend", box);
    box.appendChild(selH);
    box.appendChild(selM);

    function dispatchBoth() {
      inp.dispatchEvent(new Event("input", { bubbles: true }));
      inp.dispatchEvent(new Event("change", { bubbles: true }));
    }

    function compose() {
      const h = selH.value;
      const m = selM.value;
      const prev = inp.value;
      inp.value = h && m ? `${h}:${m}` : "";
      if (inp.value !== prev) dispatchBoth();
    }

    // külsőből hívható setter (új dolgozó beszúrásakor)
    inp._setFromValue = (v) => {
      const mm = /^\d{2}:\d{2}$/.test(v) ? v.split(":") : ["", ""];
      selH.value = mm[0] || "";
      const m = mm[1] || "";
      const allowed = ["00", "15", "30", "45"];
      selM.value = allowed.includes(m)
        ? m
        : m
        ? String(Math.round(parseInt(m, 10) / 15) * 15).padStart(2, "0")
        : "";
      compose();
    };

    inp._setFromValue(inp.value || "");
    selH.addEventListener("change", compose);
    selM.addEventListener("change", compose);
  }

  // ===== Óraszámítás =====
  const toTime = (s) => {
    if (!s || !/^\d{2}:\d{2}$/.test(s)) return null;
    const [hh, mm] = s.split(":").map(Number);
    if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
    return { hh, mm };
  };
  const minutes = (t) => t.hh * 60 + t.mm;
  const overlap = (a1, a2, b1, b2) => Math.max(0, Math.min(a2, b2) - Math.max(a1, b1));

  function hoursWithBreaks(begStr, endStr) {
    const bt = toTime(begStr);
    const et = toTime(endStr);
    if (!bt || !et) return 0;
    const b = minutes(bt),
      e = minutes(et);
    if (e <= b) return 0;

    let total = e - b;

    const bm = parseInt((breakHidden?.value || "60"), 10);
    const isHalfHour = !isNaN(bm) && bm === 30;

    if (isHalfHour) {
      total = Math.max(0, total - 30);
    } else {
      const br1s = 9 * 60 + 0,
        br1e = 9 * 60 + 15; // 09:00–09:15
      const br2s = 12 * 60 + 0,
        br2e = 12 * 60 + 45; // 12:00–12:45
      const minus = overlap(b, e, br1s, br1e) + overlap(b, e, br2s, br2e);
      total = Math.max(0, total - minus);
    }

    return Math.max(0, total) / 60;
  }

  const formatHours = (h) => (Math.round(h * 100) / 100).toFixed(2);

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
    workerList.querySelectorAll(".worker").forEach((w) => {
      sum += recalcWorker(w);
    });
    if (totalOut) totalOut.value = sum ? formatHours(sum) : "";
  }

  // ===== Szünet kapcsoló =====
  function updateBreakAndRecalc() {
    if (!breakHidden) return;
    breakHidden.value = breakHalf?.checked ? "30" : "60";
    recalcAll();
  }
  breakHalf?.addEventListener("change", updateBreakAndRecalc);
  updateBreakAndRecalc();

  // ===== Szinkron az 1. dolgozóról =====
  function markSynced(inp, isSynced) {
    if (!inp) return;
    isSynced ? (inp.dataset.synced = "1") : delete inp.dataset.synced;
  }
  function isSynced(inp) {
    return !!(inp && inp.dataset.synced === "1");
  }
  function setupManualEditUnsync(inp) {
    if (!inp) return;
    const unsync = () => markSynced(inp, false);
    inp.addEventListener("input", (e) => {
      if (e.isTrusted) unsync();
    });
    inp.addEventListener("change", (e) => {
      if (e.isTrusted) unsync();
    });
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

  function enforceNumericKeyboard(input) {
    input.setAttribute("inputmode", "numeric");
    input.setAttribute("pattern", "[0-9]*");
    input.addEventListener("input", () => {
      input.value = input.value.replace(/\D/g, "");
    });
  }

  function wireWorker(workerEl) {
    const idx = workerEl.getAttribute("data-index");

    // autocomplete
    setupAutocomplete(workerEl);

    // Ausweis numerikus
    const ausweis = workerEl.querySelector(`input[name="ausweis${idx}"]`);
    if (ausweis) enforceNumericKeyboard(ausweis);

    // idő mezők
    ["beginn", "ende"].forEach((prefix) => {
      const inp = workerEl.querySelector(`input[name^="${prefix}"]`);
      if (inp) {
        enhanceTimePicker(inp);
        inp.addEventListener("change", recalcAll);
        inp.addEventListener("input", recalcAll);
      }
    });

    const b = workerEl.querySelector(`input[name="beginn${idx}"]`);
    const e = workerEl.querySelector(`input[name="ende${idx}"]`);

    if (idx === "1") {
      markSynced(b, false);
      markSynced(e, false);
      b?.addEventListener("input", syncFromFirst);
      e?.addEventListener("input", syncFromFirst);
      b?.addEventListener("change", syncFromFirst);
      e?.addEventListener("change", syncFromFirst);
    } else {
      setupManualEditUnsync(b);
      setupManualEditUnsync(e);
    }
  }

  // első dolgozó bekötése
  wireWorker(workerList.querySelector(".worker"));
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
      <legend data-i18n="mitarbeiter">Mitarbeiter</legend> ${idx}
      <div class="grid-3">
        <div class="field">
          <label data-i18n="vorname">Vorname</label>
          <input name="vorname${idx}" type="text" />
        </div>
        <div class="field">
          <label data-i18n="nachname">Nachname</label>
          <input name="nachname${idx}" type="text" />
        </div>
        <div class="field">
          <label data-i18n="ausweis">Ausweis-Nr. / Kennzeichen</label>
          <input name="ausweis${idx}" type="text" />
        </div>
      </div>
      <div class="grid">
        <div class="field">
          <label data-i18n="vorhaltung">Vorhaltung / beauftragtes Gerät / Fahrzeug</label>
          <input name="vorhaltung${idx}" type="text" />
        </div>
      </div>
      <div class="grid-3">
        <div class="field">
          <label data-i18n="beginn">Beginn</label>
          <input name="beginn${idx}" type="time" />
        </div>
        <div class="field">
          <label data-i18n="ende">Ende</label>
          <input name="ende${idx}" type="time" />
        </div>
        <div class="field">
          <label data-i18n="stunden">Stunden (auto)</label>
          <input class="stunden-display" type="text" value="" readonly />
        </div>
      </div>
    `;
    workerList.appendChild(tpl);

    // friss fordítás az új blokkra is
    applyTranslations(currentLang);

    // szinkron az 1. dolgozóról
    const firstBeg = document.querySelector('input[name="beginn1"]')?.value || "";
    const firstEnd = document.querySelector('input[name="ende1"]')?.value || "";
    const begNew = tpl.querySelector(`input[name="beginn${idx}"]`);
    const endNew = tpl.querySelector(`input[name="ende${idx}"]`);
    if (firstBeg && begNew) begNew.value = firstBeg;
    if (firstEnd && endNew) endNew.value = firstEnd;

    wireWorker(tpl);

    // frissen beállított értékek tükrözése a selectekben
    if (begNew && begNew._setFromValue) begNew._setFromValue(begNew.value || "");
    if (endNew && endNew._setFromValue) endNew._setFromValue(endNew.value || "");

    if (begNew && begNew.value) markSynced(begNew, true);
    if (endNew && endNew.value) markSynced(endNew, true);

    recalcAll();
  });

  loadWorkers();

  // ==== Beschreibung számláló ====
  (function () {
    const besch = document.getElementById("beschreibung");
    const out = document.getElementById("besch-count");
    if (!besch || !out) return;
    const max = parseInt(besch.getAttribute("maxlength") || "1000", 10);
    function updateBeschCount() {
      const len = besch.value.length || 0;
      out.textContent = `${len} / ${max}`;
    }
    updateBeschCount();
    besch.addEventListener("input", updateBeschCount);
    besch.addEventListener("change", updateBeschCount);
  })();

  // ==== Validáció ====
  (function () {
    if (!form) return;
    const trim = (v) => (v || "").toString().trim();
    const readWorker = (fs) => {
      const idx = fs.getAttribute("data-index") || "";
      const q = (sel) => fs.querySelector(sel);
      return {
        idx,
        vorname: trim(q(`input[name="vorname${idx}"]`)?.value),
        nachname: trim(q(`input[name="nachname${idx}"]`)?.value),
        ausweis: trim(q(`input[name="ausweis${idx}"]`)?.value),
        beginn: trim(q(`input[name="beginn${idx}"]`)?.value),
        ende: trim(q(`input[name="ende${idx}"]`)?.value),
      };
    };

    form.addEventListener(
      "submit",
      function (e) {
        // ha PDF gomb indítja, ezt az ágat ne futtassuk
        if (
          e.submitter &&
          e.submitter.getAttribute("formaction") === "/generate_pdf"
        ) {
          return;
        }

        const errors = [];
        const datum = trim(document.getElementById("datum")?.value);
        const bau = trim(document.getElementById("bau")?.value);
        const bf = trim(document.getElementById("basf_beauftragter")?.value);
        const besch = trim(document.getElementById("beschreibung")?.value);

        if (!datum) errors.push("Bitte das Datum der Leistungsausführung angeben.");
        if (!bau) errors.push("Bitte Bau und Ausführungsort ausfüllen.");
        if (!bf) errors.push("Bitte den BASF-Beauftragten (Org.-Code) ausfüllen.");
        if (!besch) errors.push("Bitte die Beschreibung der ausgeführten Arbeiten ausfüllen.");

        const sets = Array.from(
          document.getElementById("worker-list").querySelectorAll(".worker")
        );
        let validWorkers = 0;

        sets.forEach((fs) => {
          const w = readWorker(fs);
          const anyFilled = !!(w.vorname || w.nachname || w.ausweis || w.beginn || w.ende);
          const allCore = !!(w.vorname && w.nachname && w.ausweis && w.beginn && w.ende);
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
          errors.push(
            "Bitte mindestens einen Mitarbeiter vollständig angeben (Vorname, Nachname, Ausweis-Nr., Beginn, Ende)."
          );
        }

        if (errors.length) {
          e.preventDefault();
          alert(errors.join("\n"));
          return false;
        }
        return true;
      },
      false
    );
  })();

  // ==== Excel letöltés (fetch + blob) ====
  (function () {
    if (!form) return;
    const submitBtn = form.querySelector('button[type="submit"]:not([formaction])');

    function setBusy(busy) {
      if (!submitBtn) return;
      if (busy) {
        submitBtn.dataset._label =
          submitBtn.textContent || submitBtn.value || "Generieren";
        submitBtn.disabled = true;
        const txt = window.__GEN_TEXT || "Wird generiert...";
        submitBtn.textContent = txt;
      } else {
        submitBtn.disabled = false;
        const lbl = submitBtn.dataset._label;
        if (lbl) submitBtn.textContent = lbl;
      }
    }

    const filenameFromDisposition = (h) => {
      if (!h) return null;
      const m =
        /filename\*?=(?:UTF-8''|")?([^\";]+)"/i.exec(h) ||
        /filename=([^;]+)/i.exec(h);
      if (!m) return null;
      try {
        return decodeURIComponent(m[1].replace(/"/g, "").trim());
      } catch {
        return m[1].replace(/"/g, "").trim();
      }
    };

    async function downloadOnce(fd) {
      const res = await fetch("/generate_excel", { method: "POST", body: fd });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);

      const ct = res.headers.get("content-type") || "";
      if (!ct.includes("spreadsheet")) {
        const txt = await res.text().catch(() => "");
        throw new Error(txt || "Ismeretlen hiba (nem érkezett fájl).");
      }
      const blob = await res.blob();
      const name =
        filenameFromDisposition(res.headers.get("content-disposition")) ||
        "leistungsnachweis.xlsx";

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    }

    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

    form.addEventListener(
      "submit",
      async (e) => {
        // ha PDF gomb indítja, most nem mi intézzük
        if (e.submitter && e.submitter.getAttribute("formaction") === "/generate_pdf") {
          return;
        }
        if (e.defaultPrevented) return; // validáció megállította

        e.preventDefault();
        const fd = new FormData(form);
        setBusy(true);
        try {
          await downloadOnce(fd);
        } catch (err1) {
          try {
            await sleep(1200);
            await downloadOnce(fd);
          } catch (err2) {
            console.error(err1, err2);
            alert("A fájl generálása most nem sikerült. Kérlek próbáld újra pár másodperc múlva.");
          }
        } finally {
          setBusy(false);
        }
      },
      false
    );
  })();

  // ==== (Opcionális) PDF Vorschau – csak akkor fut, ha van hozzá gomb ====
  (function () {
    if (!form) return;
    const pdfBtn = form.querySelector('button[formaction="/generate_pdf"]');
    if (!pdfBtn) return;

    const filenameFromDisposition = (h) => {
      if (!h) return null;
      const m =
        /filename\*?=(?:UTF-8''|")?([^\";]+)"/i.exec(h) ||
        /filename=([^;]+)/i.exec(h);
      if (!m) return null;
      try {
        return decodeURIComponent(m[1].replace(/"/g, "").trim());
      } catch {
        return m[1].replace(/"/g, "").trim();
      }
    };

    pdfBtn.addEventListener("click", async (ev) => {
      ev.preventDefault();
      ev.stopPropagation();

      const win = window.open("", "_blank");
      if (!win) {
        alert("A böngésző letiltotta a felugró ablakot. Engedélyezd, kérlek.");
        return;
      }

      win.document.open();
      win.document.write(`
        <!doctype html>
        <html><head><meta charset="utf-8"><title>PDF Vorschau</title>
        <style>html,body{height:100%;margin:0}.box{font:15px/1.4 system-ui; padding:24px}</style></head>
        <body><div class="box">PDF wird generiert…</div></body></html>
      `);
      win.document.close();

      try {
        const fd = new FormData(form);
        const res = await fetch("/generate_pdf", { method: "POST", body: fd });
        if (!res.ok) throw new Error("HTTP " + res.status);

        const ct = res.headers.get("content-type") || "";
        if (!ct.includes("pdf")) {
          const text = await res.text().catch(() => "");
          throw new Error(text || "Unerwartete Antwort – kein PDF.");
        }

        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const name =
          filenameFromDisposition(res.headers.get("content-disposition")) ||
          "leistungsnachweis_preview.pdf";

        win.document.open();
        win.document.write(`
          <!doctype html>
          <html><head><meta charset="utf-8"><title>${name}</title>
          <style>html,body,iframe{height:100%;width:100%;margin:0;border:0}</style>
          </head><body><iframe src="${url}" title="${name}" frameborder="0"></iframe></body></html>
        `);
        win.document.close();

        setTimeout(() => URL.revokeObjectURL(url), 120000);
      } catch (err) {
        win.document.open();
        win.document.write(`
          <!doctype html>
          <html><head><meta charset="utf-8"><title>Fehler</title>
          <style>body{font:14px/1.5 monospace; white-space:pre-wrap; padding:24px}</style>
          </head><body>PDF-Erzeugung fehlgeschlagen:\n${(err && err.message) || err}</body></html>
        `);
        win.document.close();
      }
    });
  })();
});
