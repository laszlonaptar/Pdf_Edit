/* static/script.js — teljes front logika
   - i18n integráció (i18n.js)
   - fordítás gomb (HR -> DE) /api/translate
   - autocomplete (datalist) /api/workers (fallback: /workers)
   - 1. sor idő átmásolása üres sorokra
   - PDF előnézet /generate_pdf
   - piszkozat mentés localStorage-ben
*/

(function () {
  "use strict";

  // ====== Util ======
  const $ = (sel, root) => (root || document).querySelector(sel);
  const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const debounce = (fn, ms) => {
    let t;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn.apply(null, args), ms);
    };
  };
  function toJSONSafe(r) {
    const ct = r.headers.get("content-type") || "";
    if (ct.includes("application/json")) return r.json();
    return r.text();
  }
  async function fetchJSON(url, opts = {}) {
    const r = await fetch(url, { headers: { "Accept": "application/json" }, ...opts });
    if (!r.ok) throw new Error(`${r.status} ${r.statusText}`);
    return toJSONSafe(r);
  }

  // ====== Backend config ======
  const defaultConfig = {
    features: { autocomplete: true, translation_button: true },
    i18n: { default: "de", available: ["de", "hr", "en"] },
  };
  let appConfig = defaultConfig;
  async function loadConfig() {
    try {
      appConfig = await fetchJSON("/api/config");
      if (!appConfig || !appConfig.features) appConfig = defaultConfig;
    } catch (e) {
      console.warn("Config fetch failed – using defaults.", e);
      appConfig = defaultConfig;
    }
  }

  // ====== i18n ======
  const I18N = window.__i18n__ || null;
  function getLang() {
    return I18N ? I18N.getLang() : "de";
  }
  function applyI18n() {
    if (I18N) I18N.applyI18n(getLang());
  }

  // ====== Fordítás (HR -> DE) ======
  const Translate = (() => {
    let btn, statusEl, srcTA, dstTA;
    async function translate() {
      const lang = getLang();
      const txt = (srcTA?.value || "").trim();
      if (!txt) return;
      btn.disabled = true;
      statusEl.textContent = lang === "hr" ? "Prevodim..." : "Übersetzen...";
      try {
        const body = JSON.stringify({ text: txt, source: "hr", target: "de" });
        const r = await fetch("/api/translate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body,
        });
        const j = await r.json();
        if (!r.ok || !j || !j.translated) throw new Error((j && j.error) || "translate error");
        if (dstTA) dstTA.value = j.translated;
        else srcTA.value = j.translated;
        statusEl.textContent = lang === "hr" ? "Prijevod gotov." : "Übersetzung fertig.";
      } catch (e) {
        console.error(e);
        statusEl.textContent = lang === "hr" ? "Greška pri prijevodu." : "Übersetzung fehlgeschlagen.";
      } finally {
        btn.disabled = false;
        setTimeout(() => (statusEl.textContent = ""), 3000);
      }
    }
    function init() {
      btn = $("#btn-translate-hr-de");
      statusEl = $("#translate-status");
      srcTA = $("#beschreibung");
      dstTA = $("#beschreibung_de"); // opcionális
      if (!btn) return;
      const show = appConfig.features.translation_button && getLang() === "hr";
      btn.style.display = show ? "" : "none";
      btn.addEventListener("click", translate);
    }
    return { init };
  })();

  // ====== Autocomplete (datalist, cache) ======
  const Auto = (() => {
    let dlLast, dlFirst, dlIds;
    let cacheAll = null; // {items:[...]} – egyszer töltsük le
    let inflight = null;

    // Betöltjük az összes dolgozót egyszer (ha a backend támogatja az üres q-t)
    async function loadAllWorkersOnce() {
      if (cacheAll) return cacheAll;
      if (inflight) return inflight;
      async function tryFetch(url) {
        try {
          const res = await fetchJSON(url);
          if (res && Array.isArray(res.items)) return res;
        } catch (e) {
          /* ignore */
        }
        return null;
      }
      inflight = (async () => {
        // Első próbálkozás: /api/workers (üres q)
        let data = await tryFetch("/api/workers");
        if (!data) data = await tryFetch("/api/workers?q=");
        if (!data) data = await tryFetch("/workers");
        if (!data) data = await tryFetch("/workers?q=");
        // Utolsó fallback: üres lista
        if (!data) data = { items: [] };
        cacheAll = data;
        inflight = null;
        return cacheAll;
      })();
      return inflight;
    }

    function uniq(arr) {
      return Array.from(new Set(arr.filter(Boolean)));
    }
    function filterClient(items, q) {
      const s = (q || "").toLowerCase();
      if (!s) return items;
      return items.filter((w) =>
        ["last_name", "first_name", "id"].some((k) => ((w[k] || "") + "").toLowerCase().includes(s))
      );
    }

    function fillDatalists(items) {
      if (dlLast) {
        dlLast.innerHTML = "";
        uniq(items.map((w) => (w.last_name || "").trim())).forEach((v) => {
          const o = document.createElement("option");
          o.value = v;
          dlLast.appendChild(o);
        });
      }
      if (dlFirst) {
        dlFirst.innerHTML = "";
        uniq(items.map((w) => (w.first_name || "").trim())).forEach((v) => {
          const o = document.createElement("option");
          o.value = v;
          dlFirst.appendChild(o);
        });
      }
      if (dlIds) {
        dlIds.innerHTML = "";
        uniq(items.map((w) => (w.id || "").trim())).forEach((v) => {
          const o = document.createElement("option");
          o.value = v;
          dlIds.appendChild(o);
        });
      }
    }

    const debouncedFill = debounce(async (q) => {
      const all = await loadAllWorkersOnce();
      const items = filterClient(all.items || [], q);
      fillDatalists(items);
    }, 150);

    function wireRow(i) {
      const ln = $(`#nachname${i}`);
      const fn = $(`#vorname${i}`);
      const id = $(`#ausweis${i}`);
      if (!ln || !fn || !id) return;

      const handler = () => {
        const q = (ln.value || fn.value || id.value || "").trim();
        debouncedFill(q);
      };
      ln.addEventListener("input", handler);
      fn.addEventListener("input", handler);
      id.addEventListener("input", handler);
    }

    function init() {
      dlLast = $("#dl-lastnames");
      dlFirst = $("#dl-firstnames");
      dlIds = $("#dl-ids");
      for (let i = 1; i <= 5; i++) wireRow(i);
      // első betöltés (cache melegítés)
      loadAllWorkersOnce().then((all) => fillDatalists(all.items || [])).catch(() => {});
    }

    return { init };
  })();

  // ====== 1. sor időmásolása üres sorokra ======
  const TimeCopy = (() => {
    function getTimeVal(id) {
      const el = $(id);
      return el ? (el.value || "").trim() : "";
    }
    function copyIfEmpty(srcId, dstId) {
      const src = $(srcId), dst = $(dstId);
      if (!src || !dst) return;
      if ((dst.value || "").trim() === "" && (src.value || "").trim() !== "") {
        dst.value = src.value;
        dst.dispatchEvent(new Event("change"));
      }
    }
    function applyCopy() {
      const b1 = getTimeVal("#beginn1");
      const e1 = getTimeVal("#ende1");
      if (!b1 && !e1) return;
      for (let i = 2; i <= 5; i++) {
        copyIfEmpty("#beginn1", `#beginn${i}`);
        copyIfEmpty("#ende1", `#ende${i}`);
      }
    }
    function init() {
      const b1 = $("#beginn1");
      const e1 = $("#ende1");
      if (b1) b1.addEventListener("change", applyCopy);
      if (e1) e1.addEventListener("change", applyCopy);
    }
    return { init };
  })();

  // ====== PDF gomb ======
  const PDF = (() => {
    function init() {
      const btn = $("#btn-pdf");
      const form = $("#main-form");
      if (!btn || !form) return;
      btn.addEventListener("click", async () => {
        try {
          const fd = new FormData(form);
          const r = await fetch("/generate_pdf", { method: "POST", body: fd });
          if (!r.ok) throw new Error("PDF gen failed");
          const blob = await r.blob();
          const url = URL.createObjectURL(blob);
          window.open(url, "_blank");
          setTimeout(() => URL.revokeObjectURL(url), 15000);
        } catch (e) {
          alert(getLang() === "hr" ? "PDF nije dostupan na poslužitelju." : "PDF előállítás nem elérhető ezen a szerveren.");
        }
      });
    }
    return { init };
  })();

  // ====== Piszkozat mentés (localStorage) ======
  const Draft = (() => {
    const KEY = "pdfedit.form.v1";
    const fields = [
      "datum", "bau", "basf_beauftragter", "beschreibung", "break_minutes",
      "nachname1", "vorname1", "ausweis1", "beginn1", "ende1", "vorhaltung1",
      "nachname2", "vorname2", "ausweis2", "beginn2", "ende2", "vorhaltung2",
      "nachname3", "vorname3", "ausweis3", "beginn3", "ende3", "vorhaltung3",
      "nachname4", "vorname4", "ausweis4", "beginn4", "ende4", "vorhaltung4",
      "nachname5", "vorname5", "ausweis5", "beginn5", "ende5", "vorhaltung5",
    ];
    function save() {
      const obj = {};
      for (const id of fields) {
        const el = document.getElementById(id);
        if (el) obj[id] = el.value;
      }
      try { localStorage.setItem(KEY, JSON.stringify(obj)); } catch (e) {}
    }
    function load() {
      try {
        const s = localStorage.getItem(KEY);
        if (!s) return;
        const obj = JSON.parse(s);
        for (const id of fields) {
          const el = document.getElementById(id);
          if (el && obj[id] != null && el.value === "") el.value = obj[id];
        }
      } catch (e) {}
    }
    function init() {
      load();
      // mentés inputokra
      fields.forEach((id) => {
        const el = document.getElementById(id);
        if (el) el.addEventListener("input", debounce(save, 200));
      });
      const form = $("#main-form");
      if (form) {
        form.addEventListener("submit", () => {
          try { localStorage.removeItem(KEY); } catch (e) {}
        });
      }
    }
    return { init, save, load };
  })();

  // ====== Boot ======
  document.addEventListener("DOMContentLoaded", async () => {
    // 1) config
    await loadConfig();

    // 2) i18n
    applyI18n();

    // 3) fordítás gomb
    Translate.init();

    // 4) autocomplete
    if (appConfig.features.autocomplete) Auto.init();

    // 5) időmásolás
    TimeCopy.init();

    // 6) PDF gomb
    PDF.init();

    // 7) piszkozat
    Draft.init();
  });
})();
