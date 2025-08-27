/* Frontend logika:
   - i18n (DE/HR)
   - Fordítás HR->DE (/api/translate) – részletes hibaüzenet
   - Autocomplete: csak gépelés után jelenik meg, substring + prefix prioritás,
     találatra kitölti a vezetéknév, keresztnév, ausweis mezőket.
   - Munkaidő: natív time input, 15 perces léptetés (step=900)
   - Pauza: 60 alap, checkbox -> 30 (rejtett break_minutes)
   - 1. sor idő másolása az üres 2–5. sorokra (felülírható)
   - PDF előnézet
   - Piszkozat mentés (localStorage)
*/
(function () {
  "use strict";
  const $  = (s, r) => (r || document).querySelector(s);
  const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
  const debounce = (fn, ms) => { let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), ms); }; };

  // ---- i18n helpers
  const I18N = window.__i18n__ || null;
  const DICT = window.__i18nDict || {};
  const getLang = () => (I18N ? I18N.getLang() : "de");
  const T = (k) => (DICT[getLang()] && DICT[getLang()][k]) || k;

  // ---- Pauza
  function initBreakSwitch() {
    const chk = $("#half_break");
    const hid = $("#break_minutes");
    if (!chk || !hid) return;
    const apply = () => { hid.value = chk.checked ? "30" : "60"; };
    chk.addEventListener("change", apply);
    apply();
  }

  // ---- Fordítás HR->DE
  function initTranslate() {
    const btn = $("#btn-translate-hr-de");
    const status = $("#translate-status");
    const src = $("#beschreibung");
    const dst = $("#beschreibung_de");
    if (!btn || !src || !status) return;

    btn.addEventListener("click", async () => {
      const text = (src.value || "").trim();
      if (!text) return;
      btn.disabled = true;
      status.textContent = T("t_busy");
      status.className = "status info";
      try {
        const r = await fetch("/api/translate", {
          method: "POST",
          headers: { "Content-Type":"application/json" },
          body: JSON.stringify({ text, source:"hr", target:"de" })
        });
        const ct = r.headers.get("content-type") || "";
        let j=null; try{ j = ct.includes("application/json") ? await r.json() : null; }catch{}
        if (!r.ok || !j || typeof j.translated!=="string") {
          const msg = (j && j.error) ? j.error : `HTTP ${r.status}`;
          throw new Error(msg);
        }
        if (dst) dst.value = j.translated; else src.value = j.translated;
        status.textContent = T("t_done");
        status.className = "status ok";
      } catch (e) {
        status.textContent = `${T("t_error")} — ${e && e.message ? e.message : ""}`;
        status.className = "status err";
      } finally {
        btn.disabled = false;
        setTimeout(()=>{ status.textContent=""; status.className="status"; }, 4000);
      }
    });
  }

  // ---- Autocomplete
  const AC = (() => {
    let cache = null, inflight = null;

    function score(item, q) {
      // Prefix találat előre
      const last  = (item.last_name  || "").toLowerCase();
      const first = (item.first_name || "").toLowerCase();
      const id    = (item.id         || "").toLowerCase();
      const s = q.toLowerCase();

      const isPrefix = [last, first, id].some(v => v.startsWith(s));
      const hasSub   = [last, first, id].some(v => v.includes(s));
      if (isPrefix) return 0;
      if (hasSub)   return 1;
      return 9;
    }

    async function loadAll() {
      if (cache) return cache;
      if (inflight) return inflight;
      const tryFetch = async (url) => {
        try {
          const r = await fetch(url, { headers: { "Accept":"application/json" }});
          if (!r.ok) return null;
          const j = await r.json();
          if (j && Array.isArray(j.items)) return j.items;
        } catch {}
        return null;
      };
      inflight = (async () => {
        let items = await tryFetch("/api/workers");
        if (!items) items = await tryFetch("/api/workers?q=");
        if (!items) items = await tryFetch("/workers");
        if (!items) items = await tryFetch("/workers?q=");
        cache = items || [];
        return cache;
      })();
      return inflight;
    }

    function attachRow(rowIdx) {
      const row = document.querySelector(`.worker-row[data-row="${rowIdx}"]`);
      if (!row) return;
      const ln = $(`#nachname${rowIdx}`), fn = $(`#vorname${rowIdx}`), id = $(`#ausweis${rowIdx}`);
      const list = $(`#ac-${rowIdx}`);

      function close() { list.style.display = "none"; list.innerHTML = ""; }
      function openWith(items) {
        list.innerHTML = "";
        if (!items.length) return close();
        for (const it of items.slice(0, 12)) {
          const d = document.createElement("div");
          d.className = "ac-item";
          d.textContent = `${it.last_name || ""} ${it.first_name || ""} – ${it.id || ""}`.trim();
          d.addEventListener("mousedown", (ev) => {
            ev.preventDefault();
            if (ln) ln.value = it.last_name  || "";
            if (fn) fn.value = it.first_name || "";
            if (id) id.value = it.id         || "";
            close();
          });
          list.appendChild(d);
        }
        list.style.display = "block";
      }

      const update = debounce(async () => {
        const q = (ln.value || fn.value || id.value || "").trim();
        if (!q) return close();               // csak gépelésre nyitunk
        const all = await loadAll();
        const filtered = all
          .map(it => ({ it, sc: score(it, q) }))
          .filter(x => x.sc < 9)
          .sort((a,b)=>a.sc-b.sc)
          .map(x => x.it);
        openWith(filtered);
      }, 120);

      // Csak inputra nyitunk (focusra NEM)
      [ln, fn, id].forEach(el => {
        if (!el) return;
        el.addEventListener("input", update);
        el.addEventListener("keydown", (e) => { if (e.key === "Escape") close(); });
        el.addEventListener("blur", () => setTimeout(close, 150));
      });
    }

    async function init() {
      await loadAll();
      for (let i=1;i<=5;i++) attachRow(i);
    }
    return { init };
  })();

  // ---- 1. sor idő átmásolása üres sorokra (bármikor felülírható)
  function initTimeCopy() {
    const b1 = $("#beginn1"), e1 = $("#ende1");
    if (!b1 || !e1) return;

    function copy() {
      for (let i=2;i<=5;i++) {
        const bi = $(`#beginn${i}`), ei = $(`#ende${i}`);
        if (bi && !bi.value) bi.value = b1.value;
        if (ei && !ei.value) ei.value = e1.value;
      }
    }
    // Induláskor és ha az első sor ideje változik
    copy();
    b1.addEventListener("change", copy);
    e1.addEventListener("change", copy);
  }

  // ---- PDF
  function initPdf() {
    const btn = $("#btn-pdf");
    const form = $("#main-form");
    if (!btn || !form) return;
    btn.addEventListener("click", async () => {
      try{
        const fd = new FormData(form);
        const r = await fetch("/generate_pdf",{ method:"POST", body:fd });
        if (!r.ok) throw new Error("pdf");
        const blob = await r.blob();
        const url = URL.createObjectURL(blob);
        window.open(url,"_blank");
        setTimeout(()=>URL.revokeObjectURL(url),15000);
      } catch {
        alert(getLang()==="hr" ? "PDF nije dostupan na poslužitelju." : "PDF előállítás nem elérhető ezen a szerveren.");
      }
    });
  }

  // ---- Piszkozat mentés
  const Draft = (() => {
    const KEY="pdfedit.form.v1";
    const F = ["datum","bau","basf_beauftragter","beschreibung","break_minutes",
      "nachname1","vorname1","ausweis1","beginn1","ende1","vorhaltung1",
      "nachname2","vorname2","ausweis2","beginn2","ende2","vorhaltung2",
      "nachname3","vorname3","ausweis3","beginn3","ende3","vorhaltung3",
      "nachname4","vorname4","ausweis4","beginn4","ende4","vorhaltung4",
      "nachname5","vorname5","ausweis5","beginn5","ende5","vorhaltung5"];
    function save(){
      const o={}; F.forEach(id=>{ const el=document.getElementById(id); if(el) o[id]=el.value; });
      try{ localStorage.setItem(KEY, JSON.stringify(o)); }catch{}
    }
    function load(){
      try{
        const s=localStorage.getItem(KEY); if(!s) return; const o=JSON.parse(s);
        F.forEach(id=>{ const el=document.getElementById(id); if(el && o[id]!=null && el.value==="") el.value=o[id]; });
      }catch{}
    }
    function init(){
      load();
      const deb = debounce(save,200);
      F.forEach(id=>{ const el=document.getElementById(id); if(el) el.addEventListener("input", deb); });
      const form=$("#main-form"); if(form) form.addEventListener("submit", ()=>{ try{localStorage.removeItem(KEY);}catch{} });
    }
    return { init };
  })();

  // ---- Boot
  document.addEventListener("DOMContentLoaded", async () => {
    if (I18N) I18N.applyI18n(getLang());
    initBreakSwitch();
    initTranslate();
    await AC.init();
    initTimeCopy();
    initPdf();
    Draft.init();
  });
})();
