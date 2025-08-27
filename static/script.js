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
  const debounce = (fn, ms) => { let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), ms); }; };
  async function fetchJSON(url, opts = {}) {
    const r = await fetch(url, { headers: { "Accept": "application/json" }, ...opts });
    if (!r.ok) throw new Error(`${r.status} ${r.statusText}`);
    const ct = r.headers.get("content-type") || "";
    return ct.includes("application/json") ? r.json() : r.text();
  }

  // ====== Backend config ======
  const defaultConfig = {
    features: { autocomplete: true, translation_button: true },
    i18n: { default: "de", available: ["de", "hr", "en"] },
  };
  let appConfig = defaultConfig;
  async function loadConfig() {
    try {
      const c = await fetchJSON("/api/config");
      appConfig = (c && c.features) ? c : defaultConfig;
    } catch { appConfig = defaultConfig; }
  }

  // ====== i18n ======
  const I18N = window.__i18n__ || null;
  const getLang = () => (I18N ? I18N.getLang() : "de");
  const applyI18n = () => { if (I18N) I18N.applyI18n(getLang()); };

  // ====== Fordítás (HR -> DE) ======
  const Translate = (() => {
    let btn, statusEl, srcTA, dstTA;
    async function translate() {
      const lang = getLang();
      const text = (srcTA?.value || "").trim();
      if (!text) return;
      btn.disabled = true;
      statusEl.textContent = lang === "hr" ? "Prevodim..." : "Übersetzen...";
      try {
        const body = JSON.stringify({ text, source: "hr", target: "de" });
        const r = await fetch("/api/translate", { method: "POST", headers: { "Content-Type": "application/json" }, body });
        const j = await r.json();
        if (!r.ok || !j || !j.translated) throw new Error((j && j.error) || "translate error");
        if (dstTA) dstTA.value = j.translated; else srcTA.value = j.translated;
        statusEl.textContent = lang === "hr" ? "Prijevod gotov." : "Übersetzung fertig.";
      } catch (e) {
        statusEl.textContent = lang === "hr" ? "Greška pri prijevodu." : "Übersetzung fehlgeschlagen.";
      } finally {
        btn.disabled = false;
        setTimeout(()=>statusEl.textContent="",3000);
      }
    }
    function init() {
      btn = $("#btn-translate-hr-de");
      if (!btn) return;
      statusEl = $("#translate-status");
      srcTA = $("#beschreibung");
      dstTA = $("#beschreibung_de"); // ha nincs, az srcTA-t írjuk felül
      const show = appConfig.features.translation_button && getLang() === "hr";
      btn.style.display = show ? "" : "none";
      btn.addEventListener("click", translate);
    }
    return { init };
  })();

  // ====== Autocomplete (datalist, egyszer tölt) ======
  const Auto = (() => {
    let dlLast, dlFirst, dlIds, cacheAll=null, inflight=null;

    async function loadAllWorkersOnce() {
      if (cacheAll) return cacheAll;
      if (inflight) return inflight;
      const tryFetch = async (url) => { try { const r = await fetchJSON(url); if (r && Array.isArray(r.items)) return r; } catch {} return null; };
      inflight = (async () => {
        let data = await tryFetch("/api/workers");
        if (!data) data = await tryFetch("/api/workers?q=");
        if (!data) data = await tryFetch("/workers");
        if (!data) data = await tryFetch("/workers?q=");
        if (!data) data = { items: [] };
        cacheAll = data; inflight = null; return cacheAll;
      })();
      return inflight;
    }

    const uniq = (arr)=>Array.from(new Set(arr.filter(Boolean)));
    function filterClient(items, q) {
      const s = (q || "").toLowerCase();
      if (!s) return items;
      return items.filter(w => ["last_name","first_name","id"].some(k => (w[k]||"").toLowerCase().includes(s)));
    }

    function fillDatalists(items) {
      if (dlLast) { dlLast.innerHTML=""; uniq(items.map(w=>(w.last_name||"").trim())).forEach(v=>{const o=document.createElement("option"); o.value=v; dlLast.appendChild(o);}); }
      if (dlFirst){ dlFirst.innerHTML=""; uniq(items.map(w=>(w.first_name||"").trim())).forEach(v=>{const o=document.createElement("option"); o.value=v; dlFirst.appendChild(o);}); }
      if (dlIds)  { dlIds.innerHTML=""; uniq(items.map(w=>(w.id||"").trim())).forEach(v=>{const o=document.createElement("option"); o.value=v; dlIds.appendChild(o);}); }
    }

    const debFill = debounce(async (q) => {
      const all = await loadAllWorkersOnce();
      fillDatalists(filterClient(all.items||[], q));
    }, 150);

    function wireRow(i){
      const ln = $(`#nachname${i}`), fn = $(`#vorname${i}`), id=$(`#ausweis${i}`);
      if (!ln||!fn||!id) return;
      const handler = () => { const q=(ln.value||fn.value||id.value||"").trim(); debFill(q); };
      ln.addEventListener("input", handler);
      fn.addEventListener("input", handler);
      id.addEventListener("input", handler);
    }

    function init(){
      dlLast=$("#dl-lastnames"); dlFirst=$("#dl-firstnames"); dlIds=$("#dl-ids");
      for(let i=1;i<=5;i++) wireRow(i);
      loadAllWorkersOnce().then(all=>fillDatalists(all.items||[])).catch(()=>{});
    }
    return { init };
  })();

  // ====== 1. sor időmásolása üres sorokra ======
  const TimeCopy = (() => {
    const copyIfEmpty = (srcSel,dstSel)=>{ const s=$(srcSel),d=$(dstSel); if(!s||!d) return; if(d.value.trim()==="" && s.value.trim()!=="") d.value=s.value; };
    function applyCopy(){ for(let i=2;i<=5;i++){ copyIfEmpty("#beginn1",`#beginn${i}`); copyIfEmpty("#ende1",`#ende${i}`);} }
    function init(){ const b1=$("#beginn1"), e1=$("#ende1"); if(b1) b1.addEventListener("change",applyCopy); if(e1) e1.addEventListener("change",applyCopy); }
    return { init };
  })();

  // ====== PDF gomb ======
  const PDF = (() => {
    function init(){
      const btn=$("#btn-pdf"), form=$("#main-form");
      if(!btn||!form) return;
      btn.addEventListener("click", async ()=>{
        try{
          const fd=new FormData(form);
          const r=await fetch("/generate_pdf",{method:"POST", body:fd});
          if(!r.ok) throw new Error("pdf");
          const blob=await r.blob();
          const url=URL.createObjectURL(blob);
          window.open(url,"_blank");
          setTimeout(()=>URL.revokeObjectURL(url),15000);
        }catch(e){
          alert(getLang()==="hr"?"PDF nije dostupan na poslužitelju.":"PDF előállítás nem elérhető ezen a szerveren.");
        }
      });
    }
    return { init };
  })();

  // ====== Piszkozat (localStorage) ======
  const Draft = (() => {
    const KEY = "pdfedit.form.v1";
    const fields = [
      "datum","bau","basf_beauftragter","beschreibung","break_minutes",
      "nachname1","vorname1","ausweis1","beginn1","ende1","vorhaltung1",
      "nachname2","vorname2","ausweis2","beginn2","ende2","vorhaltung2",
      "nachname3","vorname3","ausweis3","beginn3","ende3","vorhaltung3",
      "nachname4","vorname4","ausweis4","beginn4","ende4","vorhaltung4",
      "nachname5","vorname5","ausweis5","beginn5","ende5","vorhaltung5",
    ];
    const save = () => {
      const obj = {}; fields.forEach(id=>{ const el=document.getElementById(id); if(el) obj[id]=el.value; });
      try{ localStorage.setItem(KEY, JSON.stringify(obj)); }catch{}
    };
    const load = () => {
      try{
        const s=localStorage.getItem(KEY); if(!s) return; const obj=JSON.parse(s);
        fields.forEach(id=>{ const el=document.getElementById(id); if(el && obj[id]!=null && el.value==="") el.value=obj[id]; });
      }catch{}
    };
    function init(){
      load();
      fields.forEach(id=>{ const el=document.getElementById(id); if(el) el.addEventListener("input", debounce(save,200)); });
      const form=$("#main-form"); if(form) form.addEventListener("submit", ()=>{ try{localStorage.removeItem(KEY);}catch{} });
    }
    return { init };
  })();

  // ====== Boot ======
  document.addEventListener("DOMContentLoaded", async () => {
    await loadConfig();     // 1) config
    applyI18n();            // 2) i18n
    Translate.init();       // 3) fordítás
    if (appConfig.features.autocomplete) Auto.init();  // 4) autocomplete
    TimeCopy.init();        // 5) időmásolás
    PDF.init();             // 6) PDF
    Draft.init();           // 7) piszkozat
  });
})();
