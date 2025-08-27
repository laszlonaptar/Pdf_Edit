/* static/script.js — DE/HR i18n, fordítás (HR->DE), autocomplete, időmásolás, PDF, piszkozat */
(function () {
  "use strict";

  // ===== Util =====
  const $ = (s, r) => (r || document).querySelector(s);
  const debounce = (fn, ms) => { let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), ms); }; };
  async function fetchJSON(url, opts = {}) {
    const r = await fetch(url, { headers: { "Accept": "application/json" }, ...opts });
    if (!r.ok) throw new Error(`${r.status} ${r.statusText}`);
    const ct = r.headers.get("content-type") || "";
    return ct.includes("application/json") ? r.json() : r.text();
  }

  // ===== Config (ha leáll, van default) =====
  const defaultConfig = { features: { autocomplete: true, translation_button: true }, i18n: { default: "de", available: ["de","hr"] } };
  let appConfig = defaultConfig;
  async function loadConfig() {
    try {
      const c = await fetchJSON("/api/config");
      appConfig = (c && c.features) ? c : defaultConfig;
    } catch { appConfig = defaultConfig; }
  }

  // ===== i18n =====
  const I18N = window.__i18n__ || null;
  const getLang = () => (I18N ? I18N.getLang() : "de");
  const t = (key) => { // biztonsági fallback
    const lang = getLang();
    const dict = (window.__i18nDict || null);
    return (dict && dict[lang] && dict[lang][key]) || key;
  };
  const applyI18n = () => { if (I18N) I18N.applyI18n(getLang()); };

  // ===== Fordítás (HR -> DE) =====
  const Translate = (() => {
    let btn, statusEl, srcTA, dstTA;
    function setStatus(msg, type) {
      statusEl.textContent = msg || "";
      statusEl.className = "status" + (type ? " "+type : "");
    }
    async function translate() {
      const lang = getLang();
      const text = (srcTA?.value || "").trim();
      if (!text) return;
      btn.disabled = true;
      setStatus(lang === "hr" ? "Prevodim…" : "Übersetzen…", "info");
      try {
        const body = JSON.stringify({ text, source: "hr", target: "de" });
        const r = await fetch("/api/translate", { method: "POST", headers: { "Content-Type": "application/json" }, body });
        const j = await r.json();
        if (!r.ok || !j || !j.translated) {
          const err = (j && j.error) ? j.error : `HTTP ${r.status}`;
          throw new Error(err);
        }
        if (dstTA) dstTA.value = j.translated; else srcTA.value = j.translated;
        setStatus(lang === "hr" ? "Prijevod gotov." : "Übersetzung fertig.", "ok");
      } catch (e) {
        console.error("translate failed:", e);
        setStatus(getLang()==="hr" ? "Greška pri prijevodu." : "Übersetzung fehlgeschlagen.", "err");
      } finally {
        btn.disabled = false;
        setTimeout(()=> setStatus("", ""), 3500);
      }
    }
    function init() {
      btn = $("#btn-translate-hr-de");
      statusEl = $("#translate-status");
      srcTA = $("#beschreibung");
      dstTA = $("#beschreibung_de"); // ha nincs, az src-t írjuk felül
      if (!btn || !statusEl || !srcTA) return;
      const show = appConfig.features.translation_button && getLang()==="hr";
      btn.style.display = show ? "" : "none";
      btn.addEventListener("click", translate);
    }
    return { init };
  })();

  // ===== Autocomplete (datalist, egyszer tölt) =====
  const Auto = (() => {
    let dlLast, dlFirst, dlIds, cacheAll=null, inflight=null;

    async function loadAllOnce() {
      if (cacheAll) return cacheAll;
      if (inflight) return inflight;
      const tryFetch = async (url) => { try{ const r = await fetchJSON(url); if (r && Array.isArray(r.items)) return r; }catch{} return null; };
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
    function fillDls(items) {
      if (dlLast) { dlLast.innerHTML=""; uniq(items.map(w=>(w.last_name||"").trim())).forEach(v=>{ const o=document.createElement("option"); o.value=v; dlLast.appendChild(o); }); }
      if (dlFirst){ dlFirst.innerHTML=""; uniq(items.map(w=>(w.first_name||"").trim())).forEach(v=>{ const o=document.createElement("option"); o.value=v; dlFirst.appendChild(o); }); }
      if (dlIds)  { dlIds.innerHTML=""; uniq(items.map(w=>(w.id||"").trim())).forEach(v=>{ const o=document.createElement("option"); o.value=v; dlIds.appendChild(o); }); }
    }

    const debFill = debounce(async (q)=>{
      const all = await loadAllOnce();
      const s = (q||"").toLowerCase();
      const items = !s ? all.items : all.items.filter(w =>
        ["last_name","first_name","id"].some(k => (w[k]||"").toLowerCase().includes(s))
      );
      fillDls(items);
    }, 120);

    function wireRow(i) {
      const ln = $(`#nachname${i}`), fn = $(`#vorname${i}`), id = $(`#ausweis${i}`);
      if (!ln || !fn || !id) return;
      const h = () => debFill((ln.value||fn.value||id.value||"").trim());
      ln.addEventListener("input", h);
      fn.addEventListener("input", h);
      id.addEventListener("input", h);
    }

    function init() {
      dlLast = $("#dl-lastnames");
      dlFirst = $("#dl-firstnames");
      dlIds = $("#dl-ids");
      for (let i=1;i<=5;i++) wireRow(i);
      loadAllOnce().then(all=> fillDls(all.items||[])).catch(()=>{});
    }
    return { init };
  })();

  // ===== 1. sor időmásolása üres sorokra =====
  const TimeCopy = (() => {
    const copyIfEmpty = (srcSel,dstSel)=>{ const s=$(srcSel), d=$(dstSel); if (!s||!d) return; if (d.value.trim()==="" && s.value.trim()!=="") d.value=s.value; };
    function apply(){ for(let i=2;i<=5;i++){ copyIfEmpty("#beginn1",`#beginn${i}`); copyIfEmpty("#ende1",`#ende${i}`); } }
    function init(){ const b1=$("#beginn1"), e1=$("#ende1"); if(b1) b1.addEventListener("change",apply); if(e1) e1.addEventListener("change",apply); }
    return { init };
  })();

  // ===== PDF gomb =====
  const PDF = (() => {
    function init(){
      const btn=$("#btn-pdf"), form=$("#main-form");
      if(!btn || !form) return;
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

  // ===== Piszkozat mentés =====
  const Draft = (() => {
    const KEY="pdfedit.form.v1";
    const fields=["datum","bau","basf_beauftragter","beschreibung","break_minutes",
      "nachname1","vorname1","ausweis1","beginn1","ende1","vorhaltung1",
      "nachname2","vorname2","ausweis2","beginn2","ende2","vorhaltung2",
      "nachname3","vorname3","ausweis3","beginn3","ende3","vorhaltung3",
      "nachname4","vorname4","ausweis4","beginn4","ende4","vorhaltung4",
      "nachname5","vorname5","ausweis5","beginn5","ende5","vorhaltung5"];
    const save = ()=> {
      const obj={}; fields.forEach(id=>{ const el=document.getElementById(id); if(el) obj[id]=el.value; });
      try{ localStorage.setItem(KEY, JSON.stringify(obj)); }catch{}
    };
    const load = ()=> {
      try{
        const s=localStorage.getItem(KEY); if(!s) return; const obj=JSON.parse(s);
        fields.forEach(id=>{ const el=document.getElementById(id); if(el && obj[id]!=null && el.value==="") el.value=obj[id]; });
      }catch{}
    };
    function init(){
      load();
      const deb = debounce(save,200);
      fields.forEach(id=>{ const el=document.getElementById(id); if(el) el.addEventListener("input", deb); });
      const form=$("#main-form"); if(form) form.addEventListener("submit", ()=>{ try{localStorage.removeItem(KEY);}catch{} });
    }
    return { init };
  })();

  // ===== Boot =====
  document.addEventListener("DOMContentLoaded", async ()=>{
    await loadConfig();
    applyI18n();
    Translate.init();
    if (appConfig.features.autocomplete) Auto.init();
    TimeCopy.init();
    PDF.init();
    Draft.init();
  });
})();
