/* static/script.js — teljes (fixelt változat)
   - I18N
   - HR felületen „Prevedi na njemački” gomb (Azure /api/translate)
   - Autocomplete: /api/workers (fallback: /static/workers.csv)
   - 15 perces time picker, óraszámítás, 1. dolgozó időinek másolása
   - Validáció, Excel letöltés, PDF preview
*/

document.addEventListener("DOMContentLoaded", () => {
  // ---------- I18N ----------
  const I18N = (window.I18N || {});
  const DEFAULT_LANG = "de";
  const form  = document.getElementById("ln-form");
  const besch = document.getElementById("beschreibung"); // <— ETTŐL KEZDVE CSAK 'besch'-et használunk

  function getLang() {
    const u  = new URL(window.location.href);
    const ql = (u.searchParams.get("lang") || "").trim();
    const ls = (localStorage.getItem("app_lang") || "").trim();
    const hl = (document.documentElement.getAttribute("lang") || "").trim();
    return (ql || ls || hl || DEFAULT_LANG);
  }
  let currentLang = getLang();

  function t(key, lang) {
    const L = lang || currentLang;
    try {
      if (I18N[L] && key in I18N[L]) return I18N[L][key];
      if (I18N[DEFAULT_LANG] && key in I18N[DEFAULT_LANG]) return I18N[DEFAULT_LANG][key];
    } catch(_) {}
    return null;
  }

  function translateNode(el, lang) {
    if (!el) return;
    const key = el.getAttribute("data-i18n");
    if (key) {
      const txt = t(key, lang);
      if (txt != null) el.textContent = txt;
    }
    const phKey = el.getAttribute && el.getAttribute("data-i18n-ph");
    if (phKey && "placeholder" in el) {
      const ph = t(phKey, lang);
      if (ph != null) el.placeholder = ph;
    }
    const valKey = el.getAttribute && el.getAttribute("data-i18n-value");
    if (valKey && "value" in el) {
      const v = t(valKey, lang);
      if (v != null) el.value = v;
    }
  }

  function setHiddenLang(lang) {
    const hidden = document.getElementById("lang_hidden");
    if (hidden) hidden.value = lang;
  }

  function applyTranslations(lang) {
    const ttl = t("title", lang);
    if (ttl) document.title = ttl;
    document.querySelectorAll("[data-i18n], [data-i18n-ph], [data-i18n-value]")
      .forEach((el) => translateNode(el, lang));
    const gen = t("generating", lang);
    if (gen) window.__GEN_TEXT = gen;

    toggleTranslateUI(lang);
    setHiddenLang(lang);
  }

  const langSel = document.getElementById("lang") || document.getElementById("lang-select");
  if (langSel && [...langSel.options].some(o => o.value === currentLang)) {
    langSel.value = currentLang;
    langSel.addEventListener("change", () => {
      currentLang = langSel.value || DEFAULT_LANG;
      localStorage.setItem("app_lang", currentLang);
      document.documentElement.setAttribute("lang", currentLang);
      applyTranslations(currentLang);
    });
  }
  document.documentElement.setAttribute("lang", currentLang);
  applyTranslations(currentLang);

  // ---------- segédek ----------
  const $ = (s, r=document) => r.querySelector(s);
  async function postJSON(url, body){
    const r = await fetch(url, {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      credentials:'same-origin',
      body: JSON.stringify(body)
    });
    if(!r.ok){
      const t = await r.text().catch(()=>String(r.status));
      throw new Error(`POST ${url} -> ${r.status} ${t}`);
    }
    return await r.json();
  }

  // ---------- Fordítás UI ----------
  let translateWrap=null, translateBtn=null, deBox=null, deCount=null, infoLine=null;
  let lastSrcText = "";

  function ensureTranslateUI(){
    if (!besch || translateWrap) return;
    const section = besch.closest(".card") || besch.parentElement || document.body;
    const div = document.createElement("div");
    div.id = "translate-ui";
    div.style.marginTop = "0.75rem";
    div.style.display = "none";
    div.innerHTML = `
      <div style="display:flex;align-items:center;gap:.5rem;flex-wrap:wrap;">
        <button type="button" id="btn-translate-de" class="btn secondary">Prevedi na njemački</button>
        <span class="muted small" id="translate-info"></span>
      </div>
      <div id="translate-result" style="margin-top:.5rem;display:none;">
        <label class="muted small" style="display:block;margin-bottom:.25rem;">Deutsch (bearbeitbar)</label>
        <textarea id="beschreibung_de" rows="5" style="width:100%;"></textarea>
        <div class="muted small" id="beschreibung_de_count" style="margin-top:.25rem;">0</div>
      </div>`;
    section.appendChild(div);

    translateWrap = div;
    translateBtn  = div.querySelector("#btn-translate-de");
    deBox         = div.querySelector("#beschreibung_de");
    deCount       = div.querySelector("#beschreibung_de_count");
    infoLine      = div.querySelector("#translate-info");

    translateBtn.addEventListener("click", doTranslate);
    deBox.addEventListener("input", updateDeCount);
    deBox.addEventListener("change", updateDeCount);
  }

  function toggleTranslateUI(lang){
    ensureTranslateUI();
    if (!translateWrap) return;
    translateWrap.style.display = (lang === "hr" ? "block" : "none");
    if (translateBtn) translateBtn.textContent = (lang === "hr") ? "Prevedi na njemački" : "Übersetzen";
  }

  function updateDeCount(){
    if (!deBox || !deCount) return;
    deCount.textContent = String((deBox.value || "").length);
  }

  async function doTranslate(){
    const text = (besch?.value || "").trim();        // <— FIX: besch
    if (!text){ alert("Nincs lefordítható szöveg."); return; }
    lastSrcText = text;

    const old = translateBtn.textContent;
    translateBtn.disabled = true;
    translateBtn.textContent = (currentLang === "hr") ? "Prevođenje…" : "Übersetze…";
    try{
      const { translated } = await postJSON("/api/translate", {
        text, source: "hr", target: "de"
      });
      $("#translate-result").style.display = "block";
      deBox.value = translated || "";
      updateDeCount();
      infoLine.textContent = "Azure fordítás kész.";
      deBox.focus();
    }catch(e){
      console.error(e);
      alert("A fordítás most nem sikerült. Ellenőrizd az Azure beállításokat a szerveren.");
    }finally{
      translateBtn.disabled = false;
      translateBtn.textContent = old;
    }
  }

  // indulás
  ensureTranslateUI();
  toggleTranslateUI(currentLang);

  // ---------- Beschreibung számláló ----------
  (function(){
    const out = document.getElementById("besch-count");
    if (!besch || !out) return;                      // <— FIX: besch
    const max = parseInt(besch.getAttribute("maxlength") || "1000", 10);
    function upd(){ out.textContent = `${besch.value.length} / ${max}`; }
    upd(); besch.addEventListener("input", upd); besch.addEventListener("change", upd);
  })();

  // ---------- Autocomplete (API, fallback CSV) ----------
  let WORKERS = [];
  const byAusweis = new Map();
  const byFullName = new Map();
  const norm = (s)=> (s||"").toString().trim();
  const keyName = (nach,vor)=> `${norm(nach).toLowerCase()}|${norm(vor).toLowerCase()}`;

  async function tryApiWorkers(q){
    const url = q ? `/api/workers?q=${encodeURIComponent(q)}` : `/api/workers`;
    const r = await fetch(url, { cache:"no-store" });
    if (!r.ok) throw new Error("no api");
    const js = await r.json();
    return Array.isArray(js.items) ? js.items : [];
  }

  async function loadWorkers(){
    try{
      const items = await tryApiWorkers("a"); // melegítés
      if (items.length){
        WORKERS = items.map(x => ({
          nachname: norm(x.last_name || x.nachname || x.name),
          vorname:  norm(x.first_name || x.vorname),
          ausweis:  norm(x.id || x.ausweis)
        }));
        indexWorkers();
        refreshAllAutocompletes();
        return;
      }
    }catch(_){}
    try{
      const resp = await fetch("/static/workers.csv", { cache:"no-store" });
      if (!resp.ok) return;
      let text = await resp.text();
      parseCSV(text);
    }catch(_){}
  }

  function parseCSV(text){
    if (text && text.charCodeAt(0)===0xFEFF) text = text.slice(1);
    const lines = text.split(/\r?\n/).filter(l=>l.trim().length);
    if (!lines.length) return;
    const delim = ((l)=> ( (l.match(/;/g)||[]).length > (l.match(/,/g)||[]).length ? ";" : "," ))(lines[0]);
    const split = (line)=>{
      const re = new RegExp(`${delim}(?=(?:[^"]*"[^"]*")*[^"]*$)`, "g");
      return line.split(re).map(x=>{
        x=x.trim(); if(x.startsWith('"')&&x.endsWith('"')) x=x.slice(1,-1); return x.replace(/""/g,'"');
      });
    };
    const header = split(lines[0]).map(h=>h.toLowerCase().trim());
    const iNach = header.findIndex(h=> h==="nachname" || h==="name");
    const iVor  = header.findIndex(h=> h==="vorname");
    const iAus  = header.findIndex(h=> /(ausweis|kennzeichen)/.test(h));
    if (iNach<0 || iVor<0 || iAus<0) return;

    WORKERS = [];
    for (let i=1;i<lines.length;i++){
      const c = split(lines[i]);
      if (c.length <= Math.max(iNach,iVor,iAus)) continue;
      const w = { nachname:norm(c[iNach]), vorname:norm(c[iVor]), ausweis:norm(c[iAus]) };
      if (!w.nachname && !w.vorname && !w.ausweis) continue;
      WORKERS.push(w);
    }
    indexWorkers();
    refreshAllAutocompletes();
  }

  function indexWorkers(){
    byAusweis.clear(); byFullName.clear();
    for (const w of WORKERS){
      if (w.ausweis) byAusweis.set(w.ausweis, w);
      byFullName.set(keyName(w.nachname, w.vorname), w);
    }
  }

  function makeAutocomplete(input, getOptions, onPick){
    const dd = document.createElement("div");
    dd.style.position="absolute"; dd.style.zIndex="99999"; dd.style.background="white";
    dd.style.border="1px solid #ddd"; dd.style.borderTop="none"; dd.style.maxHeight="220px";
    dd.style.overflowY="auto"; dd.style.display="none"; dd.style.boxShadow="0 6px 14px rgba(0,0,0,.08)";
    dd.style.borderRadius="0 0 .5rem .5rem"; dd.style.fontSize="14px"; dd.setAttribute("role","listbox");
    document.body.appendChild(dd);

    function hide(){ dd.style.display="none"; }
    function show(){ dd.style.display = dd.children.length ? "block" : "none"; }
    function position(){
      const r = input.getBoundingClientRect();
      dd.style.left=`${window.scrollX + r.left}px`;
      dd.style.top =`${window.scrollY + r.bottom}px`;
      dd.style.width=`${r.width}px`;
    }
    function render(list){
      dd.innerHTML="";
      list.slice(0,12).forEach((item)=>{
        const opt=document.createElement("div");
        opt.textContent=item.label ?? item.value;
        opt.dataset.value=item.value ?? item.label ?? "";
        opt.style.padding=".5rem .75rem"; opt.style.cursor="pointer";
        opt.addEventListener("mousedown",(e)=>{
          e.preventDefault();
          input.value = opt.dataset.value;
          hide(); onPick?.(opt.dataset.value, item);
          input.dispatchEvent(new Event("change",{bubbles:true}));
        });
        opt.addEventListener("mouseover",()=>opt.style.background="#f5f5f5"));
        opt.addEventListener("mouseout", ()=>opt.style.background="white");
        dd.appendChild(opt);
      });
      position(); show();
    }
    function filter(){
      const q=(input.value||"").toLowerCase().trim();
      if(!q){ hide(); return; }
      const raw=getOptions();
      const list=raw.filter(v=>(v.label??v).toLowerCase().includes(q))
                    .map(v=>(typeof v==="string"?{value:v}:v));
      render(list);
    }
    input.addEventListener("input", filter);
    input.addEventListener("focus", ()=>{ position(); filter(); });
    input.addEventListener("blur", ()=> setTimeout(hide,120));
    window.addEventListener("scroll", position, true);
    window.addEventListener("resize", position);
  }

  function setupAutocomplete(fs){
    const idx = fs.getAttribute("data-index");
    const inpNach = fs.querySelector(`input[name="nachname${idx}"]`);
    const inpVor  = fs.querySelector(`input[name="vorname${idx}"]`);
    const inpAus  = fs.querySelector(`input[name="ausweis${idx}"]`);
    if (!inpNach || !inpVor || !inpAus) return;

    makeAutocomplete(
      inpVor,
      () => WORKERS.map(w=>({ value:w.vorname, label:`${w.vorname} — ${w.nachname} [${w.ausweis}]`, payload:w })),
      (_, item)=>{
        if (item?.payload){
          const w=item.payload; inpVor.value=w.vorname; inpNach.value=w.nachname; inpAus.value=w.ausweis; return;
        }
        const w2 = byFullName.get(keyName(inpNach.value, inpVor.value));
        if (w2) inpAus.value=w2.ausweis;
      }
    );
    makeAutocomplete(
      inpNach,
      () => WORKERS.map(w=>({ value:w.nachname, label:`${w.nachname} — ${w.vorname} [${w.ausweis}]`, payload:w })),
      (_, item)=>{
        if (item?.payload){
          const w=item.payload; inpNach.value=w.nachname; inpVor.value=w.vorname; inpAus.value=w.ausweis; return;
        }
        const w2 = byFullName.get(keyName(inpNach.value, inpVor.value));
        if (w2) inpAus.value=w2.ausweis;
      }
    );
    makeAutocomplete(
      inpAus,
      () => WORKERS.map(w=>({ value:w.ausweis, label:`${w.ausweis} — ${w.nachname} ${w.vorname}`})),
      (value)=>{
        const w=byAusweis.get(value); if (w){ inpNach.value=w.nachname; inpVor.value=w.vorname; }
      }
    );
  }

  function refreshAllAutocompletes(){
    document.querySelectorAll("#worker-list .worker").forEach(setupAutocomplete);
  }

  // ---------- time picker + órák ----------
  const workerList = document.getElementById("worker-list");
  const addBtn     = document.getElementById("add-worker");
  const totalOut   = document.getElementById("gesamtstunden_auto");
  const breakHalf  = document.getElementById("break_half");
  const breakHidden= document.getElementById("break_minutes");
  const MAX_WORKERS=5;

  function enhanceTimePicker(inp){
    if (!inp || inp.dataset.enhanced==="1") return;
    inp.dataset.enhanced="1"; inp.type="hidden";
    const box=document.createElement("div"); box.style.display="flex"; box.style.gap=".5rem"; box.style.alignItems="center";
    const selH=document.createElement("select"); selH.style.minWidth="5.2rem"; selH.add(new Option("--","")); for(let h=0;h<24;h++) selH.add(new Option(String(h).padStart(2,"0"), String(h).padStart(2,"0")));
    const selM=document.createElement("select"); selM.style.minWidth="5.2rem"; selM.add(new Option("--","")); ["00","15","30","45"].forEach(m=>selM.add(new Option(m,m)));
    inp.insertAdjacentElement("afterend", box); box.appendChild(selH); box.appendChild(selM);
    function dispatch(){ inp.dispatchEvent(new Event("input",{bubbles:true})); inp.dispatchEvent(new Event("change",{bubbles:true})); }
    function compose(){ const h=selH.value, m=selM.value; const prev=inp.value; inp.value=(h&&m)?`${h}:${m}`:""; if (inp.value!==prev) dispatch(); }
    inp._setFromValue=(v)=>{ const mm=/^\d{2}:\d{2}$/.test(v)?v.split(":"):["",""]; selH.value=mm[0]||""; const allowed=["00","15","30","45"]; selM.value=allowed.includes(mm[1]||"")?(mm[1]||""):""; compose(); };
    inp._setFromValue(inp.value||""); selH.addEventListener("change", compose); selM.addEventListener("change", compose);
  }

  const toTime=(s)=>{ if(!s||!/^\d{2}:\d{2}$/.test(s)) return null; const [hh,mm]=s.split(":").map(Number); if(hh<0||hh>23||mm<0||mm>59) return null; return {hh,mm}; };
  const minutes=(t)=>t.hh*60+t.mm;
  const overlap=(a1,a2,b1,b2)=>Math.max(0, Math.min(a2,b2)-Math.max(a1,b1));
  function hoursWithBreaks(begStr,endStr){
    const bt=toTime(begStr), et=toTime(endStr); if(!bt||!et) return 0;
    const b=minutes(bt), e=minutes(et); if(e<=b) return 0;
    let total=e-b;
    const bm=parseInt((breakHidden?.value||"60"),10);
    if (bm===30) total=Math.max(0,total-30);
    else total=Math.max(0, total - (overlap(b,e,9*60,9*60+15)+overlap(b,e,12*60,12*60+45)));
    return Math.max(0,total)/60;
  }
  const formatHours=(h)=>(Math.round(h*100)/100).toFixed(2);

  function recalcWorker(wEl){
    const beg=wEl.querySelector('input[name^="beginn"]')?.value||"";
    const end=wEl.querySelector('input[name^="ende"]')?.value||"";
    const out=wEl.querySelector(".stunden-display");
    const h=hoursWithBreaks(beg,end);
    if(out) out.value = h?formatHours(h):"";
    return h;
  }
  function recalcAll(){
    let sum=0; workerList.querySelectorAll(".worker").forEach(w=>{ sum+=recalcWorker(w); });
    if (totalOut) totalOut.value = sum?formatHours(sum):"";
  }

  function enforceNumericKeyboard(input){
    input.setAttribute("inputmode","numeric");
    input.setAttribute("pattern","[0-9]*");
    input.addEventListener("input",()=>{ input.value=input.value.replace(/\D/g,""); });
  }

  function markSynced(inp,on){ if(!inp) return; on ? (inp.dataset.synced="1") : delete inp.dataset.synced; }
  function isSynced(inp){ return !!(inp && inp.dataset.synced==="1"); }
  function setupManualEditUnsync(inp){
    if (!inp) return;
    const unsync = ()=>markSynced(inp,false);
    inp.addEventListener("input",(e)=>{ if(e.isTrusted) unsync(); });
    inp.addEventListener("change",(e)=>{ if(e.isTrusted) unsync(); });
  }

  function setupWorker(fs){
    const idx=fs.getAttribute("data-index");
    setupAutocomplete(fs);
    const ausweis=fs.querySelector(`input[name="ausweis${idx}"]`); if(ausweis) enforceNumericKeyboard(ausweis);

    ["beginn","ende"].forEach(prefix=>{
      const inp=fs.querySelector(`input[name^="${prefix}"]`);
      if (inp){ enhanceTimePicker(inp); inp.addEventListener("input",recalcAll); inp.addEventListener("change",recalcAll); }
    });

    const b=fs.querySelector(`input[name="beginn${idx}"]`);
    const e=fs.querySelector(`input[name="ende${idx}"]`);

    if (idx==="1"){
      markSynced(b,false); markSynced(e,false);
      const syncFromFirst=()=>{
        const firstBeg=document.querySelector('input[name="beginn1"]')?.value||"";
        const firstEnd=document.querySelector('input[name="ende1"]')?.value||"";
        if(!firstBeg && !firstEnd) return;
        Array.from(workerList.querySelectorAll(".worker")).slice(1).forEach(fs2=>{
          const bb=fs2.querySelector('input[name^="beginn"]');
          const ee=fs2.querySelector('input[name^="ende"]');
          if(bb && (isSynced(bb)||!bb.value)){ bb.value=firstBeg; bb._setFromValue?.(firstBeg); markSynced(bb,true); }
          if(ee && (isSynced(ee)||!ee.value)){ ee.value=firstEnd; ee._setFromValue?.(firstEnd); markSynced(ee,true); }
        });
        recalcAll();
      };
      b?.addEventListener("input",syncFromFirst);
      e?.addEventListener("input",syncFromFirst);
      b?.addEventListener("change",syncFromFirst);
      e?.addEventListener("change",syncFromFirst);
    } else {
      setupManualEditUnsync(b); setupManualEditUnsync(e);
    }
  }

  const workerList = document.getElementById("worker-list");
  const addBtn     = document.getElementById("add-worker");
  const totalOut   = document.getElementById("gesamtstunden_auto");
  const breakHalf  = document.getElementById("break_half");
  const breakHidden= document.getElementById("break_minutes");
  const MAX_WORKERS=5;

  const firstWorker = workerList?.querySelector(".worker");
  if (firstWorker){ setupWorker(firstWorker); recalcAll(); }

  addBtn?.addEventListener("click", ()=>{
    const current=workerList.querySelectorAll(".worker").length;
    if (current>=MAX_WORKERS) return;
    const idx=current+1;
    const tpl=document.createElement("fieldset");
    tpl.className="worker"; tpl.dataset.index=String(idx);
    tpl.innerHTML=`
      <legend data-i18n="mitarbeiter">Mitarbeiter</legend> ${idx}
      <div class="grid-3">
        <div class="field"><label data-i18n="vorname">Vorname</label><input name="vorname${idx}" type="text" /></div>
        <div class="field"><label data-i18n="nachname">Nachname</label><input name="nachname${idx}" type="text" /></div>
        <div class="field"><label data-i18n="ausweis">Ausweis-Nr. / Kennzeichen</label><input name="ausweis${idx}" type="text" /></div>
      </div>
      <div class="grid"><div class="field"><label data-i18n="vorhaltung">Vorhaltung / beauftragtes Gerät / Fahrzeug</label><input name="vorhaltung${idx}" type="text" /></div></div>
      <div class="grid-3">
        <div class="field"><label data-i18n="beginn">Beginn</label><input name="beginn${idx}" type="time" /></div>
        <div class="field"><label data-i18n="ende">Ende</label><input name="ende${idx}" type="time" /></div>
        <div class="field"><label data-i18n="stunden">Stunden (auto)</label><input class="stunden-display" type="text" value="" readonly /></div>
      </div>`;
    workerList.appendChild(tpl);
    applyTranslations(currentLang);
    const firstBeg=document.querySelector('input[name="beginn1"]')?.value||"";
    const firstEnd=document.querySelector('input[name="ende1"]')?.value||"";
    const begNew=tpl.querySelector(`input[name="beginn${idx}"]`);
    const endNew=tpl.querySelector(`input[name="ende${idx}"]`);
    if(firstBeg && begNew) begNew.value=firstBeg;
    if(firstEnd && endNew) endNew.value=firstEnd;
    setupWorker(tpl);
    begNew?._setFromValue?.(begNew.value||""); endNew?._setFromValue?.(endNew.value||"");
    if(begNew && begNew.value) markSynced(begNew,true);
    if(endNew && endNew.value) markSynced(endNew,true);
    recalcAll();
  });

  // dolgozók betöltése
  loadWorkers();

  // ---------- Break switch ----------
  function updateBreak(){ if (breakHidden) breakHidden.value = breakHalf?.checked ? "30":"60"; recalcAll(); }
  breakHalf?.addEventListener("change", updateBreak); updateBreak();

  // ---------- Validáció / beküldés ----------
  (function(){
    if (!form) return;
    const trim=(v)=> (v||"").toString().trim();
    const readWorker=(fs)=>{
      const idx=fs.getAttribute("data-index")||"";
      const q=(s)=> fs.querySelector(s);
      return {
        idx,
        vorname: trim(q(`input[name="vorname${idx}"]`)?.value),
        nachname:trim(q(`input[name="nachname${idx}"]`)?.value),
        ausweis: trim(q(`input[name="ausweis${idx}"]`)?.value),
        beginn:  trim(q(`input[name="beginn${idx}"]`)?.value),
        ende:    trim(q(`input[name="ende${idx}"]`)?.value),
      };
    };
    form.addEventListener("submit", (e)=>{
      if (e.submitter && e.submitter.getAttribute("formaction")==="/generate_pdf") return;
      const errors=[];
      const datum = trim(document.getElementById("datum")?.value);
      const bau   = trim(document.getElementById("bau")?.value);
      const bf    = trim(document.getElementById("basf_beauftragter")?.value);
      const beschTxt = trim(besch?.value); // <— FIX: besch
      if(!datum) errors.push("Bitte das Datum der Leistungsausführung angeben.");
      if(!bau)   errors.push("Bitte Bau und Ausführungsort ausfüllen.");
      if(!bf)    errors.push("Bitte den BASF-Beauftragten (Org.-Code) ausfüllen.");
      if(!beschTxt) errors.push("Bitte die Beschreibung der ausgeführten Arbeiten ausfüllen.");

      const sets = Array.from(document.getElementById("worker-list").querySelectorAll(".worker"));
      let validWorkers=0;
      sets.forEach((fs)=>{
        const w=readWorker(fs);
        const any=!!(w.vorname||w.nachname||w.ausweis||w.beginn||w.ende);
        const all=!!(w.vorname&&w.nachname&&w.ausweis&&w.beginn&&w.ende);
        if (all) validWorkers+=1;
        else if (any){
          const missing=[];
          if(!w.vorname) missing.push("Vorname");
          if(!w.nachname)missing.push("Nachname");
          if(!w.ausweis) missing.push("Ausweis-Nr.");
          if(!w.beginn)  missing.push("Beginn");
          if(!w.ende)    missing.push("Ende");
          errors.push(`Bitte Mitarbeiter ${w.idx}: ${missing.join(", ")} ausfüllen.`);
        }
      });
      if (errors.length){ e.preventDefault(); alert(errors.join("\n")); return false; }
      return true;
    }, false);
  })();

  // ---------- Excel letöltés ----------
  (function(){
    if (!form) return;
    const submitBtn = form.querySelector('button[type="submit"]:not([formaction])');
    function setBusy(b){
      if(!submitBtn) return;
      if(b){ submitBtn.dataset._label=submitBtn.textContent||"Generieren"; submitBtn.disabled=true; submitBtn.textContent=(window.__GEN_TEXT||"Wird generiert..."); }
      else { submitBtn.disabled=false; if(submitBtn.dataset._label) submitBtn.textContent=submitBtn.dataset._label; }
    }
    function injectTranslation(fd){
      const deTxt=(document.getElementById("beschreibung_de")?.value||"").trim();
      const srcTxt=(lastSrcText||"").trim();
      if (deTxt){ fd.set("beschreibung_src", srcTxt || (besch?.value||"")); fd.set("beschreibung_de_used","1"); fd.set("beschreibung", deTxt); } // <— FIX: besch
      else fd.set("beschreibung_de_used","0");
    }
    form.addEventListener("submit", async (e)=>{
      if (e.submitter && e.submitter.getAttribute("formaction")==="/generate_pdf") return;
      if (e.defaultPrevented) return;
      e.preventDefault();
      const fd = new FormData(form);
      injectTranslation(fd);
      setBusy(true);
      try{
        const res = await fetch("/generate_excel", { method:"POST", body: fd });
        if (!res.ok) throw new Error("HTTP "+res.status);
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a=document.createElement("a"); a.href=url; a.download="leistungsnachweis.xlsx";
        document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url), 2000);
      }catch(err){ console.error(err); alert("A fájl generálása most nem sikerült. Próbáld újra."); }
      finally{ setBusy(false); }
    });
  })();

  // ---------- PDF preview ----------
  (function(){
    if (!form) return;
    const pdfBtn = form.querySelector('button[formaction="/generate_pdf"]');
    if (!pdfBtn) return;
    function injectTranslation(fd){
      const deTxt=(document.getElementById("beschreibung_de")?.value||"").trim();
      const srcTxt=(lastSrcText||"").trim();
      if (deTxt){ fd.set("beschreibung_src", srcTxt || (besch?.value||"")); fd.set("beschreibung_de_used","1"); fd.set("beschreibung", deTxt); } // <— FIX: besch
      else fd.set("beschreibung_de_used","0");
    }
    pdfBtn.addEventListener("click", async (ev)=>{
      ev.preventDefault(); ev.stopPropagation();
      const win = window.open("", "_blank");
      if (!win){ alert("A böngésző letiltotta a felugró ablakot. Engedélyezd, kérlek."); return; }
      win.document.open();
      win.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>PDF Vorschau</title>
        <style>html,body{height:100%;margin:0}.box{font:15px/1.4 system-ui; padding:24px}</style></head>
        <body><div class="box">PDF wird generiert…</div></body></html>`);
      win.document.close();
      try{
        const fd=new FormData(form); injectTranslation(fd);
        const res=await fetch("/generate_pdf",{method:"POST",body:fd});
        if(!res.ok) throw new Error("HTTP "+res.status);
        const ct=res.headers.get("content-type")||""; if(!ct.includes("pdf")) throw new Error("Unerwartete Antwort – kein PDF.");
        const blob=await res.blob(); const url=URL.createObjectURL(blob);
        win.document.open();
        win.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>Vorschau</title>
          <style>html,body,iframe{height:100%;width:100%;margin:0;border:0}</style></head>
          <body><iframe src="${url}" frameborder="0"></iframe></body></html>`);
        win.document.close(); setTimeout(()=>URL.revokeObjectURL(url), 120000);
      }catch(err){
        win.document.open();
        win.document.write(`<pre style="white-space:pre-wrap;padding:24px">PDF-Erzeugung fehlgeschlagen:\n${(err && err.message)||err}</pre>`);
        win.document.close();
      }
    });
  })();

}); // DOMContentLoaded
