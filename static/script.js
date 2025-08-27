/* Teljes front logika:
   - DE/HR i18n
   - Fordítás (HR->DE) /api/translate (részletes hibajelzés)
   - Autocomplete: az első karakter után legördülő lista (Név – Vezetéknév – Ausweis),
     kattintásra kitölti a 3 mezőt adott sorban
   - Pauza: 60 alap, checkbox -> 30 (rejtett break_minutes mező)
   - Munkaidő: óra+perc select (perc: 00/15/30/45), rejtett HH:MM inputba szinkronizálva
   - PDF előnézet
   - Piszkozat mentés (localStorage)
*/
(function () {
  "use strict";
  const $ = (s, r) => (r || document).querySelector(s);
  const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
  const debounce = (fn, ms) => { let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), ms); }; };
  async function fetchJSON(url, opts = {}) {
    const r = await fetch(url, { headers: { "Accept": "application/json" }, ...opts });
    const ct = r.headers.get("content-type") || "";
    let body;
    try { body = ct.includes("application/json") ? await r.json() : await r.text(); }
    catch { body = null; }
    if (!r.ok) {
      const msg = (body && body.error) ? body.error : (typeof body === "string" ? body : (`HTTP ${r.status}`));
      throw new Error(msg);
    }
    return body;
  }
  const I18N = window.__i18n__ || null;
  const DICT = window.__i18nDict || {};
  const getLang = () => (I18N ? I18N.getLang() : "de");
  const T = (k) => (DICT[getLang()] && DICT[getLang()][k]) || k;

  // ---- Pauza switch
  function initBreakSwitch() {
    const chk = $("#half_break");
    const hid = $("#break_minutes");
    if (!chk || !hid) return;
    const apply = () => { hid.value = chk.checked ? "30" : "60"; };
    chk.addEventListener("change", apply);
    apply();
  }

  // ---- Time selects (óra + perc -> hidden HH:MM)
  function fillTimeSelects() {
    for (let i=1;i<=5;i++) {
      const hs = $(`#beginn${i}_h`), ms = $(`#beginn${i}_m`), hh = $(`#beginn${i}`);
      const he = $(`#ende${i}_h`), me = $(`#ende${i}_m`), eh = $(`#ende${i}`);
      const fill = (sel, from, to, step) => {
        sel.innerHTML = "";
        for (let v=from; v<=to; v+=step) {
          const opt = document.createElement("option");
          opt.value = String(v).padStart(2,"0");
          opt.textContent = opt.value;
          sel.appendChild(opt);
        }
      };
      if (hs && ms && hh) { fill(hs,0,23,1); fill(ms,0,45,15); hs.value="07"; ms.value="00"; hh.value="07:00"; }
      if (he && me && eh) { fill(he,0,23,1); fill(me,0,45,15); he.value="15"; me.value="00"; eh.value="15:00"; }

      const sync = (H, M, HIDDEN) => { HIDDEN.value = `${H.value}:${M.value}`; };
      if (hs && ms && hh) { hs.addEventListener("change",()=>sync(hs,ms,hh)); ms.addEventListener("change",()=>sync(hs,ms,hh)); }
      if (he && me && eh) { he.addEventListener("change",()=>sync(he,me,eh)); me.addEventListener("change",()=>sync(he,me,eh)); }
    }
  }

  // ---- Fordítás HR->DE
  function initTranslate() {
    const btn = $("#btn-translate-hr-de");
    const status = $("#translate-status");
    const src = $("#beschreibung");
    const dst = $("#beschreibung_de"); // ha nincs, src-t írjuk felül
    if (!btn || !src || !status) return;
    btn.addEventListener("click", async () => {
      const text = (src.value || "").trim();
      if (!text) return;
      btn.disabled = true;
      status.textContent = T("t_busy");
      status.className = "status info";
      try {
        const body = JSON.stringify({ text, source: "hr", target: "de" });
        const r = await fetch("/api/translate", { method: "POST", headers: { "Content-Type":"application/json" }, body });
        const ct = r.headers.get("content-type") || "";
        let j = null; try { j = ct.includes("application/json") ? await r.json() : null; } catch {}
        if (!r.ok || !j || !j.translated) {
          const msg = (j && j.error) ? j.error : `HTTP ${r.status}`;
          throw new Error(msg);
        }
        if (dst) dst.value = j.translated; else src.value = j.translated;
        status.textContent = T("t_done");
        status.className = "status ok";
      } catch (e) {
        status.textContent = (getLang()==="hr" ? "Greška pri prijevodu." : "Übersetzung fehlgeschlagen.") + (e && e.message ? ` — ${e.message}` : "");
        status.className = "status err";
      } finally {
        btn.disabled = false;
        setTimeout(()=>{ status.textContent=""; status.className="status"; }, 4000);
      }
    });
  }

  // ---- Autocomplete (választható lista az első karaktertől)
  const WorkerAC = (() => {
    let cacheAll = null, inflight = null;
    const uniqBy = (arr, key) => Array.from(new Map(arr.map(o => [o[key], o])).values());

    async function loadAllWorkers() {
      if (cacheAll) return cacheAll;
      if (inflight) return inflight;
      const tryFetch = async (url) => { try { const r = await fetchJSON(url); if (r && Array.isArray(r.items)) return r.items; } catch {} return null; };
      inflight = (async () => {
        let items = await tryFetch("/api/workers");
        if (!items) items = await tryFetch("/api/workers?q=");
        if (!items) items = await tryFetch("/workers");
        if (!items) items = await tryFetch("/workers?q=");
        if (!items) items = [];
        cacheAll = items;
        inflight = null;
        return cacheAll;
      })();
      return inflight;
    }

    function makeRowPicker(rowEl, inpLast, inpFirst, inpId) {
      const listLast = rowEl.querySelector(`#ac-last-${rowEl.dataset.row}`);
      const listFirst = rowEl.querySelector(`#ac-first-${rowEl.dataset.row}`);
      const listId = rowEl.querySelector(`#ac-id-${rowEl.dataset.row}`);
      const renderList = (host, items) => {
        host.innerHTML = "";
        if (!items.length) { host.style.display = "none"; return; }
        for (const w of items.slice(0, 20)) {
          const d = document.createElement("div");
          d.className = "ac-item";
          d.textContent = `${w.last_name || ""} ${w.first_name || ""} – ${w.id || ""}`.trim();
          d.addEventListener("mousedown", (ev) => { ev.preventDefault();
            if (inpLast) inpLast.value = w.last_name || "";
            if (inpFirst) inpFirst.value = w.first_name || "";
            if (inpId) inpId.value = w.id || "";
            host.style.display = "none";
          });
          host.appendChild(d);
        }
        host.style.display = "block";
      };
      const update = debounce(async () => {
        const q = (inpLast.value || inpFirst.value || inpId.value || "").toLowerCase().trim();
        const all = await loadAllWorkers();
        const items = !q ? all : all.filter(w =>
          (w.last_name||"").toLowerCase().includes(q) ||
          (w.first_name||"").toLowerCase().includes(q) ||
          (w.id||"").toLowerCase().includes(q)
        );
        const uniq = uniqBy(items, "id");
        renderList(listLast, uniq);
        renderList(listFirst, uniq);
        renderList(listId, uniq);
      }, 120);

      [inpLast, inpFirst, inpId].forEach(el => {
        if (!el) return;
        el.addEventListener("input", update);
        el.addEventListener("focus", update);
        el.addEventListener("blur", () => setTimeout(()=> {
          listLast.style.display="none"; listFirst.style.display="none"; listId.style.display="none";
        }, 180));
      });
    }

    async function init() {
      await loadAllWorkers();
      for (let i=1;i<=5;i++) {
        const rowEl = document.querySelector(`.worker-row[data-row="${i}"]`);
        if (!rowEl) continue;
        makeRowPicker(rowEl, $(`#nachname${i}`), $(`#vorname${i}`), $(`#ausweis${i}`));
      }
    }
    return { init };
  })();

  // ---- PDF
  function initPdf() {
    const btn = $("#btn-pdf");
    const form = $("#main-form");
    if (!btn || !form) return;
    btn.addEventListener("click", async () => {
      try{
        const fd = new FormData(form);
        const r = await fetch("/generate_pdf", { method: "POST", body: fd });
        if (!r.ok) throw new Error("pdf");
        const blob = await r.blob();
        const url = URL.createObjectURL(blob);
        window.open(url,"_blank");
        setTimeout(()=>URL.revokeObjectURL(url),15000);
      }catch{
        alert(getLang()==="hr" ? "PDF nije dostupan na poslužitelju." : "PDF előállítás nem elérhető ezen a szerveren.");
      }
    });
  }

  // ---- Piszkozat
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
    fillTimeSelects();
    initTranslate();
    await WorkerAC.init();
    initPdf();
    Draft.init();
  });
})();
