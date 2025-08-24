// ==== Helpers ====
const pad2 = (n) => String(n).padStart(2, "0");
const parseHHMM = (s) => {
  if (!s) return null;
  const [hh, mm] = s.split(":").map((x) => parseInt(x, 10));
  if (Number.isNaN(hh) || Number.isNaN(mm)) return null;
  return { hh, mm };
};
const minutesBetween = (a, b) => (b.hh * 60 + b.mm) - (a.hh * 60 + a.mm);

// hivatalos szabály: ha 60-as szünet, akkor 9:00–9:15 és 12:00–12:45 kerül levonásra az átfedés mértékében.
// ha 30-as szünet, akkor fix 30 perc (de max a teljes idő).
function hoursWithBreaks(begStr, endStr, pauseMin) {
  const A = parseHHMM(begStr), B = parseHHMM(endStr);
  if (!A || !B) return 0;
  let total = minutesBetween(A, B);
  if (total <= 0) return 0;

  if (pauseMin >= 60) {
    const overlap = (x1, x2, y1, y2) => {
      const s = Math.max(x1, y1), e = Math.min(x2, y2);
      return Math.max(0, e - s);
    };
    const a1 = A.hh * 60 + A.mm, a2 = B.hh * 60 + B.mm;
    const m1 = overlap(a1, a2, 9 * 60 + 0, 9 * 60 + 15);   // 09:00–09:15
    const m2 = overlap(a1, a2, 12 * 60 + 0, 12 * 60 + 45); // 12:00–12:45
    total = Math.max(0, total - (m1 + m2));
  } else {
    total = Math.max(0, total - Math.min(total, 30));
  }
  return +(total / 60).toFixed(2);
}

// deduplikálás, kis/nagybetűtől függetlenül
const uniqCaseInsensitive = (arr) => {
  const seen = new Set();
  const out = [];
  for (const v of arr) {
    const k = (v || "").trim().toLowerCase();
    if (!k || seen.has(k)) continue;
    seen.add(k);
    out.push(v);
  }
  return out;
};

// ==== Állapot az autocomplete-hez ====
let WORKERS = []; // {vorname, nachname, ausweis}

// CSV beolvasás – próbáljuk /static/workers.csv majd /workers.csv
async function loadWorkersCSV() {
  const tryUrls = ["/static/workers.csv", "/workers.csv"];
  let text = "";
  for (const url of tryUrls) {
    try {
      const res = await fetch(url, { cache: "no-store" });
      if (res.ok) {
        text = await res.text();
        break;
      }
    } catch (_) {}
  }
  if (!text) return;

  // vessző vagy pontosvessző elválasztó támogatása
  const lines = text.split(/\r?\n/).filter((l) => l.trim().length);
  // fejléc felismerés
  let header = lines[0].split(/[;,]/).map((h) => h.trim().toLowerCase());
  let startIndex = 1;
  const col = {
    vorname: header.findIndex((h) => h.includes("vorname") || h.includes("kereszt")),
    nachname: header.findIndex((h) => h.includes("nachname") || h.includes("vezetéknév") || h.includes("család")),
    ausweis: header.findIndex((h) => h.includes("ausweis") || h.includes("kenn") || h.includes("id")),
  };
  // ha nincs fejléc, tegyük fel az oszlopok sorrendjét: vorname;nachname;ausweis
  if (col.vorname < 0 || col.nachname < 0 || col.ausweis < 0) {
    header = null;
    startIndex = 0;
    col.vorname = 0; col.nachname = 1; col.ausweis = 2;
  }

  const rows = lines.slice(startIndex);
  const items = [];
  for (const line of rows) {
    const parts = line.split(/[;,]/);
    const v = (parts[col.vorname] || "").trim();
    const n = (parts[col.nachname] || "").trim();
    const a = (parts[col.ausweis] || "").trim();
    if (v || n || a) items.push({ vorname: v, nachname: n, ausweis: a });
  }
  WORKERS = items;

  // datalist-ek feltöltése
  const dlV = document.getElementById("dl-vorname");
  const dlN = document.getElementById("dl-nachname");
  const dlA = document.getElementById("dl-ausweis");
  const vset = uniqCaseInsensitive(items.map((x) => x.vorname));
  const nset = uniqCaseInsensitive(items.map((x) => x.nachname));
  const aset = uniqCaseInsensitive(items.map((x) => x.ausweis));

  dlV.innerHTML = vset.map((v) => `<option value="${escapeHtml(v)}"></option>`).join("");
  dlN.innerHTML = nset.map((v) => `<option value="${escapeHtml(v)}"></option>`).join("");
  dlA.innerHTML = aset.map((v) => `<option value="${escapeHtml(v)}"></option>`).join("");
}

function escapeHtml(s) {
  return String(s || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

// „Okos” összekötés: ha Ausweis egyedi, töltsük ki a neveket;
// ha Vorname+Nachname kombináció egyedi, töltsük ki az Ausweis-t.
function wireSmartFill(workerFieldset) {
  const inpV = workerFieldset.querySelector(".inp-vorname");
  const inpN = workerFieldset.querySelector(".inp-nachname");
  const inpA = workerFieldset.querySelector(".inp-ausweis");

  function tryFillFromAusweis() {
    const a = (inpA.value || "").trim().toLowerCase();
    if (!a) return;
    const matches = WORKERS.filter(w => (w.ausweis || "").trim().toLowerCase() === a);
    if (matches.length === 1) {
      if (!inpV.value) inpV.value = matches[0].vorname || "";
      if (!inpN.value) inpN.value = matches[0].nachname || "";
    }
  }

  function tryFillFromNames() {
    const v = (inpV.value || "").trim().toLowerCase();
    const n = (inpN.value || "").trim().toLowerCase();
    if (!v || !n) return;
    const matches = WORKERS.filter(w =>
      (w.vorname || "").trim().toLowerCase() === v &&
      (w.nachname || "").trim().toLowerCase() === n
    );
    if (matches.length === 1) {
      if (!inpA.value) inpA.value = matches[0].ausweis || "";
    }
  }

  inpA.addEventListener("change", tryFillFromAusweis);
  inpA.addEventListener("blur", tryFillFromAusweis);
  inpV.addEventListener("change", tryFillFromNames);
  inpN.addEventListener("change", tryFillFromNames);
  inpV.addEventListener("blur", tryFillFromNames);
  inpN.addEventListener("blur", tryFillFromNames);
}

// 15 perces lépés garantálása dinamikusan hozzáadott mezőknek is
function enforce15MinStep(scope) {
  scope.querySelectorAll('input[type="time"]').forEach((el) => {
    el.setAttribute("step", "900");
  });
}

// óraszám élő számítás és összesítés
function wireHoursCalc(workerFieldset) {
  const beg = workerFieldset.querySelector(".t-beginn");
  const end = workerFieldset.querySelector(".t-ende");
  const out = workerFieldset.querySelector(".stunden-display");
  const ckHalf = workerFieldset.querySelector(".ck-break-half");
  const hiddenBreak = workerFieldset.querySelector(".break-minutes");

  const recalc = () => {
    hiddenBreak.value = ckHalf.checked ? "30" : "60";
    const h = hoursWithBreaks(beg.value, end.value, parseInt(hiddenBreak.value, 10));
    out.value = h ? h.toFixed(2) : "";
    recalcTotal();
  };

  ckHalf.addEventListener("change", recalc);
  beg.addEventListener("change", recalc);
  end.addEventListener("change", recalc);
  beg.addEventListener("input", recalc);
  end.addEventListener("input", recalc);
}

function recalcTotal() {
  let total = 0;
  document.querySelectorAll(".worker").forEach((w) => {
    const val = w.querySelector(".stunden-display").value;
    if (val) total += parseFloat(val);
  });
  const g = document.getElementById("gesamtstunden_auto");
  g.value = total ? total.toFixed(2) : "";
}

// új munkavállaló blokk hozzáadása (max 5)
function addWorker() {
  const list = document.getElementById("worker-list");
  const count = list.querySelectorAll(".worker").length;
  if (count >= 5) return;

  const idx = count + 1;
  const tmpl = list.querySelector(".worker");
  const clone = tmpl.cloneNode(true);

  clone.dataset.index = idx;
  clone.querySelector("legend").textContent = `Mitarbeiter ${idx}`;

  // nevezd át a name attribútumokat sorszámozva
  clone.querySelectorAll("input, textarea, select").forEach((el) => {
    if (el.name) {
      el.name = el.name.replace(/\d+$/, "") + idx;
    }
    if (el.classList.contains("stunden-display")) {
      el.value = "";
    }
    if (el.type === "checkbox") {
      el.checked = false;
    }
    if (el.type === "hidden" && el.classList.contains("break-minutes")) {
      el.value = "60";
    }
    if (el.type === "time") {
      el.value = "";
    }
    if (el.type === "text" && !el.classList.contains("stunden-display")) {
      el.value = "";
    }
  });

  list.appendChild(clone);
  enforce15MinStep(clone);
  wireHoursCalc(clone);
  wireSmartFill(clone);
}

// leírás karakter számláló
function wireDescrCounter() {
  const ta = document.getElementById("beschreibung");
  const counter = document.getElementById("besch-count");
  const update = () => {
    const n = (ta.value || "").length;
    counter.textContent = `${n} / 1000`;
  };
  ta.addEventListener("input", update);
  update();
}

// init
window.addEventListener("DOMContentLoaded", async () => {
  // garantáljuk a 15 perces lépést a meglévő mezőkön
  enforce15MinStep(document);

  // karakter számláló
  wireDescrCounter();

  // élő óraszámítás az első blokkon
  document.querySelectorAll(".worker").forEach((w) => {
    wireHoursCalc(w);
    wireSmartFill(w);
  });

  // plusz munkavállaló gomb
  document.getElementById("add-worker").addEventListener("click", addWorker);

  // workers.csv betöltés és datalist-ek feltöltése
  await loadWorkersCSV();
});
