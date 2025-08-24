// static/script.js  (v14)

// --- Beschreibung karakter számláló ---
(function () {
  const ta = document.getElementById("beschreibung");
  const out = document.getElementById("besch-count");
  if (!ta || !out) return;
  const update = () => {
    const len = (ta.value || "").length;
    out.textContent = `${len} / ${ta.maxLength || 1000}`;
  };
  ta.addEventListener("input", update);
  update();
})();

// --- Idő -> óra számítás segédek ---
function parseHHMM(s) {
  if (!s || s.indexOf(":") < 0) return null;
  const [hh, mm] = s.split(":").map((x) => parseInt(x, 10));
  if (Number.isNaN(hh) || Number.isNaN(mm)) return null;
  return { hh, mm };
}

function toMinutes(t) {
  return t.hh * 60 + t.mm;
}

function overlapMinutes(a1, a2, b1, b2) {
  const start = Math.max(a1, b1);
  const end = Math.min(a2, b2);
  return Math.max(0, end - start);
}

function hoursWithBreaks(begStr, endStr, breakMin) {
  const beg = parseHHMM(begStr);
  const end = parseHHMM(endStr);
  if (!beg || !end) return 0;
  let total = toMinutes(end) - toMinutes(beg);
  if (total <= 0) return 0;

  let minus;
  if (breakMin >= 60) {
    // 09:00–09:15 és 12:00–12:45 automatikus levonás, ha belelóg
    const s = toMinutes(beg);
    const e = toMinutes(end);
    minus = 0;
    minus += overlapMinutes(s, e, 9 * 60 + 0, 9 * 60 + 15);
    minus += overlapMinutes(s, e, 12 * 60 + 0, 12 * 60 + 45);
  } else {
    minus = Math.min(total, 30);
  }
  return Math.max(0, (total - minus) / 60);
}

// --- Dinamikus dolgozó blokkok ---
(function () {
  const list = document.getElementById("worker-list");
  const addBtn = document.getElementById("add-worker");
  if (!list || !addBtn) return;

  const MAX_WORKERS = 5;

  function setTimeStep(fieldset) {
    // biztos, ami biztos: 15 perces lépés
    fieldset.querySelectorAll('input[type="time"]').forEach((el) => {
      el.setAttribute("step", "900");
    });
  }

  function recalcOne(fieldset) {
    const beg = fieldset.querySelector(".t-beginn")?.value || "";
    const end = fieldset.querySelector(".t-ende")?.value || "";
    const breakHidden = document.getElementById("break_minutes");
    const breakMin = breakHidden ? parseInt(breakHidden.value, 10) : 60;
    const h = hoursWithBreaks(beg, end, breakMin);
    const out = fieldset.querySelector(".stunden-display");
    if (out) out.value = h ? h.toFixed(2) : "";
  }

  function recalcAll() {
    let sum = 0;
    list.querySelectorAll(".worker").forEach((fs) => {
      const v = parseFloat(fs.querySelector(".stunden-display")?.value || "0");
      if (!Number.isNaN(v)) sum += v;
    });
    const total = document.getElementById("gesamtstunden_auto");
    if (total) total.value = sum ? sum.toFixed(2) : "";
  }

  function wireFieldset(fieldset) {
    setTimeStep(fieldset);
    fieldset.addEventListener("input", () => {
      recalcOne(fieldset);
      recalcAll();
    });
    recalcOne(fieldset);
    recalcAll();
  }

  // első blokk bekötése
  const first = list.querySelector(".worker");
  if (first) wireFieldset(first);

  // fél óra szünet checkbox
  const ck = document.getElementById("break_half");
  const breakHidden = document.getElementById("break_minutes");
  if (ck && breakHidden) {
    const syncBreak = () => {
      breakHidden.value = ck.checked ? "30" : "60";
      // váltáskor újraszámolunk mindent
      list.querySelectorAll(".worker").forEach(recalcOne);
      recalcAll();
    };
    ck.addEventListener("change", syncBreak);
    syncBreak();
  }

  function cloneWorker() {
    const count = list.querySelectorAll(".worker").length;
    if (count >= MAX_WORKERS) return;

    const tmpl = list.querySelector(".worker:last-of-type");
    const copy = tmpl.cloneNode(true);

    // új index
    const idx = count + 1;
    copy.setAttribute("data-index", String(idx));
    const legend = copy.querySelector("legend");
    if (legend) legend.textContent = `Mitarbeiter ${idx}`;

    // inputok átnevezése/ürítése
    copy.querySelectorAll("input").forEach((inp) => {
      if (inp.type === "hidden") return;
      if (inp.classList.contains("stunden-display")) {
        inp.value = "";
        return;
      }
      if (inp.classList.contains("t-beginn")) {
        inp.name = `beginn${idx}`;
        inp.value = "";
        inp.setAttribute("step", "900");
        return;
      }
      if (inp.classList.contains("t-ende")) {
        inp.name = `ende${idx}`;
        inp.value = "";
        inp.setAttribute("step", "900");
        return;
      }
      // szöveg típusok:
      if (inp.name.startsWith("vorname")) inp.name = `vorname${idx}`;
      if (inp.name.startsWith("nachname")) inp.name = `nachname${idx}`;
      if (inp.name.startsWith("ausweis")) inp.name = `ausweis${idx}`;
      if (inp.name.startsWith("vorhaltung")) inp.name = `vorhaltung${idx}`;
      if (inp.type === "text") inp.value = "";
    });

    list.appendChild(copy);
    wireFieldset(copy);
  }

  addBtn.addEventListener("click", cloneWorker);
})();
