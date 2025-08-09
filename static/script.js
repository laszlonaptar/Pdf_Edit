(function () {
  const workersDiv = document.getElementById("workers");
  const addBtn = document.getElementById("addWorker");
  const totalEl = document.getElementById("gesamt");

  const BREAKS = [
    { start: "09:00", end: "09:15" },
    { start: "12:00", end: "12:45" },
  ];

  function parse(t) {
    if (!t) return null;
    const s = t.replace(".", ":");
    const [H, M] = s.split(":").map(x => parseInt(x, 10));
    if (Number.isNaN(H) || Number.isNaN(M)) return null;
    return H * 60 + M;
    }

  function overlap(a0, a1, b0, b1) {
    return Math.max(0, Math.min(a1, b1) - Math.max(a0, b0));
  }

  function computeTotals() {
    let total = 0;
    document.querySelectorAll(".worker-row").forEach(row => {
      const b = parse(row.querySelector('input[name="beginn[]"]').value);
      const e = parse(row.querySelector('input[name="ende[]"]').value);
      if (b != null && e != null && e > b) {
        let mins = e - b;
        BREAKS.forEach(br => {
          const bs = parse(br.start), be = parse(br.end);
          mins -= overlap(b, e, bs, be);
        });
        total += Math.max(0, mins);
        row.querySelector(".row-total").textContent = (mins / 60).toFixed(2);
      } else {
        row.querySelector(".row-total").textContent = "";
      }
    });
    totalEl.value = (total / 60).toFixed(2);
  }

  function addWorker(prefill = {}) {
    const row = document.createElement("div");
    row.className = "worker-row";
    row.innerHTML = `
      <input type="text"   name="nachname[]" placeholder="Name" value="${prefill.nachname || ""}" />
      <input type="text"   name="vorname[]"  placeholder="Vorname" value="${prefill.vorname || ""}" />
      <input type="text"   name="ausweis[]"  placeholder="Ausweis-Nr." value="${prefill.ausweis || ""}" />
      <input type="time"   name="beginn[]"   value="${prefill.beginn || ""}" />
      <input type="time"   name="ende[]"     value="${prefill.ende || ""}" />
      <span class="row-total"></span>
      <button type="button" class="remove">✕</button>
    `;
    row.querySelectorAll('input[name="beginn[]"], input[name="ende[]"]').forEach(inp => {
      inp.addEventListener("change", computeTotals);
      inp.addEventListener("input", computeTotals);
    });
    row.querySelector(".remove").addEventListener("click", () => {
      row.remove();
      computeTotals();
    });
    workersDiv.appendChild(row);
  }

  // induláskor egy sor
  addWorker();

  addBtn.addEventListener("click", () => addWorker());
})();
