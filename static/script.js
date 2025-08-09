// --- settings ---
const MAX_WORKERS = 5;

// fixed breaks used also on the backend
const BREAKS = [
  { from: "09:00", to: "09:15" },
  { from: "12:00", to: "12:45" }
];

const workersWrap = document.getElementById("workers");
const addBtn = document.getElementById("addWorker");
const totalEl = document.getElementById("gesamt");
const form = document.getElementById("lnForm");

// prefill today
(function(){
  const d = new Date();
  const iso = d.toISOString().slice(0,10); // yyyy-mm-dd
  document.getElementById("datum").value = iso;
})();

function timeToMin(t){
  const [h,m] = t.split(":").map(Number);
  return h*60+m;
}
function overlapMin(a1,a2,b1,b2){
  const left = Math.max(a1,b1);
  const right = Math.min(a2,b2);
  return Math.max(0, right-left);
}
function netHours(b, e){
  if(!b || !e) return 0;
  const s = timeToMin(b), f = timeToMin(e);
  if(f<=s) return 0;
  let mins = f - s;
  BREAKS.forEach(br=>{
    mins -= overlapMin(s,f,timeToMin(br.from),timeToMin(br.to));
  });
  return Math.max(0, Math.round((mins/60)*100)/100);
}

function workerTemplate(idx){
  return `
    <div class="worker" data-idx="${idx}">
      <div class="row">
        <label>Nachname
          <input name="nachname" required>
        </label>
        <label>Vorname
          <input name="vorname" required>
        </label>
        <label>Ausweis-Nr.
          <input name="ausweis" required>
        </label>
        <label>Beginn
          <input type="time" name="beginn" value="07:00" required>
        </label>
        <label>Ende
          <input type="time" name="ende" value="16:00" required>
        </label>
        <label>Stunden
          <input name="stunden" class="stunden" readonly>
        </label>
      </div>
    </div>
  `;
}

function recalc(){
  let total = 0;
  workersWrap.querySelectorAll(".worker").forEach(w=>{
    const b = w.querySelector('input[name="beginn"]').value;
    const e = w.querySelector('input[name="ende"]').value;
    const h = netHours(b,e);
    w.querySelector('.stunden').value = h.toFixed(2);
    total += h;
  });
  totalEl.value = total.toFixed(2);
}

function addWorker(){
  const count = workersWrap.querySelectorAll(".worker").length;
  if(count >= MAX_WORKERS) return;
  workersWrap.insertAdjacentHTML("beforeend", workerTemplate(count+1));
  const w = workersWrap.lastElementChild;
  ["beginn","ende"].forEach(n=>{
    w.querySelector(`input[name="${n}"]`).addEventListener("change", recalc);
  });
  recalc();
}

addBtn.addEventListener("click", addWorker);

// add the first mandatory worker
addWorker();

// Before submit, shape fields as arrays the backend expects (vorname[], …)
form.addEventListener("submit", (e)=>{
  // convert flat inputs to array fields via hidden inputs
  const names = ["vorname","nachname","ausweis","beginn","ende"];
  names.forEach(n=>{
    const values = Array.from(workersWrap.querySelectorAll(`input[name="${n}"]`)).map(i=>i.value);
    // remove old hidden
    form.querySelectorAll(`input[name="${n}"]`).forEach(i=>{
      if(!workersWrap.contains(i)) i.remove();
    });
    // append as multiple fields with same name (FastAPI -> List[str])
    // nothing extra needed; the current inputs already have that name
    // we just ensure they are inside the form, which they are.
  });
});
