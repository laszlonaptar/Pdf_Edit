// Add up to 5 workers
const workersDiv = document.getElementById('workers');
const addBtn = document.getElementById('addWorker');
let workerCount = 0;
const MAX_WORKERS = 5;

function addWorkerRow() {
  if (workerCount >= MAX_WORKERS) return;
  workerCount++;
  const wrap = document.createElement('div');
  wrap.className = 'worker';
  wrap.innerHTML = `
    <label>Vorname
      <input type="text" name="vorname${workerCount}" />
    </label>
    <label>Nachname
      <input type="text" name="nachname${workerCount}" />
    </label>
    <label>Ausweis-Nr.
      <input type="text" name="ausweis${workerCount}" />
    </label>
    <div class="row">
      <label>Beginn
        <input type="time" name="beginn${workerCount}" />
      </label>
      <label>Ende
        <input type="time" name="ende${workerCount}" />
      </label>
    </div>
  `;
  workersDiv.appendChild(wrap);
}

addBtn.addEventListener('click', addWorkerRow);
// add first worker by default
addWorkerRow();

// Compute total hours minus fixed breaks
function parseHM(t){
  if(!t) return null;
  const [h,m] = t.split(':').map(x=>parseInt(x,10));
  return h*60+m;
}
function overlap(a1,a2,b1,b2){
  const s = Math.max(a1,b1), e = Math.min(a2,b2);
  return Math.max(0, e-s);
}
function minutesToHours(m){ return Math.round((m/60)*100)/100; }

function computeTotal(){
  let totalMin = 0;
  const break1 = [parseHM('09:00'), parseHM('09:15')];
  const break2 = [parseHM('12:00'), parseHM('12:45')];

  for(let i=1;i<=workerCount;i++){
    const b = document.querySelector(`[name="beginn${i}"]`)?.value;
    const e = document.querySelector(`[name="ende${i}"]`)?.value;
    if(!b || !e) continue;
    const mb = parseHM(b), me = parseHM(e);
    if(mb==null || me==null || me<=mb) continue;
    let dur = me - mb;
    dur -= overlap(mb, me, break1[0], break1[1]);
    dur -= overlap(mb, me, break2[0], break2[1]);
    if(dur>0) totalMin += dur;
  }
  document.getElementById('total_hours').value = minutesToHours(totalMin).toFixed(2);
}
setInterval(computeTotal, 500);
