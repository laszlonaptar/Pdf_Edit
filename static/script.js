(function () {
  // Beschreibung counter
  const ta = document.getElementById('beschreibung');
  const cnt = document.getElementById('besch-count');
  if (ta && cnt) {
    const upd = () => { cnt.textContent = (ta.value.length || 0) + ' / 1000'; };
    ta.addEventListener('input', upd); upd();
  }

  // Break checkbox -> hidden minutes
  const half = document.getElementById('break_half');
  const hid = document.getElementById('break_minutes');
  if (half && hid) {
    const sync = () => { hid.value = half.checked ? '30' : '60'; recalcAll(); };
    half.addEventListener('change', sync); sync();
  }

  // Dinamikus dolgozó hozzáadás (max 5)
  const list = document.getElementById('worker-list');
  const addBtn = document.getElementById('add-worker');

  function tplWorker(i) {
    return `
    <fieldset class="worker" data-index="${i}">
      <legend>Mitarbeiter ${i}</legend>

      <div class="grid-3">
        <div class="field"><label>Vorname</label><input name="vorname${i}" type="text" /></div>
        <div class="field"><label>Nachname</label><input name="nachname${i}" type="text" /></div>
        <div class="field"><label>Ausweis-Nr. / Kennzeichen</label><input name="ausweis${i}" type="text" /></div>
      </div>
      <div class="field">
        <label>Vorhaltung / beauftragtes Gerät / Fahrzeug</label>
        <input name="vorhaltung${i}" type="text" />
      </div>

      <div class="grid-3">
        <div class="field"><label>Beginn</label><input name="beginn${i}" class="t-beginn" type="time" step="60" /></div>
        <div class="field"><label>Ende</label><input name="ende${i}" class="t-ende" type="time" step="60" /></div>
        <div class="field"><label>Stunden (auto)</label><input class="stunden-display" type="text" value="" readonly /></div>
      </div>
    </fieldset>`;
  }

  if (addBtn && list) {
    addBtn.addEventListener('click', function () {
      const current = list.querySelectorAll('fieldset.worker').length;
      if (current >= 5) { addBtn.disabled = true; return; }
      const i = current + 1;
      list.insertAdjacentHTML('beforeend', tplWorker(i));
      if (i >= 5) addBtn.disabled = true;
      wireInputs();
    });
  }

  function parseHHMM(v) {
    if (!v) return null;
    const [h, m] = v.split(':').map(Number);
    if (Number.isNaN(h) || Number.isNaN(m)) return null;
    return h * 60 + m;
  }

  function hoursWithBreaks(beginMin, endMin, pauseMin) {
    if (beginMin === null || endMin === null || endMin <= beginMin) return 0;
    const total = endMin - beginMin;
    let minus = 0;
    if (pauseMin >= 60) {
      // 09:00–09:15 + 12:00–12:45
      minus += overlap(beginMin, endMin, 9 * 60, 9 * 60 + 15);
      minus += overlap(beginMin, endMin, 12 * 60, 12 * 60 + 45);
    } else {
      minus = Math.min(total, 30);
    }
    return Math.max(0, (total - minus) / 60);
  }

  function overlap(a1, a2, b1, b2) {
    const s = Math.max(a1, b1);
    const e = Math.min(a2, b2);
    return Math.max(0, e - s);
  }

  function wireInputs() {
    document.querySelectorAll('.t-beginn,.t-ende').forEach(el => {
      el.removeEventListener('input', recalcAll);
      el.addEventListener('input', recalcAll);
    });
  }
  wireInputs();

  function recalcAll() {
    const pauseMin = parseInt(hid ? hid.value : '60', 10) || 60;
    let sum = 0;

    document.querySelectorAll('fieldset.worker').forEach(fs => {
      const beg = fs.querySelector('.t-beginn')?.value || '';
      const end = fs.querySelector('.t-ende')?.value || '';
      const out = fs.querySelector('.stunden-display');

      const bMin = parseHHMM(beg);
      const eMin = parseHHMM(end);
      const h = hoursWithBreaks(bMin, eMin, pauseMin);
      sum += h;
      if (out) out.value = h ? h.toFixed(2) : '';
    });

    const tot = document.getElementById('gesamtstunden_auto');
    if (tot) tot.value = sum ? sum.toFixed(2) : '';
  }

  // első számítás
  recalcAll();
})();
