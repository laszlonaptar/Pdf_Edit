// egyszerű összóraszámítás (ebédszünet nélkül)
const beg = document.querySelector('input[name="beginn1"]');
const end = document.querySelector('input[name="ende1"]');
const out = document.getElementById('total_hours');

function toMin(t) {
  if (!t) return null;
  const [h, m] = t.split(':').map(Number);
  return h*60 + m;
}
function calc() {
  const b = toMin(beg.value);
  const e = toMin(end.value);
  if (b!=null && e!=null && e >= b) {
    const diff = e - b;
    const h = Math.floor(diff/60);
    const m = diff%60;
    out.value = `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
  } else {
    out.value = '';
  }
}
beg?.addEventListener('input', calc);
end?.addEventListener('input', calc);
// Optional JavaScript
