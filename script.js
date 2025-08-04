
const { jsPDF } = window.jspdf;

function createRows() {
  const container = document.getElementById("rows-container");
  for (let i = 1; i <= 5; i++) {
    container.innerHTML += `
      <div class="row">
        <div><label>Name ${i}<input type="text" id="name${i}" /></label></div>
        <div><label>Vorname ${i}<input type="text" id="vorname${i}" /></label></div>
        <div><label>Ausweis-Nr. ${i}<input type="text" id="ausweis${i}" /></label></div>
      </div>
      <div class="row">
        <div><label>Beginn<input type="time" id="beginn${i}" onchange="calc()" /></label></div>
        <div><label>Ende<input type="time" id="ende${i}" onchange="calc()" /></label></div>
        <div><label>Stunden<input type="text" id="stunden${i}" readonly /></label></div>
      </div>
    `;
  }
}
createRows();

function calc() {
  let total = 0;
  for (let i = 1; i <= 5; i++) {
    const b = document.getElementById(`beginn${i}`).value;
    const e = document.getElementById(`ende${i}`).value;
    if (b && e) {
      const start = new Date("1970-01-01T" + b + ":00");
      const end = new Date("1970-01-01T" + e + ":00");
      let diff = (end - start) / 3600000 - 1;
      diff = Math.max(0, Math.round(diff * 4) / 4);
      document.getElementById(`stunden${i}`).value = diff.toFixed(2);
      total += diff;
    }
  }
  document.getElementById("gesamt").innerText = total.toFixed(2);
}

async function generatePDF() {
  const pdf = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const bg = await fetch("template.jpg").then(r => r.blob()).then(b => URL.createObjectURL(b));
  pdf.addImage(bg, "JPEG", 0, 0, 297, 210);

  pdf.setFontSize(10);
  pdf.text(document.getElementById("datum").value, 230, 23);
  pdf.text(document.getElementById("ort").value, 25, 23);
  pdf.text(document.getElementById("beauftragter").value, 25, 29);
  pdf.text(document.getElementById("auftrag").value, 25, 35);

  pdf.text(document.getElementById("beschreibung").value, 25, 45, { maxWidth: 250 });

  let y = 70;
  for (let i = 1; i <= 5; i++) {
    pdf.text(document.getElementById("name" + i).value, 25, y);
    pdf.text(document.getElementById("vorname" + i).value, 70, y);
    pdf.text(document.getElementById("ausweis" + i).value, 120, y);
    pdf.text(document.getElementById("beginn" + i).value, 160, y);
    pdf.text(document.getElementById("ende" + i).value, 180, y);
    pdf.text(document.getElementById("stunden" + i).value, 200, y);
    y += 10;
  }

  pdf.text("Gesamt: " + document.getElementById("gesamt").innerText + " Stunden", 25, y + 5);

  pdf.save("arbeitsnachweis.pdf");
}
