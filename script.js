
function generatePDF() {
  alert('PDF-Generierung wird in der nächsten Version aktiviert.'); 
}

const container = document.getElementById("rows-container");
for (let i = 1; i <= 5; i++) {
  container.innerHTML += `
    <div class="row">
      <div><label>Name ${i}<input type="text" id="name${i}" /></label></div>
      <div><label>Vorname ${i}<input type="text" id="vorname${i}" /></label></div>
      <div><label>Ausweis-Nr. ${i}<input type="text" id="ausweis${i}" /></label></div>
    </div>
    <div class="row">
      <div><label>Beginn<input type="time" id="beginn${i}" /></label></div>
      <div><label>Ende<input type="time" id="ende${i}" /></label></div>
      <div><label>Anzahl Stunden<input type="text" id="stunden${i}" readonly /></label></div>
      <div><label>Vorhaltung<input type="text" id="vorhaltung${i}" /></label></div>
    </div>
  `;
}
