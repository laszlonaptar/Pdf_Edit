
function addMitarbeiter() {
    const container = document.getElementById('mitarbeiter-container');
    const mitarbeiter = document.createElement('div');
    mitarbeiter.classList.add('mitarbeiter');
    mitarbeiter.innerHTML = `
        <label>Nachname:</label>
        <input type="text" name="nachname[]">
        <label>Vorname:</label>
        <input type="text" name="vorname[]">
        <label>Ausweis-Nr.:</label>
        <input type="text" name="ausweis[]">
        <label>Beginn:</label>
        <input type="time" name="beginn[]">
        <label>Ende:</label>
        <input type="time" name="ende[]">
    `;
    container.appendChild(mitarbeiter);
}
