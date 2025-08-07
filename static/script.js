function addWorker() {
    const container = document.getElementById("arbeiter-container");
    const count = container.children.length + 1;
    if (count > 2) return;  // Most 2 dolgozóval

    const div = document.createElement("div");
    div.className = "arbeiter";
    div.innerHTML = `
        <h3>Arbeiter ${count}</h3>
        <label>Vorname: <input type="text" name="vorname${count}" required></label><br>
        <label>Nachname: <input type="text" name="nachname${count}" required></label><br>
        <label>Ausweis-Nr.: <input type="text" name="ausweis${count}" required></label><br>
        <label>Beginn: <input type="time" name="beginn${count}" required></label><br>
        <label>Ende: <input type="time" name="ende${count}" required></label><br>
        <label>Gesamtstunden: <input type="text" name="stunden${count}" readonly></label><br>
    `;
    container.appendChild(div);
}