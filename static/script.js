let arbeiterCount = 1;

function addArbeiter() {
    arbeiterCount++;
    const container = document.getElementById("arbeiterContainer");

    const div = document.createElement("div");
    div.className = "arbeiter";
    div.id = `arbeiter${arbeiterCount}`;
    div.innerHTML = `
        <h3>Arbeiter ${arbeiterCount}</h3>
        <label>Vorname: <input type="text" name="vorname${arbeiterCount}" required></label>
        <label>Nachname: <input type="text" name="nachname${arbeiterCount}" required></label>
        <label>Ausweis-Nr.: <input type="text" name="ausweis${arbeiterCount}" required></label>
        <label>Beginn: <input type="time" name="beginn${arbeiterCount}" required></label>
        <label>Ende: <input type="time" name="ende${arbeiterCount}" required></label>
    `;
    container.appendChild(div);
}

document.getElementById("arbeitsForm").addEventListener("change", () => {
    let total = 0;
    for (let i = 1; i <= arbeiterCount; i++) {
        const beginn = document.querySelector(`[name=beginn${i}]`);
        const ende = document.querySelector(`[name=ende${i}]`);
        if (beginn && ende && beginn.value && ende.value) {
            const [bh, bm] = beginn.value.split(":").map(Number);
            const [eh, em] = ende.value.split(":").map(Number);
            let hours = (eh * 60 + em - (bh * 60 + bm)) / 60;
            if (hours > 0) {
                // Szünetek levonása
                if (bh <= 9 && eh >= 9.25) hours -= 0.25;
                if (bh <= 12 && eh >= 12.75) hours -= 0.75;
                total += hours;
            }
        }
    }
    document.getElementById("gesamtstunden").value = total.toFixed(2);
});
