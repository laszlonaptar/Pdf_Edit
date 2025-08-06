let counter = 1;

function addArbeiter() {
    counter++;
    const section = document.getElementById("arbeiter-section");
    const div = document.createElement("div");
    div.classList.add("arbeiter");
    div.innerHTML = `
        <h3>Arbeiter ${counter}</h3>
        <label>Vorname: <input type="text" name="arbeiter_vorname" required></label>
        <label>Nachname: <input type="text" name="arbeiter_nachname" required></label>
        <label>Ausweis-Nr.: <input type="text" name="arbeiter_ausweis" required></label>
        <label>Beginn: <input type="time" name="arbeiter_beginn" required></label>
        <label>Ende: <input type="time" name="arbeiter_ende" required></label>
    `;
    section.appendChild(div);
}

document.addEventListener("input", function () {
    const beginn = document.getElementsByName("arbeiter_beginn");
    const ende = document.getElementsByName("arbeiter_ende");
    let total = 0;

    for (let i = 0; i < beginn.length; i++) {
        const start = beginn[i].value;
        const end = ende[i].value;

        if (start && end) {
            const [sh, sm] = start.split(":").map(Number);
            const [eh, em] = end.split(":").map(Number);
            let hours = (eh + em / 60) - (sh + sm / 60);

            // szünetek levonása
            if (sh <= 9 && eh >= 9.25) hours -= 0.25;
            if (sh <= 12 && eh >= 12.75) hours -= 0.75;

            total += hours;
        }
    }

    document.getElementById("gesamtstunden").value = total.toFixed(2);
});
