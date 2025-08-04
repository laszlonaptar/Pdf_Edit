
let employeeCount = 1;

function addEmployee() {
    employeeCount++;
    const container = document.getElementById("employeeFields");
    const div = document.createElement("div");
    div.innerHTML = \`
        <label>Name: <input type="text" name="Name\${employeeCount}"></label>
        <label>Vorname: <input type="text" name="Vorname\${employeeCount}"></label>
        <label>Ausweis-Nr.: <input type="text" name="Ausweis\${employeeCount}"></label>
        <label>Beginn: <input type="time" name="Beginn\${employeeCount}"></label>
        <label>Ende: <input type="time" name="Ende\${employeeCount}"></label>
        <label>Stunden: <input type="text" name="Stunden\${employeeCount}" readonly></label>
    \`;
    container.appendChild(div);
}

function parseTime(t) {
    const [h, m] = t.split(":").map(Number);
    return h + m / 60;
}

function generateXLS() {
    const form = document.forms['workForm'];
    const formData = new FormData(form);
    let totalHours = 0;

    for (let i = 1; i <= employeeCount; i++) {
        const start = formData.get("Beginn" + i);
        const end = formData.get("Ende" + i);
        if (start && end) {
            const duration = parseTime(end) - parseTime(start) - 1;  // 1 Stunde Pause
            const corrected = Math.round(duration * 4) / 4;
            form["Stunden" + i].value = corrected.toFixed(2);
            totalHours += corrected;
        }
    }

    form["Gesamtstunden"].value = totalHours.toFixed(2);

    // XLS Generierung (egyszerű CSV-ként)
    let csvContent = "data:text/csv;charset=utf-8,";
    formData.forEach((value, key) => {
        csvContent += `${key},${value}\n`;
    });

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "arbeitsnachweis.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
