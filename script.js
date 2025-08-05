
async function generateXLSX() {
    const XLSX = await import('https://cdn.sheetjs.com/xlsx-0.20.0/package/xlsx.mjs');

    const datum = document.getElementById('datum').value;
    const bauort = document.getElementById('bauort').value;
    const name = document.getElementById('name1').value;
    const vorname = document.getElementById('vorname1').value;
    const ausweis = document.getElementById('ausweis1').value;
    const beginn = document.getElementById('beginn1').value;
    const ende = document.getElementById('ende1').value;

    const wb = XLSX.utils.book_new();
    const ws_data = [
        ["Datum", datum],
        ["Bauort", bauort],
        [],
        ["Name", name],
        ["Vorname", vorname],
        ["Ausweis-Nr.", ausweis],
        ["Beginn", beginn],
        ["Ende", ende]
    ];
    const ws = XLSX.utils.aoa_to_sheet(ws_data);

    XLSX.utils.book_append_sheet(wb, ws, "Nachweis");
    XLSX.writeFile(wb, "arbeitsnachweis.xlsx");
}
