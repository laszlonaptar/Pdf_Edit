from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime, timedelta

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate_excel", methods=["POST"])
def generate_excel():
    # Dátum
    datum = request.form.get("Datum", "")
    if datum:
        datum_formatted = datetime.strptime(datum, "%Y-%m-%d").strftime("%d.%m.%Y")
    else:
        datum_formatted = ""

    # Projekt / Bauort
    projekt = request.form.get("BauUndAusfuehrungsort", "")
    
    # BASF beauftragter
    basf = request.form.get("BASFBeauftragter", "")
    
    # Tevékenység
    taetigkeit = request.form.get("Taetigkeit", "")
    
    # Kezdési és befejezési idő
    beginn = request.form.get("Beginn", "")
    ende = request.form.get("Ende", "")

    # Eszköz
    geraet = request.form.get("Geraet", "")

    # Munkaidő számítás
    total_hours = ""
    if beginn and ende:
        fmt = "%H:%M"
        start = datetime.strptime(beginn, fmt)
        end = datetime.strptime(ende, fmt)
        total = (end - start).total_seconds() / 3600

        # Reggeli szünet: 9:00–9:15 (0,25h), ebéd: 12:00–12:45 (0,75h)
        if start <= datetime.strptime("09:15", fmt) and end >= datetime.strptime("09:00", fmt):
            total -= 0.25
        if start <= datetime.strptime("12:45", fmt) and end >= datetime.strptime("12:00", fmt):
            total -= 0.75

        total_hours = round(total, 2)

    # Dolgozók
    mitarbeiter = []
    for i in range(1, 6):
        name = request.form.get(f"Nachname{i}", "")
        vorname = request.form.get(f"Vorname{i}", "")
        ausweis = request.form.get(f"Ausweis{i}", "")
        if name and vorname and ausweis:
            mitarbeiter.append((name, vorname, ausweis))

    # Excel sablon betöltése
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # Cellák kitöltése
    ws["E6"] = datum_formatted
    ws["E8"] = projekt
    ws["E10"] = basf
    ws["C14"] = taetigkeit
    ws["E12"] = geraet
    ws["K14"] = beginn
    ws["M14"] = ende
    ws["O14"] = total_hours
    ws["P14"] = total_hours

    # Dolgozók beírása (1–5 sor)
    row_start = 17
    for idx, (name, vorname, ausweis) in enumerate(mitarbeiter):
        ws[f"B{row_start + idx}"] = name
        ws[f"C{row_start + idx}"] = vorname
        ws[f"D{row_start + idx}"] = ausweis

    # Fájl memóriába mentése
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"Arbeitsnachweis_{datum}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
