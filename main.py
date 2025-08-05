from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
import io
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    datum = request.form.get('datum')
    projekt = request.form.get('projekt')
    geraet = request.form.get('geraet')
    bf = request.form.get('bf')
    beschreibung = request.form.get('beschreibung')

    mitarbeiter = []
    for i in range(1, 6):
        nachname = request.form.get(f'nachname{i}')
        vorname = request.form.get(f'vorname{i}')
        ausweis = request.form.get(f'ausweis{i}')
        if nachname and vorname and ausweis:
            mitarbeiter.append((nachname, vorname, ausweis))

    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    ws['B2'] = datum
    ws['B3'] = projekt
    ws['B4'] = geraet
    ws['F3'] = bf
    ws['B5'] = beschreibung

    for index, (nachname, vorname, ausweis) in enumerate(mitarbeiter):
        row = 8 + index
        ws[f'B{row}'] = nachname
        ws[f'C{row}'] = vorname
        ws[f'D{row}'] = ausweis

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"arbeitsnachweis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True)