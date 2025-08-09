from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from datetime import datetime, time
from io import BytesIO
from openpyxl import load_workbook

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


def _calc_hours(start: str, end: str) -> float:
    """Munkaidő számítás a rögzített szünetek levonásával."""
    if not start or not end:
        return 0.0
    try:
        s = datetime.strptime(start, "%H:%M")
        e = datetime.strptime(end, "%H:%M")
    except ValueError:
        return 0.0

    hours = (e - s).total_seconds() / 3600.0

    # 09:00–09:15 (0.25 h) és 12:00–12:45 (0.75 h) szünet
    if s.time() <= time(9, 15) and e.time() >= time(9, 0):
        hours -= 0.25
    if s.time() <= time(12, 45) and e.time() >= time(12, 0):
        hours -= 0.75

    return max(hours, 0.0)


def _first(form, *keys, default=""):
    """Az első létező kulcs értéke (alias-ok támogatása)."""
    for k in keys:
        if k in form and form[k]:
            return form[k]
    return default


def _get_list(form, *names):
    """
    Listát ad vissza dolgozói mezőkből.
    Támogatott:
      - többszörös azonos név (getlist)
      - indexelt kulcsok: name1/name2..., vorname1/vorname2... stb.
    """
    # 1) getlist, ha van
    for n in names:
        vals = form.getlist(n)
        if vals:
            return [v for v in vals if str(v).strip() != ""]
    # 2) indexelt felderítés 1..10
    out = []
    for i in range(1, 11):
        for n in names:
            key = f"{n}{i}"
            if key in form and str(form[key]).strip() != "":
                out.append(form[key])
                break
    return out


@app.post("/generate_excel")
async def generate_excel(request: Request):
    form = await request.form()

    # --- Fejléc mezők (több alias megengedett) ---
    datum       = _first(form, "datum", "date")
    projekt     = _first(form, "projekt", "bau", "bauort")
    bf          = _first(form, "bf", "beauftragter", "basf")
    taetigkeit  = _first(form, "taetigkeit", "beschreibung", "leiras")
    vorhaltung  = _first(form, "vorhaltung", "geraet", "eszkoz", default="")

    # --- Dolgozói listák (többféle elnevezés támogatása) ---
    names    = _get_list(form, "name", "nachname")
    vornamen = _get_list(form, "vorname", "keresztnev")
    ausweise = _get_list(form, "ausweis", "ausweisnr", "id", "szemelyi")
    begins   = _get_list(form, "beginn", "start")
    ends     = _get_list(form, "ende", "finish", "end")

    # Minimális validáció: legyen legalább egy dolgozó + projekt
    if not projekt or not names:
        return JSONResponse(
            status_code=400,
            content={"detail": "Hiányzó kötelező mezők: projekt/bau és legalább 1 dolgozó (név)."}
        )

    # --- Excel sablon betöltése ---
    wb = load_workbook("GP-t.xlsx")
    ws = wb.active

    # A sablonodon eddig ezek a helyek működtek stabilan:
    # Dátum (fejléc bal blokk jobb oldala), Bau (alatta bal oldal), BASF megbízott (jobb felső blokk)
    # Ha nálad más cellák a tutik, írd át ezeket a címeket.
    ws["E3"] = datum          # dátum
    ws["E4"] = projekt        # Bau / Ausführungsort
    ws["J3"] = bf             # BASF-Beauftragter (ha van ilyen mező a sablonban)

    # Tevékenység (A6–G15 tartomány – a sablonod így nézett ki)
    # Egyszerűen soronként írjuk, a wrap-ot az Excel nézi.
    lines = [ln.strip() for ln in str(taetigkeit).split("\n")]
    start_r = 6
    for i, ln in enumerate(lines[:10]):
        ws[f"A{start_r + i}"] = ln

    # Dolgozók táblázat kezdősora (a legutóbbi jó képen ez 17 volt)
    row0 = 17
    total_sum = 0.0
    n = max(len(names), len(vornamen), len(ausweise), len(begins), len(ends))
    for i in range(n):
        r = row0 + i
        name    = names[i] if i < len(names) else ""
        vorname = vornamen[i] if i < len(vornamen) else ""
        ausw    = ausweise[i] if i < len(ausweise) else ""
        b       = begins[i] if i < len(begins) else ""
        e       = ends[i] if i < len(ends) else ""

        if not name and not vorname:
            continue

        ws[f"A{r}"] = name
        ws[f"B{r}"] = vorname
        ws[f"C{r}"] = ausw
        ws[f"D{r}"] = b
        ws[f"E{r}"] = e

        h = _calc_hours(b, e)
        ws[f"F{r}"] = h
        total_sum += h

    # Összesített munkaóra (a képed alapján az alsó "Gesamtstunden" mező)
    ws["H26"] = total_sum  # ha nem stimmel a hely, írd át a pontos célcellára

    # Kész fájl visszaadása
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"arbeitsnachweis_{datum or 'heute'}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={fname}"}
    )
