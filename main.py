# main.py
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from datetime import datetime, time
from io import BytesIO
from typing import List, Dict, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


# --------- kis segédfüggvények ---------
def parse_time(s: str) -> time:
    s = (s or "").strip()
    # elfogadjuk pl. "7:00", "07:00", "07.00"
    s = s.replace(".", ":")
    return datetime.strptime(s, "%H:%M").time()


def overlap_minutes(a_start: time, a_end: time, b_start: time, b_end: time) -> int:
    """Két idősáv átfedésének hossza percben."""
    a0 = a_start.hour * 60 + a_start.minute
    a1 = a_end.hour * 60 + a_end.minute
    b0 = b_start.hour * 60 + b_start.minute
    b1 = b_end.hour * 60 + b_end.minute
    start = max(a0, b0)
    end = min(a1, b1)
    return max(0, end - start)


def net_minutes(start: time, end: time) -> int:
    """Nettó percek a reggeli (09:00–09:15) és ebéd (12:00–12:45) szünet levonása után."""
    total = overlap_minutes(start, end, start, end)
    total -= overlap_minutes(start, end, time(9, 0), time(9, 15))
    total -= overlap_minutes(start, end, time(12, 0), time(12, 45))
    return max(0, total)


def find_text(ws, text_substr: str) -> Optional[Tuple[int, int]]:
    """Megkeresi az első cellát, amelynek értékében szerepel text_substr (részszöveg)."""
    t = text_substr.lower()
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            v = str(cell.value).strip() if cell.value is not None else ""
            if t in v.lower():
                return (cell.row, cell.column)
    return None


def find_header_row_and_cols(ws, header_names: Dict[str, List[str]]) -> Tuple[int, Dict[str, int]]:
    """
    Visszaadja a fejléc sor indexét és a kért oszlopok pozícióit.
    header_names: kulcs = logikai név, érték = lehetséges felirat-változatok
    """
    # keressünk egy olyan sort, ahol legalább 3 kulcs felirata megvan
    best_row = None
    best_hits = {}
    for r in range(1, ws.max_row + 1):
        row_values = [str(ws.cell(row=r, column=c).value or "").strip().lower() for c in range(1, ws.max_column + 1)]
        hits = {}
        for key, variants in header_names.items():
            col_idx = None
            for c, val in enumerate(row_values, start=1):
                if any(v.lower() in val for v in variants):
                    col_idx = c
                    break
            if col_idx:
                hits[key] = col_idx
        if len(hits) >= 3:
            best_row = r
            best_hits = hits
            break
    if not best_row:
        # ha nem találtuk, dobjunk fel valami értelmes alapértelmezést
        raise RuntimeError("Nem találom a dolgozói táblázat fejlécét a sablonban.")
    return best_row, best_hits


def write_wrapped(ws, cell_addr: str, text: str, indent: int = 1):
    c = ws[cell_addr]
    c.value = text
    c.alignment = Alignment(wrap_text=True, horizontal="left", vertical="top", indent=indent)


# --------- frontend ---------
@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


# --------- Excel generálás ---------
@app.post("/generate_excel")
async def generate_excel(request: Request):
    form = await request.form()

    # --- űrlap mezők (rugalmas kulcskezelés) ---
    def g(key: str, default: str = "") -> str:
        return str(form.get(key, default)).strip()

    datum = g("datum") or g("date") or g("dat")
    bau = g("bau") or g("bauort") or g("projekt") or g("baustelle")
    bf = g("bf") or g("basf") or g("beauftragter") or g("beauftragter_bf")
    geraet = g("geraet") or g("vorhaltung") or g("geraet_fahrzeug") or g("gepakt")
    beschreibung = g("beschreibung") or g("taetigkeit") or g("was") or g("leiras")

    # -------- dolgozók: támogatjuk a [] és a számozott mintákat is --------
    # Listás nevek (HTML name="vorname[]" stb.)
    vorn = form.getlist("vorname[]") or form.getlist("vorname") or []
    nachn = form.getlist("nachname[]") or form.getlist("nachname") or []
    ausw = form.getlist("ausweis[]") or form.getlist("ausweis") or []

    begn = form.getlist("beginn[]") or form.getlist("beginn") or []
    ende = form.getlist("ende[]") or form.getlist("ende") or []

    # számozott (vorname1, vorname2,…)
    if not vorn and any(k.startswith("vorname") for k in form.keys()):
        # próbáljuk 1..5-ig
        for i in range(1, 6):
            if f"vorname{i}" in form:
                vorn.append(g(f"vorname{i}"))
                nachn.append(g(f"nachname{i}"))
                ausw.append(g(f"ausweis{i}"))
                begn.append(g(f"beginn{i}"))
                ende.append(g(f"ende{i}"))

    # szűrés: csak teljes sorok
    workers = []
    for i in range(max(len(vorn), len(nachn), len(ausw), len(begn), len(ende))):
        vn = vorn[i] if i < len(vorn) else ""
        nn = nachn[i] if i < len(nachn) else ""
        aw = ausw[i] if i < len(ausw) else ""
        bs = begn[i] if i < len(begn) else ""
        es = ende[i] if i < len(ende) else ""
        if any([vn, nn, aw, bs, es]):
            workers.append({"vorname": vn, "nachname": nn, "ausweis": aw, "beginn": bs, "ende": es})

    # --- sablon betöltés ---
    wb = openpyxl.load_workbook("GP-t.xlsx")
    ws = wb.active

    # --- Fejléc mezők kitöltése (felirat alapú kereséssel) ---
    try:
        pos_date = find_text(ws, "Datum der Leistungs")
        if pos_date:
            # jobbra egy cellába írunk (klasszikusan a felirat melletti adatmező)
            ws.cell(row=pos_date[0], column=pos_date[1] + 1).value = datum
        else:
            # ha nem találtuk, írjuk be "DÁTUM fallback" jelleggel valahova ésszerű helyre
            ws["B2"] = datum

        pos_bau = find_text(ws, "Bau und Ausführungsort")
        if pos_bau:
            ws.cell(row=pos_bau[0], column=pos_bau[1] + 1).value = bau

        pos_bf = find_text(ws, "BASF-Beauftragter")
        if pos_bf:
            ws.cell(row=pos_bf[0], column=pos_bf[1] + 1).value = bf

        # Gép/Eszköz: ha van külön mező, ide beírhatjuk a nagy táblázat jobb oldalán lévő oszlop tetejére is.
        pos_gear = find_text(ws, "Vorhaltung")  # "Vorhaltung / beauftragtes Gerät / Fahrzeug"
        if pos_gear:
            # egy sorral lejjebb és ugyanabba az oszlopba
            ws.cell(row=pos_gear[0] + 1, column=pos_gear[1]).value = geraet

    except Exception:
        # ha bármelyik fejléc-keresés gondot okozna, nem állítjuk le a folyamatot
        pass

    # --- Leírás: A6:G15 blokk (top-left cellába írunk, beállított tördeléssel) ---
    # Ha a sablonban nincs merge, ez nem baj: A6-ba így is beírjuk.
    write_wrapped(ws, "A6", beschreibung or "")

    # --- Dolgozói táblázat: fejléc-keresés és írás ---
    header_map = {
        "name": ["name"],
        "vorname": ["vorname"],
        "ausweis": ["ausweis", "kennzeichen"],
        "beginn": ["beginn"],
        "ende": ["ende"],
        "stunden": ["anzahl stunden", "stunden"],
    }
    header_row, cols = find_header_row_and_cols(ws, header_map)

    # az adatsor kezdete – a sablonodon jellemzően a fejléc UTÁNI 2. sor
    data_start_row = header_row + 2

    total_minutes = 0
    for idx, w in enumerate(workers[:5]):  # max 5 dolgozó
        r = data_start_row + idx
        # nevek
        if "name" in cols:
            ws.cell(row=r, column=cols["name"]).value = w["nachname"]
            ws.cell(row=r, column=cols["name"]).alignment = Alignment(horizontal="left", vertical="center")
        if "vorname" in cols:
            ws.cell(row=r, column=cols["vorname"]).value = w["vorname"]
            ws.cell(row=r, column=cols["vorname"]).alignment = Alignment(horizontal="left", vertical="center")
        if "ausweis" in cols:
            ws.cell(row=r, column=cols["ausweis"]).value = w["ausweis"]
            ws.cell(row=r, column=cols["ausweis"]).alignment = Alignment(horizontal="left", vertical="center")

        # idők + nettó órák
        try:
            t_start = parse_time(w["beginn"])
            t_end = parse_time(w["ende"])
            if "beginn" in cols:
                ws.cell(row=r, column=cols["beginn"]).value = t_start.strftime("%H:%M")
                ws.cell(row=r, column=cols["beginn"]).alignment = Alignment(horizontal="center", vertical="center")
            if "ende" in cols:
                ws.cell(row=r, column=cols["ende"]).value = t_end.strftime("%H:%M")
                ws.cell(row=r, column=cols["ende"]).alignment = Alignment(horizontal="center", vertical="center")

            mins = net_minutes(t_start, t_end)
            total_minutes += mins
            hours = round(mins / 60.0, 2)
            if "stunden" in cols:
                ws.cell(row=r, column=cols["stunden"]).value = hours
                ws.cell(row=r, column=cols["stunden"]).alignment = Alignment(horizontal="center", vertical="center")
        except Exception:
            # ha bármi baj van a formátummal, ne dőljön el a folyamat
            pass

    # --- Gesamtstunden mező ---
    pos_total = find_text(ws, "Gesamtstunden")
    if pos_total:
        ws.cell(row=pos_total[0], column=pos_total[1] + 1).value = round(total_minutes / 60.0, 2)

    # --- mentés és visszaadás ---
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"Arbeitsnachweis_{datum or datetime.now().strftime('%Y-%m-%d')}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
