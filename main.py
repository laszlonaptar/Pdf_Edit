# main.py
from io import BytesIO
from datetime import datetime, time, timedelta

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from openpyxl import load_workbook

# =========================
# SABLON CELLAPOZÍCIÓK
# (ha a sablonod mást kíván, ezeket a címeket kell átírni)
CELL_DATE = "E3"          # Dátum
CELL_BAU = "E4"           # Bau / Ausführungsort
CELL_BF  = "J3"           # BASF-Beauftragter
TOTAL_CELL = "H26"        # Összes óraszám (alsó összesítés cella)
DESC_FIRST_ROW = 6        # Napi leírás kezdő sora (A6..A15)
DESC_MAX_LINES = 10       # ennyi sort írunk ki a leírásból
WORKERS_FIRST_ROW = 17    # első dolgozó sora
# =========================

# Fix szünetek (zárt intervallumok)
BREAKS = [
    (time(9, 0),  time(9, 15)),   # 09:00–09:15  -> 0.25 óra
    (time(12, 0), time(12, 45)),  # 12:00–12:45 -> 0.75 óra
]

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


# ---------- Segédfüggvények ----------

def _first(form, *keys, default=""):
    """Vegye az első, létező és nem üres mezőt a megadott kulcsok közül."""
    for k in keys:
        v = form.get(k)
        if v is None:
            # többértékes kulcs (MultiDict) esetén próbáljuk listaként
            vals = form.getlist(k) if hasattr(form, "getlist") else []
            if vals:
                v = vals[0]
        if v is not None and str(v).strip() != "":
            return str(v)
    return default


def _get_list(form, *keys):
    """Olvasson ki többször előforduló mezőket listába (name, vorname, stb.)."""
    out = []
    for k in keys:
        vals = []
        if hasattr(form, "getlist"):
            vals = form.getlist(k) or []
        else:
            v = form.get(k)
            if v is not None:
                # ha vesszővel elválasztott érkezik
                vals = [x.strip() for x in str(v).split(",")]
        vals = [str(x).strip() for x in vals if str(x).strip() != ""]
        if vals:
            # összevonjuk (a legelső kulcs döntő; a többi kiegészíthet)
            if not out:
                out = vals[:]
            else:
                # ha hosszabb, bővítjük üresekkel, hogy zip-szerűen illeszkedjen
                if len(vals) > len(out):
                    out += [""] * (len(vals) - len(out))
                for i, v in enumerate(vals):
                    if out[i] == "" and v != "":
                        out[i] = v
    return out


def _parse_hhmm(s: str) -> time | None:
    s = (s or "").strip()
    for fmt in ("%H:%M", "%H.%M", "%H%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except Exception:
            pass
    return None


def _overlap_minutes(a1: time, a2: time, b1: time, b2: time) -> int:
    """Két időintervallum átfedésének hossza percben (fél nyitott: [start, end])."""
    dt = datetime.combine(datetime.today(), time(0, 0))
    A1, A2 = dt.replace(hour=a1.hour, minute=a1.minute), dt.replace(hour=a2.hour, minute=a2.minute)
    B1, B2 = dt.replace(hour=b1.hour, minute=b1.minute), dt.replace(hour=b2.hour, minute=b2.minute)
    start = max(A1, B1)
    end = min(A2, B2)
    delta = (end - start).total_seconds() / 60
    return max(0, int(delta))


def _calc_hours(beg: str, end: str) -> float:
    """Bruttó munkaidő – fix szünetek levonása (órában)."""
    tb = _parse_hhmm(beg)
    te = _parse_hhmm(end)
    if not tb or not te:
        return 0.0
    if te <= tb:
        # ha átnyúlik éjfélen, korrigáljuk
        dt0 = datetime.combine(datetime.today(), tb)
        dt1 = datetime.combine(datetime.today(), te) + timedelta(days=1)
    else:
        dt0 = datetime.combine(datetime.today(), tb)
        dt1 = datetime.combine(datetime.today(), te)

    gross_min = int((dt1 - dt0).total_seconds() // 60)

    # szünetek levonása
    minus = 0
    for sb, se in BREAKS:
        minus += _overlap_minutes(tb, te, sb, se)

    net_min = max(0, gross_min - minus)
    return round(net_min / 60.0, 2)


# ---------- Routes ----------

@app.get("/", response_class=HTMLResponse)
async def form_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/generate_excel")
async def generate_excel(request: Request):
    try:
        form = await request.form()

        # fejléc mezők (több kulcsnév támogatása a régi verziók miatt)
        datum      = _first(form, "datum", "date")
        projekt    = _first(form, "projekt", "bau", "bauort")
        bf         = _first(form, "bf", "beauftragter", "basf")
        taetigkeit = _first(form, "taetigkeit", "beschreibung", "leiras")
        vorhaltung = _first(form, "vorhaltung", "geraet", "eszkoz", default="")

        # dolgozók
        names    = _get_list(form, "name", "nachname")
        vornamen = _get_list(form, "vorname", "keresztnev")
        ausweise = _get_list(form, "ausweis", "ausweisnr", "id", "szemelyi")
        begins   = _get_list(form, "beginn", "start")
        ends     = _get_list(form, "ende", "finish", "end")

        if not projekt or not names:
            return JSONResponse(
                status_code=400,
                content={"detail": "Kötelező mezők hiányoznak: projekt/bau és legalább egy dolgozó (név)."}
            )

        # Excel sablon betöltése
        wb = load_workbook("GP-t.xlsx")
        ws = wb.active

        # Fejléc cellák
        ws[CELL_DATE] = str(datum or "")
        ws[CELL_BAU]  = str(projekt or "")
        ws[CELL_BF]   = str(bf or "")

        # Napi leírás A6..A15 (soronként egy bejegyzés)
        lines = [ln.strip() for ln in str(taetigkeit or "").split("\n")]
        for i, ln in enumerate(lines[:DESC_MAX_LINES]):
            ws[f"A{DESC_FIRST_ROW + i}"] = ln

        # Dolgozók kitöltése
        r0 = WORKERS_FIRST_ROW
        total = 0.0
        n = max(len(names), len(vornamen), len(ausweise), len(begins), len(ends))
        for i in range(n):
            r = r0 + i
            name    = str(names[i]) if i < len(names) else ""
            vorname = str(vornamen[i]) if i < len(vornamen) else ""
            ausw    = str(ausweise[i]) if i < len(ausweise) else ""
            b       = str(begins[i]) if i < len(begins) else ""
            e       = str(ends[i]) if i < len(ends) else ""

            if not name and not vorname:
                continue

            # A..F oszlopok: Név, Keresztnév, Ausweis, Kezd, Vég, Óra
            ws[f"A{r}"] = name
            ws[f"B{r}"] = vorname
            ws[f"C{r}"] = ausw
            ws[f"D{r}"] = b
            ws[f"E{r}"] = e
            h = _calc_hours(b, e)
            ws[f"F{r}"] = h
            total += h

            # G (Vorhaltung/Gerät) – egyelőre minden sorba ugyanazt tesszük (ha kell)
            if vorhaltung:
                ws[f"G{r}"] = str(vorhaltung)

        # Összes óra
        ws[TOTAL_CELL] = total

        # Válasz (xlsx letöltés)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        fname = f"arbeitsnachweis_{(datum or 'heute').replace('.', '-')}.xlsx"
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={fname}"}
        )

    except Exception as e:
        # részletes stacktrace a Render logba
        import sys, traceback
        traceback.print_exc(file=sys.stderr)
        return JSONResponse(status_code=500, content={"error": str(e)})
