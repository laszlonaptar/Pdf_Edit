from __future__ import annotations

# ========================== Stdlib & Typing ==========================
import io
import os
import re
import json
import smtplib
import mimetypes
from email.message import EmailMessage
from datetime import datetime, time as dtime
from typing import List, Optional, Dict

# ========================== FastAPI & Friends ==========================
from fastapi import (
    FastAPI,
    Request,
    Response,
    HTTPException,
    Query,
    Body,
)
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import (
    HTMLResponse,
    FileResponse,
    StreamingResponse,
    JSONResponse,
    RedirectResponse,
)
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates

# ========================== Pydantic v2 ==========================
from pydantic import BaseModel, Field, field_validator

# ========================== HTTP client ==========================
import httpx

# ========================== Excel ==========================
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

# ========================== Optional Google Drive ==========================
# Ha NINCS telepítve a kliens, akkor ezek opcionálisak maradnak.
try:
    from google.oauth2.service_account import Credentials as GCreds
    from googleapiclient.discovery import build as gbuild
    from googleapiclient.http import MediaIoBaseUpload
    HAS_GOOGLE = True
except Exception:
    HAS_GOOGLE = False


# ======================================================================
#                           APP KONFIG
# ======================================================================
APP_TITLE = "PDF-Edit"
TZ = "Europe/Berlin"  # információs jellegű
DEFAULT_TEMPLATE = os.getenv("XLS_TEMPLATE", "GP-t.xlsx")

app = FastAPI(title=APP_TITLE)

# Statikus és sablon könyvtárak bekötése (ha léteznek)
if os.path.isdir("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="templates") if os.path.isdir("templates") else None

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=os.getenv("CORS_ALLOW_ORIGINS", "*").split(","),
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ======================================================================
#                           HEALTH & ROOT
# ======================================================================
@app.get("/healthz")
def healthz_get():
    return {"ok": True}

@app.head("/healthz")
def healthz_head():
    return Response(status_code=200)

LANDING_HTML = """<!doctype html>
<html lang="de">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>PDF-Edit – Start</title>
</head>
<body>
  <main>
    <h1>PDF-Edit ist online</h1>
    <p>Weiter zur Anwendung: <a href="/app">/app</a></p>
  </main>
</body>
</html>"""

@app.get("/", response_class=HTMLResponse)
def index_get():
    resp = HTMLResponse(content=LANDING_HTML, status_code=200)
    resp.headers["Cache-Control"] = "no-store"
    return resp

@app.head("/")
def index_head():
    resp = Response(status_code=200)
    resp.headers["Cache-Control"] = "no-store"
    return resp

@app.get("/app", response_class=HTMLResponse)
def app_page(request: Request):
    if templates and os.path.exists(os.path.join("templates", "index.html")):
        resp = templates.TemplateResponse("index.html", {"request": request})
    else:
        resp = HTMLResponse("<h2>App</h2><p>templates/index.html nicht gefunden.</p>", status_code=200)
    resp.headers["Cache-Control"] = "no-store"
    return resp

@app.get("/favicon.ico")
def favicon():
    path = "static/favicon.ico"
    if os.path.exists(path):
        return FileResponse(path, media_type="image/x-icon")
    return Response(status_code=204)


# ======================================================================
#                           FEATURE CONFIG
# ======================================================================
@app.get("/api/config")
def api_config():
    """
    A frontend ezt olvassa ki: mit jelenítsen meg.
    """
    return {
        "features": {
            "autocomplete": True,
            "translation_button": True,
            "email_send": bool(os.getenv("SMTP_HOST")),     # csak ha be van állítva SMTP
            "drive_upload": bool(os.getenv("GDRIVE_FOLDER_ID") and os.getenv("GDRIVE_SA_JSON")),
        },
        "i18n": {
            "default": "de",
            "available": ["de", "hr", "en"],
        },
    }


# ======================================================================
#                           AUTOCOMPLETE
# ======================================================================
DEFAULT_WORKERS = [
    {"last_name": "Muster", "first_name": "Max", "id": "A001"},
    {"last_name": "Beispiel", "first_name": "Erika", "id": "A002"},
    {"last_name": "Novak", "first_name": "Ivan", "id": "HR015"},
]

def load_workers() -> List[Dict[str, str]]:
    """
    Források prioritása:
    1) WORKERS_JSON környezeti változó (JSON lista)
    2) workers.json fájl
    3) beépített DEFAULT_WORKERS
    """
    env_json = os.getenv("WORKERS_JSON")
    if env_json:
        try:
            data = json.loads(env_json)
            if isinstance(data, list):
                return data
        except Exception:
            pass
    if os.path.exists("workers.json"):
        try:
            with open("workers.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            pass
    return DEFAULT_WORKERS

@app.get("/api/workers")
def api_workers(q: Optional[str] = Query(default=None, description="Name oder Ausweis-ID Filter")):
    workers = load_workers()
    if not q:
        return {"items": workers}
    q_low = q.lower().strip()
    filtered = [
        w for w in workers if any(
            q_low in (w.get(k, "") or "").lower() for k in ("last_name", "first_name", "id")
        )
    ]
    return {"items": filtered}


# ======================================================================
#                           TRANSLATION (proxy)
# ======================================================================
class TranslateReq(BaseModel):
    text: str = Field(..., description="Input text")
    source: Optional[str] = Field(None, description="Source language, e.g. 'hr'")
    target: str = Field(..., description="Target language, e.g. 'de'")

@app.post("/api/translate")
async def api_translate(req: TranslateReq):
    """
    LibreTranslate (vagy kompatibilis) proxy.
    Env:
      LT_ENDPOINT = 'https://libretranslate.de/translate' (példa)
      LT_API_KEY  = '...' (opcionális)
    """
    endpoint = os.getenv("LT_ENDPOINT")
    if not endpoint:
        raise HTTPException(status_code=501, detail="Translation backend not configured (LT_ENDPOINT).")

    payload = {"q": req.text, "source": req.source or "auto", "target": req.target, "format": "text"}
    headers = {"Content-Type": "application/json"}
    api_key = os.getenv("LT_API_KEY")
    if api_key:
        payload["api_key"] = api_key
    timeout = httpx.Timeout(10.0, connect=5.0)
    async with httpx.AsyncClient(timeout=timeout) as client:
        r = await client.post(endpoint, json=payload, headers=headers)
        if r.status_code != 200:
            raise HTTPException(status_code=502, detail=f"Translator error: {r.status_code}")
        data = r.json()
        translated = data.get("translatedText")
        if not translated:
            raise HTTPException(status_code=502, detail="Invalid response from translator")
        return {"translated": translated}


# ======================================================================
#                           EXCEL GENERÁLÁS
# ======================================================================
BREAKS = [
    (dtime(9, 0), dtime(9, 15), 0.25),
    (dtime(12, 0), dtime(12, 45), 0.75),
]

def parse_hhmm(t: str) -> dtime:
    m = re.fullmatch(r"\s*(\d{1,2}):(\d{2})\s*", t or "")
    if not m:
        raise ValueError("Ungültiges Zeitformat, erwartet HH:MM")
    hh = int(m.group(1)); mm = int(m.group(2))
    if not (0 <= hh < 24 and 0 <= mm < 60):
        raise ValueError("Ungültige Stunde/Minute")
    return dtime(hh, mm)

def overlap_minutes(a_start: dtime, a_end: dtime, b_start: dtime, b_end: dtime) -> int:
    a0 = a_start.hour*60 + a_start.minute
    a1 = a_end.hour*60 + a_end.minute
    b0 = b_start.hour*60 + b_start.minute
    b1 = b_end.hour*60 + b_end.minute
    lo = max(a0, b0); hi = min(a1, b1)
    return max(0, hi - lo)

def compute_hours(start: dtime, end: dtime) -> float:
    if end <= start:
        raise ValueError("Ende muss nach Beginn sein")
    total_min = (end.hour*60 + end.minute) - (start.hour*60 + start.minute)
    deduct = 0
    for b0, b1, hrs in BREAKS:
        minutes = overlap_minutes(start, end, b0, b1)
        if minutes > 0:
            deduct += hrs * min(1.0, minutes / ((b1.hour*60 + b1.minute) - (b0.hour*60 + b0.minute)))
    net_hours = max(0.0, total_min/60.0 - deduct)
    return round(net_hours, 2)

class WorkerIn(BaseModel):
    last_name: str
    first_name: str
    id: str = Field(..., description="Ausweis-Nr.")
    start: Optional[str] = None
    end: Optional[str] = None
    vorhaltung: Optional[str] = None

class GenerateExcelIn(BaseModel):
    date: str = Field(..., description="YYYY-MM-DD")
    project: str = Field(..., description="Bau und Ausführungsort")
    bf: str = Field(..., description="Bauleiter/Fachbauleiter")
    description: str = Field(..., description="Was wurde gemacht?")
    workers: List[WorkerIn]
    copy_first_worker_time: bool = Field(default=True, description="Copy first worker time to others if missing")

    @field_validator("workers")
    @classmethod
    def at_least_one_worker(cls, v):
        if not v or len(v) < 1:
            raise ValueError("Mindestens 1 Arbeiter erforderlich")
        if len(v) > 5:
            raise ValueError("Maximal 5 Arbeiter unterstützt")
        return v

def fill_description(ws: Worksheet, text: str):
    """
    A6–A15 sorok (10 sor) – egyszerű 70 karakteres wrap.
    Ha más cellakiosztást használtok, itt kell igazítani.
    """
    lines: List[str] = []
    words = re.split(r"(\s+)", (text or "").strip())
    current = ""
    for token in words:
        if len(current) + len(token) > 70 and current:
            lines.append(current)
            current = token.lstrip()
        else:
            current += token
    if current:
        lines.append(current)
    for i in range(10):
        cell = f"A{6+i}"
        ws[cell] = lines[i] if i < len(lines) else ""

def write_workers(ws: Worksheet, workers: List[WorkerIn]):
    """
    Feltételezett mapping:
      Sorok: 18..22 (max 5 dolgozó)
      Oszlopok: A=Nachname, B=Vorname, C=Ausweis, D=Beginn, E=Ende, F=Std (nettó), G=Vorhaltung
    """
    base_row = 18
    first_time = None
    if workers and workers[0].start and workers[0].end:
        first_time = (workers[0].start, workers[0].end)

    for idx in range(5):
        r = base_row + idx
        w = workers[idx] if idx < len(workers) else None
        if not w:
            for col in "ABCDEFG":
                ws[f"{col}{r}"] = ""
            continue

        ws[f"A{r}"] = w.last_name
        ws[f"B{r}"] = w.first_name
        ws[f"C{r}"] = w.id

        s = w.start or (first_time[0] if (first_time and idx > 0) else None)
        e = w.end or (first_time[1] if (first_time and idx > 0) else None)
        ws[f"D{r}"] = s or ""
        ws[f"E{r}"] = e or ""

        net = ""
        if s and e:
            try:
                net = compute_hours(parse_hhmm(s), parse_hhmm(e))
            except Exception:
                net = ""
        ws[f"F{r}"] = net
        ws[f"G{r}"] = w.vorhaltung or ""

def write_header(ws: Worksheet, date_str: str, project: str, bf: str):
    ws["C2"] = project
    ws["C3"] = bf
    try:
        d = datetime.strptime(date_str, "%Y-%m-%d").date()
        ws["C1"] = d.strftime("%d.%m.%Y")
    except Exception:
        ws["C1"] = date_str

@app.post("/generate_excel")
def generate_excel(payload: GenerateExcelIn):
    template_path = DEFAULT_TEMPLATE
    if not os.path.exists(template_path):
        raise HTTPException(status_code=500, detail=f"Template nicht gefunden: {template_path}")

    try:
        wb = load_workbook(template_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Vorlage kann nicht geöffnet werden: {e}")

    ws = wb.active

    write_header(ws, payload.date, payload.project, payload.bf)
    fill_description(ws, payload.description or "")
    write_workers(ws, payload.workers)

    # Összóra: F18..F22 → F23 (szükség esetén igazítsd a cél cellát)
    try:
        total = 0.0
        for r in range(18, 23):
            v = ws[f"F{r}"].value
            if isinstance(v, (int, float)):
                total += float(v)
            elif isinstance(v, str):
                try:
                    total += float(v.replace(",", "."))
                except Exception:
                    pass
        ws["F23"] = round(total, 2)
    except Exception:
        pass

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = f"Arbeitsnachweis_{payload.date}.xlsx"
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
        "Cache-Control": "no-store",
    }
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


# ======================================================================
#                           E-MAIL KÜLDÉS (opcionális)
# ======================================================================
class EmailPayload(BaseModel):
    to: str
    subject: str
    body: str
    filename: Optional[str] = None  # opcionális csatolmány fájlnév (a /generate_excelből kapott)
    content_b64: Optional[str] = None  # vagy base64 tartalom közvetlen (nem kötelező)

def _smtp_send(msg: EmailMessage):
    host = os.getenv("SMTP_HOST")
    port = int(os.getenv("SMTP_PORT", "587"))
    user = os.getenv("SMTP_USER")
    pwd  = os.getenv("SMTP_PASS")
    tls  = os.getenv("SMTP_TLS", "1") == "1"
    if not host:
        raise HTTPException(status_code=500, detail="SMTP nincs konfigurálva")

    with smtplib.SMTP(host, port) as s:
        if tls:
            s.starttls()
        if user:
            s.login(user, pwd or "")
        s.send_message(msg)

@app.post("/api/send_email")
def api_send_email(p: EmailPayload):
    """
    Egyszerű SMTP e-mail küldés.
    Env:
      SMTP_HOST, SMTP_PORT(=587), SMTP_USER, SMTP_PASS, SMTP_FROM, SMTP_TLS(=1)
    """
    from_addr = os.getenv("SMTP_FROM", os.getenv("SMTP_USER", "noreply@example.com"))
    msg = EmailMessage()
    msg["From"] = from_addr
    msg["To"] = p.to
    msg["Subject"] = p.subject
    msg.set_content(p.body or "")

    # Ha csatolmányt is küldenénk base64 nélkül, akkor a frontend küldje be a tartalmat base64-ben.
    # Jelen verzióban csak sima szövegküldésre optimális.
    # (Ha kell bővíthetjük fájlolvasással a szerveren, de ez általában nem kívánatos.)

    _smtp_send(msg)
    return {"ok": True}


# ======================================================================
#                           GOOGLE DRIVE FELTÖLTÉS (opcionális)
# ======================================================================
class DriveUploadIn(BaseModel):
    filename: str
    content_b64: str  # a kliens küldi base64-ben
    folder_id: Optional[str] = None  # ha nincs megadva, GDRIVE_FOLDER_ID env

@app.post("/api/drive_upload")
def api_drive_upload(p: DriveUploadIn):
    """
    Service Account alapú Drive feltöltés.
    Env:
      GDRIVE_SA_JSON   = a service account JSON tartalma (egysoros JSON)
      GDRIVE_FOLDER_ID = alapértelmezett mappa ID
    """
    if not HAS_GOOGLE:
        raise HTTPException(status_code=501, detail="Google kliens nincs telepítve ezen a környezeten")

    sa_json = os.getenv("GDRIVE_SA_JSON")
    if not sa_json:
        raise HTTPException(status_code=500, detail="GDRIVE_SA_JSON nincs beállítva")

    folder_id = p.folder_id or os.getenv("GDRIVE_FOLDER_ID")
    if not folder_id:
        raise HTTPException(status_code=400, detail="Nincs Drive mappa ID (folder_id vagy GDRIVE_FOLDER_ID)")

    try:
        info = json.loads(sa_json)
    except Exception:
        raise HTTPException(status_code=500, detail="GDRIVE_SA_JSON hibás formátum")

    creds = GCreds.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/drive.file"],
    )
    service = gbuild("drive", "v3", credentials=creds)

    # base64 dekód
    import base64
    raw = base64.b64decode(p.content_b64)

    media = MediaIoBaseUpload(io.BytesIO(raw), mimetype=mimetypes.guess_type(p.filename)[0] or "application/octet-stream")
    file_metadata = {"name": p.filename, "parents": [folder_id]}
    r = service.files().create(body=file_metadata, media_body=media, fields="id, webViewLink, webContentLink").execute()

    # Publikus olvasás (opcionális – ha kell publikus link)
    try:
        service.permissions().create(
            fileId=r["id"],
            body={"role": "reader", "type": "anyone"},
        ).execute()
        pub = service.files().get(fileId=r["id"], fields="webViewLink, webContentLink").execute()
    except Exception:
        pub = {"webViewLink": None, "webContentLink": None}

    return {"id": r.get("id"), "view": pub.get("webViewLink"), "download": pub.get("webContentLink")}


# ======================================================================
#                   BACKWARD-COMPAT / ALIAS ENDPOINTOK
# ======================================================================
@app.get("/config")
def _alias_config():
    return api_config()

@app.get("/workers")
def _alias_workers(q: Optional[str] = Query(default=None)):
    return api_workers(q=q)

@app.post("/api/translate_text")
async def _alias_translate(req: TranslateReq):
    return await api_translate(req)

@app.post("/api/generate_excel")
def _alias_generate_excel(payload: GenerateExcelIn):
    return generate_excel(payload)


# ======================================================================
#                           UVICORN INDÍTÁS
# ======================================================================
def _port() -> int:
    try:
        return int(os.getenv("PORT", "10000"))
    except Exception:
        return 10000

if __name__ == "__main__":
    # Notebook/fejlesztői környezetben csak explicit kérésre induljon
    if os.getenv("RUN_UVICORN", "0") == "1":
        import uvicorn
        uvicorn.run(app, host="0.0.0.0", port=_port())
