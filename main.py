import io
import os
import re
import time
import uuid
from collections import OrderedDict
from typing import Optional

import requests
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook

app = FastAPI(title="Excel Pincode Finder API", version="1.0.0", description="Uses a third-party Post Office lookup API; not an official India Post service.")

origins_env = os.getenv("FRONTEND_ORIGIN", "*")
origins = ["*"] if origins_env.strip() == "*" else [x.strip() for x in origins_env.split(",") if x.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

# In-memory downloads: suitable for a small single-instance deployment.
# For multi-instance production deployments, use object storage instead.
DOWNLOADS = {}
MAX_UPLOAD_BYTES = 15 * 1024 * 1024
API_URL = "https://api.postalpincode.in/postoffice/{}"
SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "ExcelPincodeFinder/1.0 (address-to-postal-code lookup)"})


def norm(value):
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def tokens(value):
    stop = {"near", "behind", "opposite", "road", "rd", "street", "st", "flat",
            "house", "no", "plot", "nearby", "india", "pin", "pincode"}
    return {t for t in norm(value).split() if len(t) > 2 and t not in stop}


def find_col(headers, override, aliases, required=True):
    if override:
        wanted = norm(override)
        for idx, h in enumerate(headers, start=1):
            if norm(h) == wanted:
                return idx
        raise HTTPException(400, f"Column '{override}' was not found in the first row.")
    normalized = [norm(h) for h in headers]
    for alias in aliases:
        if norm(alias) in normalized:
            return normalized.index(norm(alias)) + 1
    if required:
        raise HTTPException(400, f"Could not identify column. Available headers: {headers}")
    return None


def get_post_offices(query):
    query = str(query or "").strip()
    if not query:
        return []
    try:
        response = SESSION.get(API_URL.format(requests.utils.quote(query, safe="")), timeout=12)
        response.raise_for_status()
        payload = response.json()
        if not isinstance(payload, list) or not payload:
            return []
        item = payload[0] or {}
        if item.get("Status") != "Success":
            return []
        return item.get("PostOffice") or []
    except (requests.RequestException, ValueError):
        return []


def score_office(office, address, city, state):
    office_name = norm(office.get("Name"))
    district = norm(office.get("District"))
    region = norm(office.get("Region"))
    state_name = norm(office.get("State"))
    address_tokens = tokens(address)
    city_tokens = tokens(city)
    score = 0
    if office_name and office_name in norm(address):
        score += 8
    score += 3 * len(address_tokens & tokens(office.get("Name")))
    score += 2 * len(city_tokens & (tokens(office.get("Name")) | tokens(office.get("District"))))
    if state and state_name and (state_name in norm(state) or norm(state) in state_name):
        score += 4
    elif state and state_name:
        score -= 5
    if district and district in norm(address):
        score += 3
    if office.get("BranchType") == "Head Office":
        score += 1
    return score


def choose_exact(address, city, state):
    # Query meaningful address chunks from most specific to less specific.
    raw = [x.strip() for x in re.split(r"[,;\n]", str(address or "")) if x.strip()]
    candidates = []
    seen = set()
    for q in raw[:8]:
        qn = norm(q)
        if len(qn) < 4 or qn in seen:
            continue
        seen.add(qn)
        candidates.extend(get_post_offices(q))
        time.sleep(0.08)
    if not candidates:
        # Try combined address phrase, often accepted by the post-office search endpoint.
        candidates = get_post_offices(str(address or "")[:100])

    unique = OrderedDict()
    for item in candidates:
        pin = str(item.get("Pincode") or "").strip()
        if re.fullmatch(r"\d{6}", pin):
            key = (norm(item.get("Name")), pin)
            unique[key] = item
    ranked = sorted(
        unique.values(),
        key=lambda x: score_office(x, address, city, state),
        reverse=True
    )
    if not ranked:
        return None
    best = ranked[0]
    best_score = score_office(best, address, city, state)
    # Require positive evidence: name/address overlap or a locality term plus state agreement.
    name_match = norm(best.get("Name")) in norm(address) if best.get("Name") else False
    token_match = bool(tokens(address) & tokens(best.get("Name")))
    state_match = not state or norm(best.get("State")) in norm(state) or norm(state) in norm(best.get("State"))
    if best_score >= 8 and state_match and (name_match or token_match):
        return best
    return None


def choose_city_fallback(city, state):
    if not city:
        return None
    results = get_post_offices(city)
    valid = [x for x in results if re.fullmatch(r"\d{6}", str(x.get("Pincode") or ""))]
    if state:
        matching = [x for x in valid if norm(state) in norm(x.get("State")) or norm(x.get("State")) in norm(state)]
        if matching:
            valid = matching
        elif valid:
            return None  # City name/state conflict: do not guess.
    if not valid:
        return None
    # Prefer a Head Office; then an office whose name exactly matches the city; stable first result otherwise.
    valid.sort(key=lambda x: (
        x.get("BranchType") == "Head Office",
        norm(x.get("Name")) == norm(city)
    ), reverse=True)
    return valid[0]


def cell_text(ws, row, col):
    value = ws.cell(row=row, column=col).value
    return "" if value is None else str(value).strip()


@app.get("/health")
def health():
    return {"status": "ok", "service": "excel-pincode-finder"}


@app.post("/process")
async def process_excel(
    file: UploadFile = File(...),
    sheet_name: str = Form("Addresses"),
    address_column: Optional[str] = Form(None),
    pincode_column: Optional[str] = Form(None),
    city_column: Optional[str] = Form(None),
):
    filename = file.filename or "workbook.xlsx"
    if not filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Please upload an .xlsx workbook.")
    raw = await file.read()
    if len(raw) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "File is larger than the 15 MB upload limit.")
    try:
        wb = load_workbook(io.BytesIO(raw))
    except Exception:
        raise HTTPException(400, "Could not read this Excel workbook.")

    actual_sheet = next((s for s in wb.sheetnames if norm(s) == norm(sheet_name)), None)
    if not actual_sheet:
        raise HTTPException(400, f"Worksheet '{sheet_name}' was not found. Available sheets: {wb.sheetnames}")
    ws = wb[actual_sheet]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    addr_col = find_col(headers, address_column, ["Address", "Full Address", "Address Field", "Customer Address"])
    pin_col = find_col(headers, pincode_column, ["Pincode", "Pin Code", "PIN Code", "Postal Code", "Postcode"])
    city_col = find_col(headers, city_column, ["City", "Town", "Location", "District"], required=False)

    filled_exact = 0
    filled_fallback = 0
    unresolved = 0
    already_present = 0
    audit = []

    for row in range(2, ws.max_row + 1):
        existing = cell_text(ws, row, pin_col)
        address = cell_text(ws, row, addr_col)
        city = cell_text(ws, row, city_col) if city_col else ""
        if existing:
            already_present += 1
            continue
        if not address and not city:
            unresolved += 1
            audit.append({"row": row, "address": "", "pincode": "", "status": "UNRESOLVED", "reason": "No address or city"})
            continue

        state = ""
        # Extract a likely state from comma-separated address tail or city text.
        parts = [p.strip() for p in re.split(r"[,;\n]", address) if p.strip()]
        if len(parts) >= 2:
            state = parts[-1]
        exact = choose_exact(address, city, state)
        selected = exact
        status = "EXACT"
        if not selected:
            selected = choose_city_fallback(city, state)
            status = "CITY_FALLBACK"

        if selected:
            pin = str(selected.get("Pincode"))
            ws.cell(row=row, column=pin_col).value = pin
            if status == "EXACT":
                filled_exact += 1
            else:
                filled_fallback += 1
            audit.append({
                "row": row, "address": address, "pincode": pin, "status": status,
                "post_office": selected.get("Name"), "district": selected.get("District"),
                "state": selected.get("State")
            })
        else:
            unresolved += 1
            audit.append({"row": row, "address": address, "pincode": "", "status": "UNRESOLVED",
                          "reason": "No reliable locality or city match returned"})
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    token = uuid.uuid4().hex
    DOWNLOADS[token] = (output.getvalue(), filename.rsplit(".", 1)[0] + "_pincodes_filled.xlsx", time.time())
    # Expire old in-memory files.
    for key, (_, _, created) in list(DOWNLOADS.items()):
        if time.time() - created > 1800:
            DOWNLOADS.pop(key, None)
    return {
        "download_token": token,
        "sheet": actual_sheet,
        "rows_checked": max(ws.max_row - 1, 0),
        "filled_exact": filled_exact,
        "filled_city_fallback": filled_fallback,
        "unresolved": unresolved,
        "already_had_pincode": already_present,
        "audit": audit,
    }


@app.get("/download/{token}")
def download(token: str):
    item = DOWNLOADS.pop(token, None)
    if not item:
        raise HTTPException(404, "Download expired or was already used. Please process the workbook again.")
    data, filename, _ = item
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )
