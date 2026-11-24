
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


app = FastAPI(
    title="Excel Pincode Finder API",
    version="1.1.0",
    description="Address-based Indian PIN-code lookup using a third-party Post Office API."
)

# --------------------------------------------------
# CONFIGURATION
# --------------------------------------------------

origins_env = os.getenv("FRONTEND_ORIGIN", "*")
origins = (
    ["*"]
    if origins_env.strip() == "*"
    else [x.strip() for x in origins_env.split(",") if x.strip()]
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

DOWNLOADS = {}
MAX_UPLOAD_BYTES = 15 * 1024 * 1024

API_URL = "https://api.postalpincode.in/postoffice/{}"

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "ExcelPincodeFinder/1.1"
})


# --------------------------------------------------
# TEXT NORMALIZATION
# --------------------------------------------------

def norm(value):
    if value is None:
        return ""

    return re.sub(
        r"[^a-z0-9]+",
        " ",
        str(value).lower()
    ).strip()


def tokens(value):
    stop = {
        "near", "behind", "opposite", "road", "rd",
        "street", "st", "flat", "house", "no",
        "plot", "nearby", "india", "pin", "pincode"
    }

    return {
        t for t in norm(value).split()
        if len(t) > 2 and t not in stop
    }


def cell_text(ws, row, col):
    if not col:
        return ""

    value = ws.cell(row=row, column=col).value

    return "" if value is None else str(value).strip()


# --------------------------------------------------
# COLUMN DETECTION
# --------------------------------------------------

def find_col(headers, override, aliases, required=True):

    if override:
        wanted = norm(override)

        for idx, header in enumerate(headers, start=1):
            if norm(header) == wanted:
                return idx

        raise HTTPException(
            400,
            f"Column '{override}' was not found."
        )

    normalized = [norm(h) for h in headers]

    for alias in aliases:
        if norm(alias) in normalized:
            return normalized.index(norm(alias)) + 1

    if required:
        raise HTTPException(
            400,
            f"Could not identify column. Available headers: {headers}"
        )

    return None


def detect_header_row(
    ws,
    address_override=None,
    pincode_override=None,
    city_override=None
):

    address_aliases = [
        "Address",
        "Full Address",
        "Address Field",
        "Customer Address",
        "Location Address"
    ]

    pincode_aliases = [
        "Pincode",
        "Pin Code",
        "PIN Code",
        "Postal Code",
        "Postcode"
    ]

    city_aliases = [
        "City",
        "Town",
        "Location",
        "District"
    ]

    best = None

    # Search the first 30 rows for the actual header.
    for row_num in range(1, min(ws.max_row, 30) + 1):

        headers = [
            ws.cell(row_num, col).value
            for col in range(1, ws.max_column + 1)
        ]

        normalized = [norm(h) for h in headers]

        def locate(override, aliases):

            if override:
                wanted = norm(override)

                return (
                    normalized.index(wanted) + 1
                    if wanted in normalized
                    else None
                )

            for alias in aliases:
                wanted = norm(alias)

                if wanted in normalized:
                    return normalized.index(wanted) + 1

            return None

        address_col = locate(
            address_override,
            address_aliases
        )

        pincode_col = locate(
            pincode_override,
            pincode_aliases
        )

        city_col = locate(
            city_override,
            city_aliases
        )

        # Address header is mandatory.
        if address_col is None:
            continue

        score = 10

        if pincode_col:
            score += 5

        if city_col:
            score += 2

        if best is None or score > best[0]:
            best = (
                score,
                row_num,
                address_col,
                pincode_col,
                city_col
            )

    if best is None:
        raise HTTPException(
            400,
            "Could not find an Address header in the first 30 rows. "
            "Check your worksheet and column names."
        )

    (
        _,
        header_row,
        address_col,
        pincode_col,
        city_col
    ) = best

    if pincode_override and pincode_col is None:
        raise HTTPException(
            400,
            "Specified Pincode column was not found."
        )

    if city_override and city_col is None:
        raise HTTPException(
            400,
            "Specified City column was not found."
        )

    return (
        header_row,
        address_col,
        pincode_col,
        city_col
    )


# --------------------------------------------------
# POST OFFICE API
# --------------------------------------------------

def get_post_offices(query):

    query = str(query or "").strip()

    if not query:
        return []

    try:
        response = SESSION.get(
            API_URL.format(
                requests.utils.quote(query, safe="")
            ),
            timeout=12
        )

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


# --------------------------------------------------
# MATCHING LOGIC
# --------------------------------------------------

def score_office(office, address, city, state):

    office_name = norm(office.get("Name"))
    district = norm(office.get("District"))
    state_name = norm(office.get("State"))

    address_tokens = tokens(address)
    city_tokens = tokens(city)

    score = 0

    if office_name and office_name in norm(address):
        score += 8

    score += 3 * len(
        address_tokens & tokens(office.get("Name"))
    )

    score += 2 * len(
        city_tokens & (
            tokens(office.get("Name"))
            | tokens(office.get("District"))
        )
    )

    if state and state_name:
        if state_name in norm(state) or norm(state) in state_name:
            score += 4
        else:
            score -= 5

    if district and district in norm(address):
        score += 3

    if office.get("BranchType") == "Head Office":
        score += 1

    return score


def choose_exact(address, city, state):

    raw = [
        x.strip()
        for x in re.split(r"[,;\n]", str(address or ""))
        if x.strip()
    ]

    candidates = []
    seen = set()

    for query in raw[:8]:

        query_norm = norm(query)

        if len(query_norm) < 4 or query_norm in seen:
            continue

        seen.add(query_norm)

        candidates.extend(get_post_offices(query))

        time.sleep(0.08)

    if not candidates:
        candidates = get_post_offices(
            str(address or "")[:100]
        )

    unique = OrderedDict()

    for item in candidates:

        pin = str(item.get("Pincode") or "").strip()

        if re.fullmatch(r"\d{6}", pin):

            key = (norm(item.get("Name")), pin)

            unique[key] = item

    ranked = sorted(
        unique.values(),
        key=lambda x: score_office(
            x, address, city, state
        ),
        reverse=True
    )

    if not ranked:
        return None

    best = ranked[0]

    best_score = score_office(
        best, address, city, state
    )

    name_match = (
        norm(best.get("Name")) in norm(address)
        if best.get("Name")
        else False
    )

    token_match = bool(
        tokens(address) & tokens(best.get("Name"))
    )

    state_match = (
        not state
        or norm(best.get("State")) in norm(state)
        or norm(state) in norm(best.get("State"))
    )

    if (
        best_score >= 8
        and state_match
        and (name_match or token_match)
    ):
        return best

    return None


def choose_city_fallback(city, state):

    if not city:
        return None

    results = get_post_offices(city)

    valid = [
        x for x in results
        if re.fullmatch(
            r"\d{6}",
            str(x.get("Pincode") or "")
        )
    ]

    if state:

        matching = [
            x for x in valid
            if (
                norm(state) in norm(x.get("State"))
                or norm(x.get("State")) in norm(state)
            )
        ]

        if matching:
            valid = matching

        elif valid:
            return None

    if not valid:
        return None

    valid.sort(
        key=lambda x: (
            x.get("BranchType") == "Head Office",
            norm(x.get("Name")) == norm(city)
        ),
        reverse=True
    )

    return valid[0]


# --------------------------------------------------
# HEALTH CHECK
# --------------------------------------------------

@app.get("/health")
def health():

    return {
        "status": "ok",
        "service": "excel-pincode-finder"
    }


# --------------------------------------------------
# EXCEL PROCESSING
# --------------------------------------------------

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
        raise HTTPException(
            400,
            "Please upload an .xlsx workbook."
        )

    raw = await file.read()

    if len(raw) > MAX_UPLOAD_BYTES:
        raise HTTPException(
            413,
            "File is larger than the 15 MB upload limit."
        )

    try:
        wb = load_workbook(io.BytesIO(raw))

    except Exception:
        raise HTTPException(
            400,
            "Could not read this Excel workbook."
        )

    actual_sheet = next(
        (
            s for s in wb.sheetnames
            if norm(s) == norm(sheet_name)
        ),
        None
    )

    if not actual_sheet:
        raise HTTPException(
            400,
            f"Worksheet '{sheet_name}' was not found. "
            f"Available sheets: {wb.sheetnames}"
        )

    ws = wb[actual_sheet]

    # AUTOMATIC HEADER DETECTION
    header_row, addr_col, pin_col, city_col = detect_header_row(
        ws,
        address_column,
        pincode_column,
        city_column
    )

    # Create Pincode column if missing.
    if pin_col is None:

        pin_col = ws.max_column + 1

        ws.cell(
            row=header_row,
            column=pin_col
        ).value = "Pincode"

    filled_exact = 0
    filled_fallback = 0
    unresolved = 0
    already_present = 0

    audit = []

    # Start immediately below the detected header.
    for row in range(header_row + 1, ws.max_row + 1):

        existing = cell_text(
            ws, row, pin_col
        )

        address = cell_text(
            ws, row, addr_col
        )

        city = (
            cell_text(ws, row, city_col)
            if city_col
            else ""
        )

        # Never overwrite an existing Pincode.
        if existing:

            already_present += 1
            continue

        if not address and not city:

            unresolved += 1

            audit.append({
                "row": row,
                "address": "",
                "pincode": "",
                "status": "UNRESOLVED",
                "reason": "No address or city"
            })

            continue

        state = ""

        parts = [
            p.strip()
            for p in re.split(r"[,;\n]", address)
            if p.strip()
        ]

        if len(parts) >= 2:
            state = parts[-1]

        # First try exact locality matching.
        exact = choose_exact(
            address,
            city,
            state
        )

        selected = exact
        status = "EXACT"

        # If exact match fails, try city fallback.
        if not selected:

            selected = choose_city_fallback(
                city,
                state
            )

            status = "CITY_FALLBACK"

        if selected:

            pin = str(
                selected.get("Pincode")
            ).strip()

            # Validate before writing.
            if not re.fullmatch(r"\d{6}", pin):

                unresolved += 1

                audit.append({
                    "row": row,
                    "address": address,
                    "pincode": "",
                    "status": "UNRESOLVED",
                    "reason": "Invalid Pincode returned"
                })

                continue

            ws.cell(
                row=row,
                column=pin_col
            ).value = pin

            if status == "EXACT":
                filled_exact += 1

            else:
                filled_fallback += 1

            audit.append({
                "row": row,
                "address": address,
                "pincode": pin,
                "status": status,
                "post_office": selected.get("Name"),
                "district": selected.get("District"),
                "state": selected.get("State")
            })

        else:

            unresolved += 1

            audit.append({
                "row": row,
                "address": address,
                "pincode": "",
                "status": "UNRESOLVED",
                "reason": "No reliable locality or city match returned"
            })

    # --------------------------------------------------
    # SAVE OUTPUT WORKBOOK
    # --------------------------------------------------

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    token = uuid.uuid4().hex

    output_filename = (
        filename.rsplit(".", 1)[0]
        + "_pincodes_filled.xlsx"
    )

    DOWNLOADS[token] = (
        output.getvalue(),
        output_filename,
        time.time()
    )

    # Expire old downloads after 30 minutes.
    for key, (_, _, created) in list(DOWNLOADS.items()):

        if time.time() - created > 1800:
            DOWNLOADS.pop(key, None)

    return {
        "download_token": token,
        "sheet": actual_sheet,
        "header_row": header_row,
        "rows_checked": max(
            ws.max_row - header_row,
            0
        ),
        "filled_exact": filled_exact,
        "filled_city_fallback": filled_fallback,
        "unresolved": unresolved,
        "already_had_pincode": already_present,
        "audit": audit
    }


# --------------------------------------------------
# EXCEL DOWNLOAD
# --------------------------------------------------

@app.get("/download/{token}")
def download(token: str):

    item = DOWNLOADS.pop(token, None)

    if not item:
        raise HTTPException(
            404,
            "Download expired or was already used. "
            "Please process the workbook again."
        )

    data, filename, _ = item

    return StreamingResponse(
        io.BytesIO(data),
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition":
            f'attachment; filename="{filename}"'
        }
    )
