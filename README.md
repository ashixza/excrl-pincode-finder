# Excel Pincode Finder

A deployable web app for filling missing Indian PIN codes in Excel workbooks.

## Architecture

- **Frontend:** static HTML/CSS/JavaScript hosted on Netlify
- **Backend:** Python FastAPI hosted on Render
- **Excel:** `openpyxl` preserves workbook sheets, formatting, and existing values
- **PIN lookup:** third-party Post Office lookup API (`api.postalpincode.in`), which exposes India Post post-office data but is not an official India Post endpoint

## What it does

1. Upload an `.xlsx` workbook.
2. Finds the sheet named `Addresses` (case-insensitive).
3. Detects address and pincode columns using common header names. You can override the column names in the UI.
4. Fills **only blank pincode cells**.
5. First searches address/locality terms. If it cannot find a convincing exact locality match, it tries the city field and chooses a city-level fallback (Head Office preferred where available).
6. Marks each result as `EXACT` or `CITY_FALLBACK` in the on-screen audit. It does not add or alter audit columns in the workbook.
7. Downloads the updated workbook. All other sheets and columns are retained.

## Accuracy and safety

A city can have multiple PIN codes. A city fallback is therefore a best-effort city-level choice, **not a guarantee of delivery to the exact street address**. The UI clearly reports fallback rows so they can be reviewed. If the city itself is ambiguous or no post office is returned, the pincode is left blank and reported as unresolved.

Existing pincodes are never overwritten. Only empty pincode cells in the `Addresses` sheet are written.

The third-party Post Office API is an external public service and may be unavailable, rate-limited, or return incomplete matches. The app does not claim that a six-digit format check proves an address match.

## Local run

### Backend

```bash
cd backend
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

### Frontend

Open `frontend/index.html` in a browser, or serve the `frontend` directory with any static server. Set the API URL in the UI to `http://localhost:8000`.

## Deploy

### 1. Deploy backend to Render

- Push this project to a Git repository.
- In Render, create **New > Web Service** and connect the repository.
- Set **Root Directory** to `backend`.
- Runtime: Python.
- Build command: `pip install -r requirements.txt`
- Start command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
- Add environment variable `FRONTEND_ORIGIN` with your Netlify site URL after deploying the frontend (for initial testing, `*` is allowed; use your exact site origin in production).

A ready-to-use `backend/render.yaml` is included.

### 2. Deploy frontend to Netlify

- In Netlify, choose **Add new site > Import an existing project** and connect the same repository.
- Set **Base directory** to `frontend`.
- Publish directory: `.` (the frontend folder).
- No build command is required.
- Set `VITE_API_BASE` is not required; enter your Render backend URL in the app's API URL field and click **Save API URL**.

A `netlify.toml` is included at the project root. If deploying manually, drag the `frontend` folder into Netlify Drop, then enter your Render URL in the UI.

## Privacy

The workbook is processed in memory by the backend and is not intentionally saved to disk. The generated file is returned as a download. Do not upload files containing data you are not authorized to process. For production use, add authentication, upload limits, request throttling, and an organization-approved privacy review.

## Supported workbook assumptions

- `.xlsx` format (not legacy `.xls`).
- Target worksheet named `Addresses`.
- Headers are in the first row.
- Common header variants are recognized (`Address`, `Full Address`, `Pincode`, `PIN Code`, etc.).
- If the workbook uses unusual headers, enter the exact column header text in the UI.

## API

`POST /process` multipart form:
- `file`: `.xlsx`
- `sheet_name`: defaults to `Addresses`
- `address_column`: optional exact header override
- `pincode_column`: optional exact header override
- `city_column`: optional exact header override

Response: JSON summary, audit rows, and a short-lived download token. Download with `GET /download/{token}`.
