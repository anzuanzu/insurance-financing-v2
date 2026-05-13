# Insurance Financing Quote API

This backend is a cloud API PoC for premium-financing quote scenarios.

Current scope:

- Product: `WUS`
- Financing type: `premium`
- Dividend option: `第7年起持續增購保額`
- Inputs: gender, age, coverage year, face amount or premium, LTV ratio
- Outputs: paid premium, table premium, projected benefit, financing amount, self-fund amount, coverage multiples

## Why this PoC uses a rebuild engine first

The current local environment does not include `LibreOffice`, so this PoC wires the API to the existing `scripts/rebuild_wus_coverage.py` logic that was already validated against the workbook.

The Docker image installs `libreoffice-calc` so we can later swap in a true workbook execution adapter without changing the HTTP contract.

Today the service supports an engine switch:

- `QUOTE_ENGINE=rebuild`
- `QUOTE_ENGINE=libreoffice`
- `QUOTE_ENGINE=auto`

Right now `libreoffice` has an initial `WUS` implementation that writes into the workbook core input cells and reads back workbook outputs after headless recalculation. Locally we are still using `rebuild`, because this machine does not have `LibreOffice`.

When `QUOTE_ENGINE=auto`:

- `faceAmount` driven `WUS` requests prefer workbook execution when `LibreOffice` is available
- `premium`-only `WUS` requests still fall back to the rebuild engine
- if `LibreOffice` is not available, all requests fall back to the rebuild engine

## Run locally

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
PYTHONPATH=../backend:.. QUOTE_ENGINE=rebuild uvicorn app.main:app --reload
```

Open:

- `http://127.0.0.1:8000/health`
- `http://127.0.0.1:8000/docs`

## Example request

```json
{
  "productCode": "WUS",
  "financingType": "premium",
  "gender": "female",
  "age": 69,
  "coverageYear": 20,
  "faceAmount": 1000000,
  "ltvRatio": 0.5,
  "dividendOption": "第7年起持續增購保額"
}
```

## Example response

```json
{
  "productCode": "WUS",
  "financingType": "premium",
  "engine": "rebuild_wus_coverage",
  "sourceWorkbook": "WUS(金控)國泰人壽鑽美滿利率變動型美元終身壽險(定期給付型)-11501版(宣告4.25%-1150101啟用)(保費融資)_i值.xlsx",
  "dividendOption": "第7年起持續增購保額",
  "gender": "female",
  "age": 69,
  "coverageYear": 20,
  "faceAmount": 1000000,
  "premium": 601623,
  "tablePremium": 607700,
  "projectedBenefit": 1133923,
  "coverageBefore": 1.8845,
  "coverageAfter": 3.769,
  "financingAmount": 300812,
  "selfFundAmount": 300811,
  "currency": "USD",
  "notes": []
}
```

## Deploy

This folder is ready to deploy to platforms that accept Docker, such as Render, Railway, Cloud Run, or Azure Web App for Containers.

A starter Render blueprint is included at:

- [render.yaml](/Users/yanganru/Documents/亞灣專區保單融資保費融資債券質押網站開發/render.yaml)

For a real workbook-backed run:

```bash
QUOTE_ENGINE=libreoffice uvicorn app.main:app --host 0.0.0.0 --port 8000
```

Current workbook-backed constraints:

- `WUS` only
- `faceAmount` input is supported
- `premium`-only reverse lookup is still handled by the rebuild engine

## Smoke test after deploy

After the API is running, you can validate the service with:

```bash
cd backend
QUOTE_API_BASE_URL=http://127.0.0.1:8000 python3 scripts/smoke_test_wus_api.py
```

For a cloud deployment, replace `QUOTE_API_BASE_URL` with your deployed API URL:

```bash
QUOTE_API_BASE_URL=https://your-api.example.com python3 scripts/smoke_test_wus_api.py
```

This smoke test checks:

- `/health`
- `WUS` face-amount quote flow
- `WUS` premium-input quote flow
- response contract fields needed by the frontend
