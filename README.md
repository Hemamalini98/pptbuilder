# SlideFormatter (pptbuilder)

A web application for applying a branded PowerPoint template to content slide decks, with support for injecting figures extracted from PDF source files.

---

## Overview

SlideFormatter automates the process of reformatting a raw content `.pptx` to match a master template's styles — fonts, colors, layouts, spacing — and lets users map figures from a source PDF onto the correct slides before exporting the finished deck.

**Workflow (5 steps):**

1. **Template Master** — Upload or select a branded `.pptx` template
2. **Upload Content** — Upload the raw content `.pptx` to be reformatted
3. **PDF Figures** — Upload the source PDF; the app extracts figures and captions
4. **Review Presentation** — Preview slides, map PDF figures to slide placeholders
5. **Export Deck** — Download the styled `.pptx`, an Excel figure report, or an accessibility report

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3.11, FastAPI, uvicorn |
| PPTX processing | python-pptx, lxml |
| PDF extraction | PyMuPDF (fitz) |
| Excel export | openpyxl |
| Frontend | React 18, TypeScript, Zustand, Vite, Tailwind CSS |
| Containerization | Docker, Docker Compose |

---

## Getting Started

### Prerequisites

- Python 3.11+
- Node.js 18+ (for local frontend dev)

### Local Development

**Backend:**

```bash
pip install -r requirements.txt
python server.py
# API available at http://localhost:8001
```

**Frontend:**

```bash
cd frontend
npm install
npm run dev
# UI available at http://localhost:5173
```

### Docker (recommended)

```bash
docker compose up --build
```

- Frontend: [http://localhost:5050](http://localhost:5050)
- Backend API: [http://localhost:8001](http://localhost:8001)

---

## Project Structure

```
pptbuilder/
├── server.py           # FastAPI app and all API routes
├── main.py             # Template extraction — reads .pptx, serializes styles to JSON
├── convert.py          # Core styling engine — applies template styles to content slides
├── report.py           # Figure diagnostics report generation
├── accessibility.py    # Accessibility report generation
├── excel_report.py     # Excel export of figure/mapping data
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
├── uploads/            # Session-isolated working files (gitignored)
└── frontend/
    └── src/
        ├── App.tsx
        ├── store.ts                    # Zustand global state
        └── components/
            ├── Step1Template.tsx
            ├── Step2Upload.tsx
            ├── Step3Figures.tsx
            ├── Step4Mapping.tsx
            └── Step5Export.tsx
```

---

## API Reference

| Method | Endpoint | Description |
|---|---|---|
| `POST` | `/api/upload-template` | Upload a new template `.pptx` |
| `GET` | `/api/templates` | List available templates |
| `POST` | `/api/select-template` | Set the active template for the session |
| `POST` | `/api/upload-ppt` | Upload the content `.pptx` |
| `POST` | `/api/process-ppt` | Apply template styles to the content deck |
| `POST` | `/api/upload-pdf` | Upload the source PDF |
| `GET` | `/api/pdf/captions` | Get extracted figure captions |
| `POST` | `/api/extract` | Extract a figure from the PDF |
| `POST` | `/api/add-image` | Inject an extracted figure into a slide |
| `GET` | `/api/download` | Download the final styled `.pptx` |
| `GET` | `/api/download-excel` | Download the figure report as `.xlsx` |
| `GET` | `/api/accessibility-report` | Download an accessibility report |
| `GET` | `/api/report-data` | Get figure mapping data as JSON |
| `POST` | `/api/reset` | Clear all session state and uploaded files |

Sessions are identified via a `session_id` cookie (or `X-Session-Id` header), and each session gets an isolated working directory under `uploads/`.
