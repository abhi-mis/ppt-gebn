<p align="center">
  <img src="https://img.shields.io/badge/python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/flask-2.x-000000?style=for-the-badge&logo=flask&logoColor=white" />
  <img src="https://img.shields.io/badge/socketio-realtime-010101?style=for-the-badge&logo=socketdotio&logoColor=white" />
  <img src="https://img.shields.io/badge/license-MIT-green?style=for-the-badge" />
</p>

<h1 align="center">ReportAI — Backend</h1>

<p align="center">
  <b>Automated project status report generation from Excel to PowerPoint</b><br/>
  <sub>Upload Excel data → Extract & transform → Generate polished PPT slides → AI-powered summaries</sub>
</p>

<p align="center">
  <a href="https://github.com/abhinav3254">
    <img src="https://img.shields.io/badge/built%20by-Abhinav%20Jha-blueviolet?style=flat-square" />
  </a>
</p>

---

## What It Does

ReportAI takes one or more Excel files containing project data (milestones, deliverables, risks, phases, financials) and automatically generates a branded PowerPoint status report using a hardcoded template. It handles everything: reading Excel sheets, mapping data to slides, fitting tables to slide dimensions, styling, and optionally generating AI summaries per slide.

---

## Architecture

```
app.py                          ← Flask entry point + REST API
├── excel_processing/
│   ├── read_excel.py           ← Reads visible Excel sheets → styled JSON
│   ├── structure.py            ← Sheet structure utilities
│   └── validation.py           ← Input validation
├── ppt_generation/
│   ├── slides.py               ← Core slide builder (tables, layout, pagination)
│   ├── ppt_helper.py           ← Cell styling, column widths, milestones, roles
│   ├── sheet_config.py         ← Per-sheet width/margin/behavior constants
│   └── template_loader.py      ← Template utilities
├── sockets/
│   ├── events.py               ← WebSocket event handlers
│   ├── handlers.py             ← Socket routing
│   ├── refine_comments.py      ← AI-powered comment refinement
│   ├── generate_ppt_images.py  ← Slide → image rendering
│   └── socket_service.py       ← SocketIO service layer
├── utils/
│   ├── llm_helper.py           ← LLM abstraction (Claude/OpenAI)
│   ├── process_slides_and_stream.py ← Streaming slide summarization
│   ├── ppt_slide_summary.py    ← Summary extraction
│   ├── file_utils.py           ← File helpers
│   └── image_utils.py          ← Image processing
├── ai_helper/
│   └── open_ai_helper.py       ← OpenAI/Claude API integration
├── ppt_sanitizer.py            ← Post-generation font/style fixes
├── fix_ppt.py                  ← PPT repair and cleanup
├── fix_ppt_fonts.py            ← Font normalization
├── constant.py                 ← Global constants
├── templates/
│   └── default_template.pptx   ← Hardcoded PPT template (place yours here)
├── prompts/
│   └── summarization_prompt.txt← AI summarization prompt
├── uploads/                    ← Temp uploaded files
├── extracted/                  ← Extracted JSON from Excel
└── generated/                  ← Generated PPT output files
```

---

## API Endpoints

### `POST /api/upload-report`

Generates a PPT report from Excel files.

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `excel_files` | File[] | Yes | One or more `.xlsx` files |
| `project_title` | string | No | Title shown on slide 1 |
| `author` | string | No | Author name on slide 1 |
| `report_date` | string | No | Date in `YYYY-MM-DD` format |
| `date_format` | string | No | `UK` (DD/MM/YYYY) or `US` (MM/DD/YYYY) |
| `roles` | JSON string | No | Distribution roles `{"role": "name"}` |
| `auto_summarize` | string | No | `"true"` to enable AI summaries |

**Response:**
```json
{
  "pptDownloadUrl": "http://localhost:5000/api/download-report/...",
  "extractedJson": ["sheet1.json", "sheet2.json"],
  "new_fixed_ppt": "uuid_report-fixed.pptx",
  "excel_count": 2
}
```

### `POST /api/update-project-details`

Updates project details in an already-generated PPT (title, author, date, roles).

```json
{
  "ppt_name": "uuid_report-fixed.pptx",
  "project_title": "My Project",
  "author": "Abhinav Jha",
  "report_date": "2026-04-26",
  "date_format": "UK",
  "roles": { "Project Manager": "Alice", "Lead Architect": "Bob" }
}
```

### `GET /api/download-report/<file_name>`

Downloads the generated PPT file.

---

## WebSocket Events

| Event | Direction | Description |
|-------|-----------|-------------|
| `slide_summary` | Server → Client | Streams per-slide AI summaries |
| `summary_complete` | Server → Client | All summaries done |
| `generate_summaries` | Client → Server | Trigger on-demand summarization |
| `refine_comments` | Client → Server | AI-refine cell comments |
| `refine_progress` | Server → Client | Refinement progress updates |
| `restyle_ppt` | Client → Server | Apply style config to PPT |
| `restyle_complete` | Server → Client | Restyle finished |

---

## Sheet Configuration

`ppt_generation/sheet_config.py` controls per-sheet table behavior:

```python
SHEET_OVERRIDES = {
    "fs": {
        "table_width_in": 12.2,
        "left_margin_in": 0.5,
        "right_margin_in": 0.3,
        "max_col_width_in": 2.5,
        "skip_rag_legend": True,
    },
    "phase": {
        "table_width_in": 12.2,
        "skip_rag_legend": True,
    },
    "assumptions": {
        "compact_row_height_in": 0.22,
    },
}
```

Uses prefix matching — `"FS - Sprint 1"` picks up the `"fs"` config. Add any sheet name to customize its layout.

---

## Key Design Decisions

**Multi-Excel support** — Each Excel file generates its own set of slides appended sequentially. The first Excel populates header slides (title, roles, milestones); subsequent ones skip headers and append data only.

**Solo sheets** — Key Milestones and Deliverables always get their own dedicated slides, never sharing with other sheet data. Configured via `SOLO_SHEETS` in `slides.py`.

**Table overflow protection** — Three layers prevent tables from exceeding slide width: per-sheet config widths → proportional scale factor → hard safety clamp.

**Smart header wrapping** — Single-word headers (RAG, ID, S.No) never wrap; font auto-shrinks instead. Multi-word headers wrap at word boundaries only when the column is too narrow.

**Template-first** — The PPT template is hardcoded at `templates/default_template.pptx`. No user upload needed. Custom templates still accepted via the API for backward compatibility.

---

## Setup

```bash
# Clone
git clone https://github.com/abhinav3254/report-ai-backend.git
cd report-ai-backend

# Install dependencies
pip install -r requirements.txt

# Place your PPT template
cp your_template.pptx templates/default_template.pptx

# Create required directories
mkdir -p uploads extracted generated

# Run
python app.py
```

Server starts at `http://localhost:8000`

---

## Environment

| Dependency | Version |
|-----------|---------|
| Python | 3.10+ |
| Flask | 2.x |
| Flask-SocketIO | 5.x |
| python-pptx | 0.6.x |
| openpyxl | 3.1+ |
| pandas | 2.x |

---

<p align="center">
  <sub>Built with focus by <a href="https://github.com/abhinav3254"><b>Abhinav Jha</b></a></sub>
</p>
