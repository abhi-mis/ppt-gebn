<p align="center">
  <img src="https://img.shields.io/badge/react-18-61DAFB?style=for-the-badge&logo=react&logoColor=black" />
  <img src="https://img.shields.io/badge/typescript-5-3178C6?style=for-the-badge&logo=typescript&logoColor=white" />
  <img src="https://img.shields.io/badge/vite-5-646CFF?style=for-the-badge&logo=vite&logoColor=white" />
  <img src="https://img.shields.io/badge/license-MIT-green?style=for-the-badge" />
</p>

<h1 align="center">ReportAI — Frontend</h1>

<p align="center">
  <b>A sleek dark-themed UI for automated project report generation</b><br/>
  <sub>Upload Excel → Configure project details → Generate PPT → Chat with AI about your slides</sub>
</p>

<p align="center">
  <a href="https://github.com/abhinav3254">
    <img src="https://img.shields.io/badge/built%20by-Abhinav%20Jha-blueviolet?style=flat-square" />
  </a>
</p>

---

## Overview

The frontend is a single-page React application that provides an end-to-end interface for generating PowerPoint status reports from Excel data. Users upload Excel files, fill in project details, generate the report, and then interact with it through an AI-powered chat, real-time slide previews, and on-the-fly style customization.

---

## Screens

### Landing Page
The entry point. Users upload one or more Excel files, optionally fill in project details (title, author, date, distribution roles), toggle auto-summarization, and hit generate. A live pipeline tracker shows progress through upload → extraction → generation → fixing stages.

### Workspace
After generation, the app transitions to a three-panel workspace:

| Panel | Description |
|-------|-------------|
| **Sidebar** (left) | Four tabs — Actions, Details, Style, AI — for downloading, editing project details, restyling the PPT, and configuring AI behavior |
| **Chat** (center) | AI-powered conversational interface to ask questions about slide content |
| **Activity Feed** (right) | Real-time slide summaries, images, and processing logs streamed via WebSocket |

---

## Project Structure

```
src/
├── App.tsx                      ← Root: Landing or Workspace based on report state
├── main.tsx                     ← Entry point + providers
├── index.css                    ← Global styles + CSS variables
│
├── api/
│   ├── axios.ts                 ← Axios instance (base URL config)
│   └── socket.ts                ← Socket.IO client instance
│
├── components/
│   ├── Landing.tsx              ← Upload page with Excel multi-file + project details
│   ├── Workspace.tsx            ← Three-panel layout (sidebar + chat + feed)
│   ├── Sidebar.tsx              ← Actions / Details / Style / AI tabs
│   ├── ChatPanel.tsx            ← AI chat interface
│   ├── ActivityFeed.tsx         ← Real-time slide summaries + images
│   ├── MultiFileUpload.tsx      ← Drag-and-drop multi-file upload with file list
│   ├── FileUpload.tsx           ← Single file upload (used for backward compat)
│   └── ProjectDetailsForm.tsx   ← Reusable project details form (title, author, roles)
│
├── context/
│   ├── ReportContext.tsx         ← Global report state (upload result, clear)
│   ├── SocketContext.tsx         ← WebSocket state (summaries, restyle, refine)
│   └── AIOptionsContext.tsx      ← AI toggle settings (persisted in sessionStorage)
│
└── types/
    └── index.ts                 ← TypeScript interfaces (ReportData, StyleConfig, etc.)
```

---

## Key Components

### `MultiFileUpload`
Handles multiple Excel file uploads with drag-and-drop, deduplication by name+size, a scrollable file list with remove buttons, and a file count badge.

### `ProjectDetailsForm`
A reusable form used in both the Landing page (pre-generation) and the Sidebar Details tab (post-generation). Fields include project title, author, report date, date format (UK/US), and a dynamic distribution roles list with add/remove and a "Prefill defaults" button. Persists to `sessionStorage`.

### `Sidebar`
Four tabs:

| Tab | Purpose |
|-----|---------|
| **Actions** | Download PPT, Refine Comments, Generate Summaries |
| **Details** | Edit project details + Apply to PPT (calls `/api/update-project-details`) |
| **Style** | Header/data/title font, size, color pickers + Apply Style (WebSocket restyle) |
| **AI** | Toggle switches for summary grouping, jargon detection, risk flags, etc. |

### `SocketContext`
Manages all WebSocket events: slide summaries (streamed), refine comments (with progress tracking), restyle status, and on-demand summary generation. Exposes clean hooks like `requestGenerateSummaries(pptName)`.

---

## API Integration

### REST (Axios)

| Method | Endpoint | Used By |
|--------|----------|---------|
| `POST` | `/api/upload-report` | Landing → upload Excel + details |
| `POST` | `/api/update-project-details` | Sidebar Details → edit after generation |
| `GET` | `/api/download-report/:name` | Sidebar Actions → download PPT |

### WebSocket (Socket.IO)

| Event (emit) | Triggered From |
|-------------|----------------|
| `generate_summaries` | Sidebar Actions |
| `refine_comments` | Sidebar Actions |
| `restyle_ppt` | Sidebar Style tab |

| Event (listen) | Handled In |
|----------------|------------|
| `slide_summary` | ActivityFeed |
| `summary_complete` | SocketContext |
| `refine_progress` | Sidebar |
| `restyle_complete` | Sidebar |

---

## Design System

The UI uses a custom dark theme with CSS variables defined in `index.css`:

| Variable | Purpose |
|----------|---------|
| `--bg-primary` | Main background |
| `--bg-secondary` | Card/panel background |
| `--bg-tertiary` | Input/button background |
| `--accent` | Primary accent (purple) |
| `--cyan` | Secondary accent |
| `--text-primary` | Main text |
| `--text-muted` | Secondary/label text |
| `--success` / `--error` | Status colors |
| `--border` | Border color |

Typography uses `--font-sans` for body and `--font-mono` for labels/badges.

---

## Setup

```bash
# Clone
git clone https://github.com/abhinav3254/report-ai-frontend.git
cd report-ai-frontend

# Install
npm install

# Configure backend URL (optional — defaults to http://localhost:8000/api)
echo "VITE_API_BASE_URL=http://localhost:8000/api" > .env
echo "VITE_SOCKET_BASE_URL=http://localhost:8000" >> .env

# Run
npm run dev
```

App starts at `http://localhost:5173`

---

## Build

```bash
npm run build
```

Output goes to `dist/`. Serve with any static file server.

---

## Environment

| Dependency | Version |
|-----------|---------|
| React | 18 |
| TypeScript | 5 |
| Vite | 5 |
| Axios | 1.x |
| Socket.IO Client | 4.x |

---

<p align="center">
  <sub>Built with focus by <a href="https://github.com/abhinav3254"><b>Abhinav Jha</b></a></sub>
</p>
