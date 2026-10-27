# Diffly — Document Comparison Studio

Compare two documents (PDF or DOCX, in any combination) and get back the
differences highlighted in place: downloadable copies of both files with every
changed word marked, and a synced side-by-side viewer in the browser.

---

## Prerequisites

| | |
| --- | --- |
| **Python** | 3.10 or later |
| **OS** | Windows, macOS or Linux — nothing here is platform-specific |
| **Disk** | ~600 MB for the virtual environment |

No Microsoft Word, LibreOffice or other external converter is needed. PDFs are
rendered by PyMuPDF and DOCX files are rendered directly to HTML in-process.

---

## Run it locally

From the `doc_compare/` directory:

**Windows (PowerShell)**

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run streamlit_app.py
```

**macOS / Linux**

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Streamlit opens <http://localhost:8501> in your browser. That is the whole
setup — the app runs with no configuration at all. MIS e-mail reporting stays
switched off until you configure it, which is what you want locally.

> If PowerShell blocks `Activate.ps1` with an execution-policy error, either run
> `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` for that session,
> or skip activation and call the interpreter directly:
> `venv\Scripts\python.exe -m streamlit run streamlit_app.py`.

### Using it

1. Upload an **Original** and a **Revised** file. Either can be PDF or DOCX, and
   you can compare a PDF against a Word file.
2. Press **Compare documents**.
3. Read the summary, download the highlighted copies (individually or as a ZIP),
   and review the two documents side by side.

In the side-by-side viewer:

| Control | What it does |
| --- | --- |
| **Sync** + **Mouse** / **Page** | Scroll both panes together, either by wheel or a page at a time |
| **Page** stepper, **Align** | Step through pages; Align handles two documents that sit a page apart |
| **Change** stepper | Jump between highlighted differences, current one ringed in orange |
| **Zoom** (per pane) | Enlarge one side without disturbing the other |
| **Layout** | Switch between side-by-side columns and stacked rows |
| Keyboard | `←` `→` pages · `N` `P` changes · `S` sync |

Yellow marks changed text, pale orange marks a change inside a table.

---

## Optional: MIS e-mail reporting

Every completed comparison can e-mail a usage report. **It is off by default and
needs nothing to run the app.** To switch it on:

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Then fill in the three values in `.streamlit/secrets.toml`:

| Key | Notes |
| --- | --- |
| `GMAIL_USERNAME` | The Gmail account the report is sent from |
| `GMAIL_APP_PASSWORD` | A Gmail **App Password**, not the account password. Google Account → Security → 2-Step Verification → App passwords |
| `MIS_EMAIL_TO` | Comma-separated recipients. Leave empty and nothing is sent |

`.streamlit/secrets.toml` is git-ignored and must stay that way — it holds a live
credential. Environment variables of the same names take precedence over the
file, so `MIS_EMAIL_TO=""` switches reporting off for a single run without
editing anything.

Reporting is deliberately invisible in the interface: there is no indicator and
no toggle. To confirm it works, check the inbox after a comparison, or look for a
`MIS mail:` line in the terminal. A failure to send is logged and the comparison
carries on unaffected.

Comparisons are also appended to `~/.doc_compare/usage_logs.jsonl`. Nothing reads
that file back; it is just a local record.

---

## How it fits together

```
streamlit_app.py     the UI and the entry point  ──┐
preview_viewer.py    the embedded side-by-side viewer
comparison_core.py   the engine: extract, diff, highlight, render  ←┐
mis_mailer.py        builds and sends the MIS report               │
usage_storage.py     resolves the usage-log path                   │
                                                                   │
fastapi_app.py       a parallel REST + React deployment  ──────────┘
frontend/            its Vite/React client
```

`comparison_core.py` is shared, so the Streamlit app and the API always compare
documents identically.

| File | Role |
| --- | --- |
| `streamlit_app.py` | **Run this.** The Streamlit UI. |
| `preview_viewer.py` | Builds the side-by-side viewer as one self-contained HTML document. Streamlit cannot sync scrolling between two of its own containers, so the two panes, the toolbar and all the interaction live in a single iframe. |
| `comparison_core.py` | Word extraction, the diff, highlighting, and preview rendering. |
| `mis_mailer.py` | The MIS report. |
| `usage_storage.py` | Usage-log path resolution. |
| `.streamlit/config.toml` | Theme and the 200 MB upload limit. |
| `DEPLOY_STREAMLIT.md` | Deploying to Streamlit Community Cloud, and the design notes behind the viewer and the reporting. |

### The other deployment

`fastapi_app.py` plus `frontend/` serve the same engine as a REST API with a
React client, deployed to Render via `render.yaml`. It is **not** needed to run
the Streamlit app, and `pip install -r requirements.txt` pulls its three extra
packages along harmlessly. To run it instead:

```bash
cd frontend && npm ci && npm run build && cd ..
uvicorn fastapi_app:app --reload
```

---

## Troubleshooting

**The side-by-side preview is blank.**
Make sure Streamlit is 1.58 or later (`streamlit --version`). The viewer hands
raw HTML to `st.iframe` so it travels in the iframe's `srcdoc`; on older
versions, or if that HTML is ever wrapped into a `data:` URL again, Chromium's
2 MB cap on `data:` URL navigation silently blanks the frame for PDF pairs,
whose page images make the payload large. DOCX pairs stay under the cap, so
"DOCX works but PDF doesn't" is the signature of exactly that problem.

**A large PDF is slow to preview.**
Every page ships to the browser as an image, at roughly 200 KB of HTML per
rendered page — so a 50-page document compared against another comes to about
20 MB. Comparison itself is unaffected; only the preview is. The levers are capping `max_pages` in the
`compare_documents_full` call in `streamlit_app.py`, or lowering the render scale
in `_render_pdf_preview_base64` in `comparison_core.py`.

**"Could not extract text from one or both documents".**
The file is a scanned image with no text layer. There is no OCR step, so these
cannot be compared.

**Uploads over 200 MB are rejected.**
Raise `maxUploadSize` in `.streamlit/config.toml`.

**I need to see what the engine is doing.**
Set `DOC_COMPARE_DEBUG=1`. Never enable this on a hosted deployment — the debug
logs echo the text of every uploaded document.
