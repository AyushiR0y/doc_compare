# Deploying Document Comparison Studio to Streamlit

## Which file is the app?

| File | What it is | Status |
| --- | --- | --- |
| `streamlit_app.py` | **Deploy this.** Streamlit UI on the current engine, with MIS e-mail. | Current |
| `preview_viewer.py` | The embedded side-by-side viewer (sync scrolling, page and change stepping). | Current |
| `comparison_core.py` | The comparison engine. Shared by the Streamlit app and the API. | Current |
| `mis_mailer.py` | Builds and sends the MIS report. | Current |
| `fastapi_app.py` + `frontend/` | FastAPI + React build of the same engine (Render deployment). | Parallel deployment |
| `app.py` | The previous Streamlit app, with its own older copy of the comparison logic. | Superseded |

`app.py` is left in place but is no longer the entry point. It carries a duplicate,
older version of the diff logic, so it will not pick up engine improvements.

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy to Streamlit Community Cloud

1. Push the repository to GitHub (`.streamlit/secrets.toml` is git-ignored and must
   stay that way).
2. On <https://share.streamlit.io> choose **New app**, pick the repo and branch.
3. Set **Main file path** to `doc_compare/streamlit_app.py`.
4. Open **Advanced settings -> Python version** and select **3.10** or later.
5. Open **Settings -> Secrets** and paste:

   ```toml
   GMAIL_USERNAME = "commissionsmis@gmail.com"
   GMAIL_APP_PASSWORD = "your-16-character-app-password"
   MIS_EMAIL_TO = "Ayushi.Roy01@bajajlife.com,bhagyashree.daga@bajajlife.com"
   ```

6. Deploy. Reporting is invisible in the interface by design, so confirm it is
   working by checking the inbox after the first comparison, or by looking for a
   `MIS mail:` line in the app logs.

`requirements.txt` is shared with the FastAPI deployment, so it also installs
`fastapi`/`uvicorn`. They are unused by the Streamlit app and harmless.

## The interface

One page, no sidebar: pick two files, compare, read the summary, download the
highlighted files, and review them side by side. The viewer owns its own controls,
so there is no settings panel competing with it for width - which matters, because
the sidebar used to squeeze the viewer below its stacking breakpoint on 1280px
screens and on 1366px ones at Windows 125% scaling.

## The side-by-side viewer

Streamlit cannot synchronise scrolling between two of its own containers, so
`preview_viewer.py` renders both documents, the toolbar and all the interaction
into a single HTML document, embedded with `st.iframe` via a `data:` URL. The
iframe gets an opaque origin, which is why it is fully isolated from the host
page - and why nothing outside it can read into it.

It offers:

* **Sync toggle** and two modes - **Mouse** (the wheel moves both panes by the
  same amount) and **Page** (stepping moves both panes a page at a time).
* **Page stepping** with a page indicator, plus **Align** for two documents that
  sit a page apart. Pages line up 1:1 unless you press Align; the resulting
  offset is shown in the toolbar so it is never a silent surprise. Page mode
  needs two PDFs - a DOCX renders as one continuous sheet, and the toolbar says so.
* **Change stepping** - jump between highlighted differences in either document,
  with the current one ringed in orange. Per-page change counts show under each page.
* **Zoom** and keyboard shortcuts: arrows for pages, `N`/`P` for changes, `S` for sync.

## MIS reporting

Every completed comparison sends one mail to every address in `MIS_EMAIL_TO`,
containing:

* **Usage details** - date and time in IST and UTC, processing time, session ID,
  client IP, resolved location, and the environment the app ran in.
* **Files uploaded** - file name, format, size, page count, word count and the
  number of highlighted words, per document.
* **Changes highlighted** - total highlighted changes, the split across the two
  documents, matching words, match rate and change rate.
* **Outputs generated** - the names of the two highlighted files.
* **Attachment** - the same record as JSON, for consolidating into a wider MIS.

Reporting is a back-office concern, so it is **deliberately invisible in the
interface** - there is no indicator, no toggle, and no mention of it anywhere on
screen. The mail is sent while the user waits; if it fails, the failure is written
to the app log as a `MIS mail:` line and the comparison carries on unaffected.

Transport: `smtp.gmail.com` over **587/STARTTLS**, falling back to **465/SSL**.
587 is tried first because corporate networks frequently drop the implicit-TLS
connection on 465 — that is exactly what happens on the Bajaj network.

### Credentials

`GMAIL_APP_PASSWORD` is deliberately **not** in any committed file. Settings are read
from the environment first, then `st.secrets`, then the non-secret defaults in
`mis_mailer.py`, so only the password strictly has to be set.

The environment taking precedence means a deployment can redirect or switch reporting
off without touching the secrets - running with `MIS_EMAIL_TO=""` leaves it
unconfigured and nothing is sent, which is how the test runs avoid mailing anyone.

If the password ever needs rotating, generate a new App Password on the Gmail account
(Google Account -> Security -> 2-Step Verification -> App passwords) and update the
Secrets box. Nothing in the repository needs to change.

## Usage log

Comparisons are also appended to `~/.doc_compare/usage_logs.jsonl`, the same log the
FastAPI build writes. Nothing in the app reads it back - the in-app usage dashboard
was removed - so it is just a local record.

> On Streamlit Community Cloud that file lives on ephemeral disk and is wiped whenever
> the app restarts or is redeployed. **The MIS e-mails are the durable record.** Point
> `USAGE_LOG_FILE_PATH` at persistent storage if you want the log to survive too.

## Debug logging

Engine debug logs echo the text of uploaded documents, so they are off by default.
Set `DOC_COMPARE_DEBUG=1` locally if you need them. Never enable this on a hosted
deployment — the logs would capture the contents of every uploaded file.
