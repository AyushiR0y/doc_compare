import React, { useMemo, useRef, useState, useEffect } from "react";

const API_BASE = import.meta.env.VITE_API_BASE_URL || "http://127.0.0.1:8000";

function getErrorMessage(error) {
  if (error instanceof Error) {
    return error.message;
  }
  return "Something went wrong while processing the documents.";
}

function DocPreview({ doc, title, scrollRef }) {
  if (!doc?.preview) {
    return null;
  }

  const preview = doc.preview;

  return (
    <section className="preview-card">
      <h3>{title}</h3>
      <p className="preview-subtitle">{doc.name}</p>

      {preview.type === "html" ? (
        <div
          ref={scrollRef}
          className="word-preview preview-surface"
          dangerouslySetInnerHTML={{ __html: preview.html }}
        />
      ) : (
        <div ref={scrollRef} className="pdf-preview preview-surface">
          {preview.pages?.map((page) => (
            <figure key={page.page} className="page-figure preview-page" data-page={page.page}>
              <img
                src={`data:${page.mime_type};base64,${page.image_base64}`}
                alt={`Page ${page.page} preview`}
              />
              <figcaption>Page {page.page}</figcaption>
            </figure>
          ))}
          {preview.truncated ? (
            <p className="muted-note">
              Showing first {preview.pages.length} pages out of {preview.total_pages}.
            </p>
          ) : null}
        </div>
      )}
    </section>
  );
}

function LoadingPreview() {
  return (
    <section className="preview-grid loading-grid" aria-live="polite">
      <article className="preview-card glass-card shimmer-block" />
      <article className="preview-card glass-card shimmer-block" />
    </section>
  );
}

export default function App() {
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [syncScroll, setSyncScroll] = useState(true);
  const [currentPageIndex, setCurrentPageIndex] = useState(0);
  const [syncMode, setSyncMode] = useState("mouse"); // 'mouse' or 'page'
  const [pageOffset, setPageOffset] = useState(0);
  const doc1ScrollRef = useRef(null);
  const doc2ScrollRef = useRef(null);
  const scrollLockRef = useRef(false);

  const canCompare = useMemo(() => file1 && file2 && !loading, [file1, file2, loading]);

  function getPreviewPages(container) {
    if (!container) {
      return [];
    }

    return Array.from(container.querySelectorAll(".preview-page"));
  }

  function getPageCount(doc) {
    return doc?.preview?.page_count || doc?.preview?.pages?.length || 0;
  }

  function getPageIndexForContainer(container) {
    if (!container) {
      return 0;
    }

    const pages = getPreviewPages(container);
    if (!pages.length) {
      return 0;
    }

    let selectedIndex = 0;
    const currentScrollTop = container.scrollTop;

    for (let index = 0; index < pages.length; index += 1) {
      const page = pages[index];
      const pageTop = page.offsetTop;
      if (pageTop <= currentScrollTop + 24) {
        selectedIndex = index;
      } else {
        break;
      }
    }

    return selectedIndex;
  }

  function scrollContainerToPage(container, pageIndex) {
    if (!container) {
      return;
    }

    const pages = getPreviewPages(container);
    if (!pages.length) {
      return;
    }

    const clampedIndex = Math.max(0, Math.min(pageIndex, pages.length - 1));
    pages[clampedIndex].scrollIntoView({ behavior: "auto", block: "start", inline: "nearest" });
  }

  function syncBothPreviewsToPage(pageIndex) {
    const el1 = doc1ScrollRef.current;
    const el2 = doc2ScrollRef.current;

    scrollLockRef.current = true;
    // apply offset for doc2
    const offset = pageOffset || 0;
    scrollContainerToPage(el1, pageIndex);
    scrollContainerToPage(el2, pageIndex + offset);

    window.requestAnimationFrame(() => {
      scrollLockRef.current = false;
    });
  }

  function toggleSync(enabled) {
    if (enabled && syncMode === "page") {
      // compute offset based on current visible pages
      const el1 = doc1ScrollRef.current;
      const el2 = doc2ScrollRef.current;
      const idx1 = getPageIndexForContainer(el1);
      const idx2 = getPageIndexForContainer(el2);
      setPageOffset(idx2 - idx1);
      setCurrentPageIndex(idx1);
    }
    setSyncScroll(enabled);
  }

  useEffect(() => {
    if (!result) {
      return;
    }

    setCurrentPageIndex(0);
  }, [result]);

  useEffect(() => {
    if (!result) {
      return;
    }

    if (!syncScroll) {
      return;
    }

    syncBothPreviewsToPage(currentPageIndex);
  }, [currentPageIndex, result, syncScroll]);

  // Attach native listeners for page-based scrolling
  useEffect(() => {
    if (!result) return;
    const el1 = doc1ScrollRef.current;
    const el2 = doc2ScrollRef.current;
    if (!el1 || !el2) return;

    function onScroll1() {
      if (!syncScroll || scrollLockRef.current) return;
      const pageIndex = getPageIndexForContainer(el1);
      setCurrentPageIndex(pageIndex);
    }

    function onScroll2() {
      if (!syncScroll || scrollLockRef.current) return;
      const pageIndex = getPageIndexForContainer(el2);
      setCurrentPageIndex(pageIndex);
    }

    el1.addEventListener("scroll", onScroll1, { passive: true });
    el1.addEventListener("scroll", onScroll1, { passive: true });
    el2.addEventListener("scroll", onScroll2, { passive: true });

    let onWheel1, onWheel2;
    function normalizeDeltaY(e) {
      return e.deltaMode === 1 ? e.deltaY * 16 : e.deltaMode === 2 ? e.deltaY * el1.clientHeight : e.deltaY;
    }

    if (syncMode === "mouse") {
      onWheel1 = (e) => {
        if (!syncScroll) return;
        e.preventDefault();
        if (scrollLockRef.current) return;
        scrollLockRef.current = true;
        const deltaY = normalizeDeltaY(e);
        const next1 = Math.max(0, Math.min(el1.scrollTop + deltaY, el1.scrollHeight - el1.clientHeight));
        const next2 = Math.max(0, Math.min(el2.scrollTop + deltaY, el2.scrollHeight - el2.clientHeight));
        el1.scrollTop = next1;
        el2.scrollTop = next2;
        window.requestAnimationFrame(() => (scrollLockRef.current = false));
      };

      onWheel2 = (e) => {
        if (!syncScroll) return;
        e.preventDefault();
        if (scrollLockRef.current) return;
        scrollLockRef.current = true;
        const deltaY = normalizeDeltaY(e);
        const next1 = Math.max(0, Math.min(el1.scrollTop + deltaY, el1.scrollHeight - el1.clientHeight));
        const next2 = Math.max(0, Math.min(el2.scrollTop + deltaY, el2.scrollHeight - el2.clientHeight));
        el1.scrollTop = next1;
        el2.scrollTop = next2;
        window.requestAnimationFrame(() => (scrollLockRef.current = false));
      };

      el1.addEventListener("wheel", onWheel1, { passive: false });
      el2.addEventListener("wheel", onWheel2, { passive: false });
    }

    return () => {
      el1.removeEventListener("scroll", onScroll1);
      el2.removeEventListener("scroll", onScroll2);
      if (onWheel1) el1.removeEventListener("wheel", onWheel1);
      if (onWheel2) el2.removeEventListener("wheel", onWheel2);
    };

  }, [result, syncScroll, syncMode]);

  async function runPreviewComparison(event) {
    event.preventDefault();

    if (!file1 || !file2) {
      setError("Please upload both documents.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const formData = new FormData();
      formData.append("file1", file1);
      formData.append("file2", file2);

      const response = await fetch(`${API_BASE}/compare-preview?max_pages=0&include_images=true`, {
        method: "POST",
        body: formData,
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.detail || "Preview request failed.");
      }

      setResult(data);
    } catch (requestError) {
      setResult(null);
      setError(getErrorMessage(requestError));
    } finally {
      setLoading(false);
    }
  }

  async function downloadComparedFiles() {
    if (!file1 || !file2) {
      setError("Please upload both documents before downloading.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const formData = new FormData();
      formData.append("file1", file1);
      formData.append("file2", file2);

      const response = await fetch(`${API_BASE}/compare`, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let text = "";
        try {
          text = await response.text();
        } catch (e) {
          /* ignore */
        }
        console.error("Download request failed", response.status, text);
        throw new Error(`Download failed (status ${response.status})`);
      }

      const blob = await response.blob();
      const blobUrl = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = blobUrl;
      anchor.download = "comparison_result.zip";
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      URL.revokeObjectURL(blobUrl);
    } catch (requestError) {
      setError(getErrorMessage(requestError));
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="page">
      <header className="hero">
        <p className="eyebrow">Deterministic Document Diff</p>
        <h1>Document Comparison Studio</h1>
        <p>
          This comparison is deterministic and rule-based. No AI is used in matching or highlighting.
        </p>
        <p>
          Highlights indicate changed or added words and sections, and for PDFs the highlighted pages show where changes appear.
        </p>
      </header>

      <form className="uploader-grid" onSubmit={runPreviewComparison}>
        <section className="upload-box">
          <h2>Document 1</h2>
          <p className="upload-help">Upload PDF or DOCX</p>
          <input
            type="file"
            accept=".pdf,.docx"
            onChange={(e) => setFile1(e.target.files?.[0] || null)}
          />
          <p className="filename">{file1 ? file1.name : "No file selected"}</p>
        </section>

        <section className="upload-box">
          <h2>Document 2</h2>
          <p className="upload-help">Upload PDF or DOCX</p>
          <input
            type="file"
            accept=".pdf,.docx"
            onChange={(e) => setFile2(e.target.files?.[0] || null)}
          />
          <p className="filename">{file2 ? file2.name : "No file selected"}</p>
        </section>

        <section className="controls">
          <p className="muted-note">Preview loads all pages.</p>

          <div className="button-row">
            <button type="submit" disabled={!canCompare}>
              {loading ? "Generating Preview..." : "Run Preview Comparison"}
            </button>
            <button
              type="button"
              className="secondary"
              disabled={!file1 || !file2 || loading}
              onClick={downloadComparedFiles}
            >
              Download Highlighted Files
            </button>
          </div>
        </section>
      </form>

      {error ? <p className="error">{error}</p> : null}

      {result ? (
        <section className="summary">
          <h2>Comparison Summary</h2>
          <p className="summary-note">
            Highlights are based on changed or added text segments. Total highlighted changes: {result.summary.highlighted_changes}.
          </p>
          <div className="stats-grid">
            <article>
              <span>Total Words Doc 1</span>
              <strong>{result.summary.total_words1}</strong>
            </article>
            <article>
              <span>Total Words Doc 2</span>
              <strong>{result.summary.total_words2}</strong>
            </article>
            <article>
              <span>Highlighted Changes Doc 1</span>
              <strong>{result.summary.diff_words1}</strong>
            </article>
            <article>
              <span>Highlighted Changes Doc 2</span>
              <strong>{result.summary.diff_words2}</strong>
            </article>
            <article>
              <span>Match Rate</span>
              <strong>{result.summary.match_rate}%</strong>
            </article>
          </div>
        </section>
      ) : null}

      {result ? (
        <>
          <div className="preview-toolbar">
            <div>
              <h2>Document Preview</h2>
              <p className="summary-note">Use the arrows to move both documents page by page. Scroll sync keeps the same page index aligned.</p>
            </div>
            <div className="preview-nav">
              {syncMode === "page" ? (
                <>
                  <button
                    type="button"
                    className="secondary"
                    style={{ minWidth: 110 }}
                    onClick={() => {
                      const nextPage = Math.max(currentPageIndex - 1, 0);
                      setCurrentPageIndex(nextPage);
                      syncBothPreviewsToPage(nextPage);
                    }}
                    disabled={loading || currentPageIndex <= 0 || !syncScroll}
                  >
                    ← Prev Page
                  </button>
                  <button
                    type="button"
                    className="secondary"
                    style={{ minWidth: 110 }}
                    onClick={() => {
                      const maxPageIndex = Math.max(getPageCount(result.doc1), getPageCount(result.doc2)) - 1;
                      const nextPage = Math.min(currentPageIndex + 1, Math.max(maxPageIndex, 0));
                      setCurrentPageIndex(nextPage);
                      syncBothPreviewsToPage(nextPage);
                    }}
                    disabled={loading || !syncScroll}
                  >
                    Next Page →
                  </button>
                </>
              ) : null}
              <div className="sync-mode-group">
                <button type="button" className={syncMode === "mouse" ? "sync-mode-button active" : "sync-mode-button"} onClick={() => setSyncMode("mouse")}>
                  Mouse sync
                </button>
                <button type="button" className={syncMode === "page" ? "sync-mode-button active" : "sync-mode-button"} onClick={() => setSyncMode("page")}>
                  Page sync
                </button>
              </div>
              <label className="sync-switch">
                <input
                  type="checkbox"
                  checked={syncScroll}
                  onChange={(event) => toggleSync(event.target.checked)}
                />
                <span className="sync-switch-track">
                  <span className="sync-switch-thumb" />
                </span>
                <span className="sync-switch-label">
                  {syncScroll ? `${syncMode === "mouse" ? "Mouse" : "Page"} sync on` : `${syncMode === "mouse" ? "Mouse" : "Page"} sync off`}
                </span>
              </label>
              {syncMode === "page" ? (
                <span className="page-indicator">
                  Page {currentPageIndex + 1} / {Math.max(getPageCount(result.doc1), getPageCount(result.doc2))}
                </span>
              ) : null}
            </div>
          </div>
          <section className="preview-grid">
            <DocPreview
              doc={result.doc1}
              title="Preview: Document 1"
              scrollRef={doc1ScrollRef}
            />
            <DocPreview
              doc={result.doc2}
              title="Preview: Document 2"
              scrollRef={doc2ScrollRef}
            />
          </section>
        </>
      ) : null}

      {loading && !result ? <LoadingPreview /> : null}
    </main>
  );
}
