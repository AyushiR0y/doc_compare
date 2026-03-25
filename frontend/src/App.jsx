import React, { useMemo, useState } from "react";

const API_BASE = import.meta.env.VITE_API_BASE_URL || "http://127.0.0.1:8000";

function getErrorMessage(error) {
  if (error instanceof Error) {
    return error.message;
  }
  return "Something went wrong while processing the documents.";
}

function DocPreview({ doc, title }) {
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
          className="word-preview"
          dangerouslySetInnerHTML={{ __html: preview.html }}
        />
      ) : (
        <div className="pdf-preview">
          {preview.pages?.map((page) => (
            <figure key={page.page} className="page-figure">
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

  const canCompare = useMemo(() => file1 && file2 && !loading, [file1, file2, loading]);

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
        const data = await response.json();
        throw new Error(data.detail || "Download failed.");
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
          <p className="summary-note">Highlights are based on changed or added text segments.</p>
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
              <span>Changed Words Doc 1</span>
              <strong>{result.summary.diff_words1}</strong>
            </article>
            <article>
              <span>Changed Words Doc 2</span>
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
        <section className="preview-grid">
          <DocPreview doc={result.doc1} title="Preview: Document 1" />
          <DocPreview doc={result.doc2} title="Preview: Document 2" />
        </section>
      ) : null}

      {loading && !result ? <LoadingPreview /> : null}
    </main>
  );
}
