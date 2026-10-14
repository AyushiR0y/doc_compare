"""Builds the side-by-side document viewer embedded in the Streamlit app.

Streamlit cannot synchronise scrolling between two of its own containers, so both
documents, the toolbar and all the interaction live in one self-contained HTML
document rendered in a single iframe. It carries over what the React build does
(sync toggle, mouse and page sync modes, page stepping, page indicator) and adds
zoom, change-to-change navigation and keyboard shortcuts.
"""

from __future__ import annotations

import base64
import json
from typing import Any

BRAND = "#005eac"


def _pane_body(doc: dict[str, Any]) -> tuple[str, int]:
    """Render one document's pages. Returns (html, page_count)."""
    preview = doc.get("preview") or {}

    if preview.get("type") == "html":
        # DOCX renders as one continuous sheet; changed words carry class "dc".
        return (
            '<div class="pg" data-page="1"><div class="sheet doc">'
            + preview.get("html", "")
            + "</div></div>",
            1,
        )

    pages = preview.get("pages") or []
    chunks: list[str] = []

    for page in pages:
        encoded = page.get("image_base64")
        if not encoded:
            continue

        width = page.get("width_px") or 1
        height = page.get("height_px") or 1

        # Change anchors are positioned in percentages so they survive zooming.
        anchors = []
        for x0, y0, x1, y1 in page.get("changes") or []:
            anchors.append(
                '<i class="chg" style="left:{:.3f}%;top:{:.3f}%;width:{:.3f}%;height:{:.3f}%"></i>'.format(
                    x0 / width * 100,
                    y0 / height * 100,
                    max(x1 - x0, 1) / width * 100,
                    max(y1 - y0, 1) / height * 100,
                )
            )

        count = page.get("change_count", 0)
        badge = (
            f'<span class="badge">{count} change{"s" if count != 1 else ""}</span>'
            if count
            else '<span class="badge none">no changes</span>'
        )

        chunks.append(
            f'<div class="pg" data-page="{page["page"]}">'
            f'<div class="sheet" style="aspect-ratio:{width}/{height}">'
            f'<img loading="lazy" src="data:{page.get("mime_type", "image/jpeg")};base64,{encoded}" '
            f'alt="Page {page["page"]}"/>{"".join(anchors)}</div>'
            f'<div class="pgcap">Page {page["page"]}{badge}</div></div>'
        )

    if not chunks:
        chunks.append('<div class="pg empty">No preview available.</div>')

    return "".join(chunks), len(pages)


_CSS = """
*{box-sizing:border-box}
html,body{margin:0;padding:0;height:100%;overflow:hidden}
body{font-family:Inter,"Segoe UI",system-ui,Arial,sans-serif;background:#eef4fb;color:#0f172a}
.dcv{display:flex;flex-direction:column;height:100vh}

/* ---------- toolbar ---------- */
.tbar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;
  padding:8px 12px;background:#fff;border:1px solid #dde8f4;border-radius:12px;
  margin:0 0 9px 0;box-shadow:0 2px 10px rgba(15,23,42,.05);flex:0 0 auto}
.tgroup{display:flex;align-items:center;gap:6px}
.tgroup.sep{border-left:1px solid #e3ecf7;padding-left:12px}
.tgroup.right{margin-left:auto}
.tlabel{font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#64748b}

button{font:inherit;cursor:pointer;border:1px solid #dde8f4;background:#fff;color:#0f172a;
  border-radius:8px;padding:5px 10px;font-size:12.5px;font-weight:600;line-height:1.3;
  transition:background .12s,border-color .12s,color .12s}
button:hover:not(:disabled){background:#f2f8ff;border-color:#9fc7e8}
button:disabled{opacity:.4;cursor:not-allowed}
button.icon{padding:5px 9px;font-size:13px}

.seg{display:inline-flex;background:#eff5fb;border:1px solid #dde8f4;border-radius:9px;padding:2px;gap:2px}
.seg button{border:none;background:transparent;padding:4px 11px;border-radius:7px;color:#475569}
.seg button.on{background:#FFF;color:__BRAND__;box-shadow:0 1px 4px rgba(15,23,42,.12)}

.sw{width:40px;height:22px;border-radius:999px;background:#cbd5e1;border:none;padding:0;position:relative;
  transition:background .16s}
.sw.on{background:__BRAND__}
.sw .knob{position:absolute;top:3px;left:3px;width:16px;height:16px;border-radius:50%;background:#fff;
  transition:transform .16s;box-shadow:0 1px 3px rgba(0,0,0,.3)}
.sw.on .knob{transform:translateX(18px)}

.ind{font-size:12.5px;font-weight:700;color:#334155;min-width:84px;text-align:center;
  font-variant-numeric:tabular-nums}
.ind.sm{min-width:62px}
.hint{font-size:11.5px;color:#94a3b8}
.offset{font-size:11px;font-weight:700;color:#92400e;background:#fff7e6;border:1px solid #fcd34d;
  border-radius:999px;padding:2px 8px;white-space:nowrap}

/* ---------- panes ---------- */
.panes{display:grid;grid-template-columns:1fr 1fr;gap:10px;flex:1 1 auto;min-height:0}
.panes.stacked{grid-template-columns:1fr;grid-template-rows:1fr 1fr}
.pane{display:flex;flex-direction:column;min-height:0;min-width:0;background:#fff;
  border:1px solid #dde8f4;border-radius:12px;overflow:hidden;box-shadow:0 2px 10px rgba(15,23,42,.05)}
.pane.active{border-color:__BRAND__;box-shadow:0 0 0 2px rgba(0,94,172,.14)}
.phead{display:flex;align-items:center;gap:8px;padding:8px 12px;border-bottom:1px solid #e8f0f9;
  background:linear-gradient(180deg,#fff,#f8fbff);flex:0 0 auto;cursor:pointer}
.ext{font-size:10px;font-weight:800;letter-spacing:.05em;color:#fff;background:__BRAND__;
  border-radius:4px;padding:2px 6px;flex:0 0 auto}
.pname{font-size:13px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1 1 auto}
.pmeta{font-size:11.5px;color:#64748b;white-space:nowrap;flex:0 0 auto}

.scroll{flex:1 1 auto;overflow:auto;background:#eef4fb;padding:12px;scroll-behavior:auto;
  overscroll-behavior:contain}
.pg{width:calc(100% * var(--z,1));margin:0 auto 14px;max-width:none}
.pg:last-child{margin-bottom:4px}
.sheet{position:relative;background:#fff;border:1px solid #d8e4f2;border-radius:6px;overflow:hidden;
  box-shadow:0 3px 12px rgba(15,23,42,.1)}
.sheet img{display:block;width:100%;height:auto}
.sheet.doc{padding:4px 10px;min-height:120px}
.pgcap{display:flex;align-items:center;gap:8px;justify-content:center;
  font-size:11.5px;color:#64748b;margin-top:5px}
.badge{background:#fff7e6;border:1px solid #fcd34d;color:#92400e;border-radius:999px;
  padding:1px 8px;font-size:10.5px;font-weight:700}
.badge.none{background:#f1f5f9;border-color:#e2e8f0;color:#94a3b8}
.pg.empty{padding:28px;text-align:center;color:#94a3b8;font-size:13px}

/* change anchors sit over the baked-in highlight; only the current one shows */
.chg{position:absolute;display:block;pointer-events:none;border-radius:2px}
.chg.cur{outline:2.5px solid #ea580c;outline-offset:2px;background:rgba(234,88,12,.16);
  animation:pulse 1.1s ease-out 1}
.dc.cur{outline:2.5px solid #ea580c;outline-offset:1px;border-radius:2px}
@keyframes pulse{0%{box-shadow:0 0 0 0 rgba(234,88,12,.55)}100%{box-shadow:0 0 0 14px rgba(234,88,12,0)}}

@media (max-width:900px){.panes{grid-template-columns:1fr;grid-template-rows:1fr 1fr}}
"""


def build_viewer_html(
    doc1: dict[str, Any],
    doc2: dict[str, Any],
    summary: dict[str, Any],
    stacked: bool = False,
) -> str:
    body1, pages1 = _pane_body(doc1)
    body2, pages2 = _pane_body(doc2)

    paginated = (doc1.get("preview", {}).get("type") == "pdf_images"
                 and doc2.get("preview", {}).get("type") == "pdf_images")

    config = {
        "pages1": pages1,
        "pages2": pages2,
        "paginated": paginated,
        "changes1": summary.get("diff_words1", 0),
        "changes2": summary.get("diff_words2", 0),
    }

    def head(doc: dict[str, Any], changed: int, pages: int) -> str:
        kind = "page" if doc.get("preview", {}).get("type") == "pdf_images" else "document"
        meta = f"{pages} {kind}{'s' if pages != 1 else ''} · {changed} changed"
        return (
            f'<span class="ext">{doc.get("extension", "doc").upper()}</span>'
            f'<span class="pname" title="{doc.get("name", "")}">{doc.get("name", "")}</span>'
            f'<span class="pmeta">{meta}</span>'
        )

    page_nav = (
        '<div class="tgroup sep" id="pageGrp">'
        '<span class="tlabel">Page</span>'
        '<button class="icon" id="prevPage" title="Previous page (←)">&#9664;</button>'
        '<span class="ind" id="pageInd">1 / 1</span>'
        '<button class="icon" id="nextPage" title="Next page (→)">&#9654;</button>'
        '<button id="alignBtn" title="Treat the pages now on screen as the same page">Align</button>'
        '<span class="offset" id="offInd" style="display:none"></span>'
        "</div>"
        if paginated
        else '<div class="tgroup sep"><span class="hint">Page sync needs two PDFs</span></div>'
    )

    css = _CSS.replace("__BRAND__", BRAND)

    return f"""<!doctype html><html><head><meta charset="utf-8"><style>{css}</style></head>
<body><div class="dcv">
  <div class="tbar">
    <div class="tgroup">
      <span class="tlabel">Sync</span>
      <button class="sw on" id="syncBtn" title="Toggle synchronised scrolling (S)"><span class="knob"></span></button>
      <div class="seg" id="modeSeg">
        <button data-mode="mouse" class="on" title="Wheel scrolls both panes together">Mouse</button>
        <button data-mode="page" title="Step both panes page by page">Page</button>
      </div>
    </div>
    {page_nav}
    <div class="tgroup sep">
      <span class="tlabel">Change</span>
      <div class="seg" id="paneSeg">
        <button data-pane="1" class="on">Doc 1</button>
        <button data-pane="2">Doc 2</button>
      </div>
      <button class="icon" id="prevChg" title="Previous change (P)">&#9650;</button>
      <span class="ind sm" id="chgInd">0 / 0</span>
      <button class="icon" id="nextChg" title="Next change (N)">&#9660;</button>
    </div>
    <div class="tgroup sep right">
      <span class="tlabel">Zoom</span>
      <button class="icon" id="zoomOut">&minus;</button>
      <span class="ind sm" id="zoomInd">100%</span>
      <button class="icon" id="zoomIn">+</button>
      <button id="zoomReset">Reset</button>
    </div>
  </div>

  <div class="panes{' stacked' if stacked else ''}" id="panes">
    <section class="pane active" id="pane1">
      <header class="phead" data-pane="1">{head(doc1, config["changes1"], pages1)}</header>
      <div class="scroll" id="s1">{body1}</div>
    </section>
    <section class="pane" id="pane2">
      <header class="phead" data-pane="2">{head(doc2, config["changes2"], pages2)}</header>
      <div class="scroll" id="s2">{body2}</div>
    </section>
  </div>
</div>
<script>
(function(){{
  var CFG = {json.dumps(config)};
  var s1 = document.getElementById('s1'), s2 = document.getElementById('s2');
  var lock = false, sync = true, mode = 'mouse', zoom = 1;
  var page = 0, offset = 0, activePane = 1, chgIdx = -1;

  function pagesOf(el){{ return Array.prototype.slice.call(el.querySelectorAll('.pg')); }}
  function changesOf(el){{ return Array.prototype.slice.call(el.querySelectorAll('.chg,.dc')); }}
  function clamp(v, lo, hi){{ return Math.max(lo, Math.min(v, hi)); }}
  function maxScroll(el){{ return Math.max(0, el.scrollHeight - el.clientHeight); }}

  /* ---------- page indicator ---------- */
  function pageIndexOf(el){{
    var ps = pagesOf(el), sel = 0;
    for (var i = 0; i < ps.length; i++) {{
      if (ps[i].offsetTop <= el.scrollTop + 24) sel = i; else break;
    }}
    return sel;
  }}
  function renderPage(){{
    var ind = document.getElementById('pageInd');
    if (!ind) return;
    var total = Math.max(CFG.pages1, CFG.pages2);
    ind.textContent = (page + 1) + ' / ' + total;
    var pp = document.getElementById('prevPage'), np = document.getElementById('nextPage');
    if (pp) pp.disabled = page <= 0;
    if (np) np.disabled = page >= total - 1;
  }}
  function goToPage(idx){{
    var total = Math.max(CFG.pages1, CFG.pages2);
    page = clamp(idx, 0, total - 1);
    lock = true;
    var p1 = pagesOf(s1), p2 = pagesOf(s2);
    if (p1[page]) s1.scrollTop = p1[page].offsetTop - 8;
    if (sync) {{
      var t = clamp(page + offset, 0, p2.length - 1);
      if (p2[t]) s2.scrollTop = p2[t].offsetTop - 8;
    }}
    requestAnimationFrame(function(){{ lock = false; }});
    renderPage();
  }}

  /* ---------- change navigation ---------- */
  function activeEl(){{ return activePane === 1 ? s1 : s2; }}
  function renderChg(){{
    var list = changesOf(activeEl());
    /* Show no position until a change has actually been stepped to, otherwise
       the counter reads "1 / n" at rest and the first click looks like a no-op. */
    var pos = !list.length ? '0' : (chgIdx < 0 ? '\\u2013' : (clamp(chgIdx, 0, list.length - 1) + 1));
    document.getElementById('chgInd').textContent = pos + ' / ' + list.length;
    document.getElementById('prevChg').disabled = !list.length;
    document.getElementById('nextChg').disabled = !list.length;
  }}
  function goToChange(step){{
    var el = activeEl(), list = changesOf(el);
    if (!list.length) return;
    chgIdx = chgIdx < 0 ? (step > 0 ? 0 : list.length - 1) : clamp(chgIdx + step, 0, list.length - 1);
    list.forEach(function(n){{ n.classList.remove('cur'); }});
    var target = list[chgIdx];
    target.classList.add('cur');
    /* reflow the animation so repeat visits still pulse */
    target.style.animation = 'none'; void target.offsetWidth; target.style.animation = '';

    var before = el.scrollTop;
    var top = 0, node = target;
    while (node && node !== el) {{ top += node.offsetTop; node = node.offsetParent; }}
    var next = clamp(top - el.clientHeight / 2, 0, maxScroll(el));
    lock = true;
    el.scrollTop = next;
    if (sync) {{ mirror(el === s1 ? s2 : s1, next - before); }}
    requestAnimationFrame(function(){{ lock = false; }});
    renderChg();
  }}
  function mirror(other, delta){{
    other.scrollTop = clamp(other.scrollTop + delta, 0, maxScroll(other));
  }}

  /* ---------- wheel sync ---------- */
  function normDelta(e, el){{
    if (e.deltaMode === 1) return e.deltaY * 16;
    if (e.deltaMode === 2) return e.deltaY * el.clientHeight;
    return e.deltaY;
  }}
  function wheel(e){{
    if (!sync || mode !== 'mouse') return;      /* let the pane scroll on its own */
    e.preventDefault();
    if (lock) return;
    lock = true;
    var d = normDelta(e, s1);
    s1.scrollTop = clamp(s1.scrollTop + d, 0, maxScroll(s1));
    s2.scrollTop = clamp(s2.scrollTop + d, 0, maxScroll(s2));
    requestAnimationFrame(function(){{ lock = false; }});
    page = pageIndexOf(s1); renderPage();
  }}
  s1.addEventListener('wheel', wheel, {{passive:false}});
  s2.addEventListener('wheel', wheel, {{passive:false}});

  s1.addEventListener('scroll', function(){{ if (!lock) {{ page = pageIndexOf(s1); renderPage(); }} }}, {{passive:true}});
  s2.addEventListener('scroll', function(){{ if (!lock && !sync) {{ renderPage(); }} }}, {{passive:true}});

  /* ---------- controls ---------- */
  /* Pages line up 1:1 by default - page 3 sits next to page 3. Two documents
     that sit a page apart (an extra cover page, say) can be aligned with the
     Align button, which is the only thing that sets an offset; it then shows in
     the toolbar. Capturing it automatically would silently lock in whatever
     misalignment the panes happened to have. */
  function captureOffset(){{
    offset = pageIndexOf(s2) - pageIndexOf(s1);
    renderOffset();
  }}
  function renderOffset(){{
    var el = document.getElementById('offInd');
    if (!el) return;
    el.textContent = offset ? ('offset ' + (offset > 0 ? '+' : '') + offset) : '';
    el.style.display = offset ? '' : 'none';
  }}
  function resetOffset(){{ offset = 0; renderOffset(); }}

  var syncBtn = document.getElementById('syncBtn');
  syncBtn.addEventListener('click', function(){{
    sync = !sync;
    syncBtn.classList.toggle('on', sync);
  }});

  document.getElementById('modeSeg').addEventListener('click', function(e){{
    var b = e.target.closest('button'); if (!b) return;
    mode = b.dataset.mode;
    Array.prototype.forEach.call(this.children, function(c){{ c.classList.toggle('on', c === b); }});
  }});

  document.getElementById('paneSeg').addEventListener('click', function(e){{
    var b = e.target.closest('button'); if (!b) return;
    setActive(parseInt(b.dataset.pane, 10));
  }});
  function setActive(n){{
    activePane = n; chgIdx = -1;
    document.getElementById('pane1').classList.toggle('active', n === 1);
    document.getElementById('pane2').classList.toggle('active', n === 2);
    Array.prototype.forEach.call(document.getElementById('paneSeg').children, function(c){{
      c.classList.toggle('on', parseInt(c.dataset.pane, 10) === n);
    }});
    renderChg();
  }}
  Array.prototype.forEach.call(document.querySelectorAll('.phead'), function(h){{
    h.addEventListener('click', function(){{ setActive(parseInt(h.dataset.pane, 10)); }});
  }});

  var pp = document.getElementById('prevPage'), np = document.getElementById('nextPage');
  if (pp) pp.addEventListener('click', function(){{ goToPage(page - 1); }});
  if (np) np.addEventListener('click', function(){{ goToPage(page + 1); }});

  var alignBtn = document.getElementById('alignBtn');
  if (alignBtn) alignBtn.addEventListener('click', function(){{
    if (offset) {{ resetOffset(); }} else {{ captureOffset(); }}
  }});
  document.getElementById('prevChg').addEventListener('click', function(){{ goToChange(-1); }});
  document.getElementById('nextChg').addEventListener('click', function(){{ goToChange(1); }});

  function setZoom(z){{
    zoom = clamp(z, 0.5, 3);
    document.querySelector('.dcv').style.setProperty('--z', zoom);
    s1.style.setProperty('--z', zoom); s2.style.setProperty('--z', zoom);
    document.getElementById('zoomInd').textContent = Math.round(zoom * 100) + '%';
  }}
  document.getElementById('zoomIn').addEventListener('click', function(){{ setZoom(zoom + 0.15); }});
  document.getElementById('zoomOut').addEventListener('click', function(){{ setZoom(zoom - 0.15); }});
  document.getElementById('zoomReset').addEventListener('click', function(){{ setZoom(1); }});

  document.addEventListener('keydown', function(e){{
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
    var k = e.key.toLowerCase();
    if (k === 'arrowright' && CFG.paginated) {{ e.preventDefault(); goToPage(page + 1); }}
    else if (k === 'arrowleft' && CFG.paginated) {{ e.preventDefault(); goToPage(page - 1); }}
    else if (k === 'n') {{ e.preventDefault(); goToChange(1); }}
    else if (k === 'p') {{ e.preventDefault(); goToChange(-1); }}
    else if (k === 's') {{ e.preventDefault(); syncBtn.click(); }}
  }});

  setZoom(1); renderPage(); renderChg();
}})();
</script></body></html>"""


def viewer_data_url(html: str) -> str:
    """Wrap the viewer document as a data URL for st.iframe."""
    return "data:text/html;base64," + base64.b64encode(html.encode("utf-8")).decode("ascii")
