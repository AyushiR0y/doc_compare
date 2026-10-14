"""Builds the side-by-side document viewer embedded in the Streamlit app.

Streamlit cannot synchronise scrolling between two of its own containers, so both
documents, the toolbar and all the interaction live in one self-contained HTML
document rendered in a single iframe.

PDF pages arrive as rendered images. DOCX arrives as one continuous HTML flow,
which is then paginated in the browser - blocks are measured and distributed
across page-shaped sheets - so a Word file reads like the PDF beside it and page
stepping works for both. Zoom scales the page width and the type together, so the
layout is self-similar and pagination holds steady as you zoom.
"""

from __future__ import annotations

import base64
import json
from typing import Any

BRAND = "#005eac"
A4_ASPECT = 1.414  # height / width, used when there is no PDF to match


def _page_aspect(doc: dict[str, Any]) -> float | None:
    """Height/width of a PDF document's first page, if it is a PDF."""
    preview = doc.get("preview") or {}
    if preview.get("type") != "pdf_images":
        return None
    for page in preview.get("pages") or []:
        width = page.get("width_px")
        height = page.get("height_px")
        if width and height:
            return round(height / width, 4)
    return None


def _pane_body(doc: dict[str, Any]) -> str:
    """Render one document's pages."""
    preview = doc.get("preview") or {}

    if preview.get("type") == "html":
        # One flow now; the browser slices it into page sheets on load.
        return (
            '<div class="pg docsrc"><div class="sheet doc">'
            + preview.get("html", "")
            + "</div></div>"
        )

    chunks: list[str] = []
    for page in preview.get("pages") or []:
        encoded = page.get("image_base64")
        if not encoded:
            continue

        width = page.get("width_px") or 1
        height = page.get("height_px") or 1

        # Change anchors are positioned in percentages so they survive zooming.
        anchors = "".join(
            '<i class="chg" style="left:{:.3f}%;top:{:.3f}%;width:{:.3f}%;height:{:.3f}%"></i>'.format(
                x0 / width * 100,
                y0 / height * 100,
                max(x1 - x0, 1) / width * 100,
                max(y1 - y0, 1) / height * 100,
            )
            for x0, y0, x1, y1 in page.get("changes") or []
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
            f'alt="Page {page["page"]}"/>{anchors}</div>'
            f'<div class="pgcap">Page {page["page"]}{badge}</div></div>'
        )

    return "".join(chunks) or '<div class="pg empty">No preview available.</div>'


_CSS = """
*{box-sizing:border-box}
html,body{margin:0;padding:0;height:100%;overflow:hidden}
body{font-family:Inter,"Segoe UI",system-ui,Arial,sans-serif;background:#eef4fb;color:#0f172a}
.dcv{display:flex;flex-direction:column;height:100vh}

/* ---------- toolbar ---------- */
.tbar{display:flex;align-items:center;gap:9px;flex-wrap:wrap;
  padding:7px 11px;background:#fff;border:1px solid #dde8f4;border-radius:11px;
  margin:0 0 9px 0;box-shadow:0 1px 6px rgba(15,23,42,.05);flex:0 0 auto}
.tgroup{display:flex;align-items:center;gap:5px}
.tgroup.sep{border-left:1px solid #e8eff7;padding-left:11px}
.tgroup.right{margin-left:auto}
.tlabel{font-size:10.5px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#8ea2b8}

button{font:inherit;cursor:pointer;border:1px solid #dde8f4;background:#fff;color:#0f172a;
  border-radius:7px;padding:4px 9px;font-size:12px;font-weight:600;line-height:1.35;
  transition:background .12s,border-color .12s,color .12s}
button:hover:not(:disabled){background:#f2f8ff;border-color:#9fc7e8}
button:disabled{opacity:.35;cursor:not-allowed}
button.icon{padding:4px 8px;font-size:12px}

.seg{display:inline-flex;background:#eff5fb;border:1px solid #dde8f4;border-radius:8px;padding:2px;gap:2px}
.seg button{border:none;background:transparent;padding:3px 10px;border-radius:6px;color:#546b85}
.seg button.on{background:#fff;color:__BRAND__;box-shadow:0 1px 3px rgba(15,23,42,.14)}

.sw{width:38px;height:21px;border-radius:999px;background:#cbd5e1;border:none;padding:0;position:relative;
  transition:background .16s}
.sw.on{background:__BRAND__}
.sw .knob{position:absolute;top:3px;left:3px;width:15px;height:15px;border-radius:50%;background:#fff;
  transition:transform .16s;box-shadow:0 1px 3px rgba(0,0,0,.3)}
.sw.on .knob{transform:translateX(17px)}

.ind{font-size:12px;font-weight:700;color:#334155;min-width:72px;text-align:center;
  font-variant-numeric:tabular-nums}
.ind.sm{min-width:56px}
.offset{font-size:10.5px;font-weight:700;color:#92400e;background:#fff7e6;border:1px solid #fcd34d;
  border-radius:999px;padding:2px 7px;white-space:nowrap}

/* ---------- panes ---------- */
.panes{display:grid;grid-template-columns:1fr 1fr;gap:10px;flex:1 1 auto;min-height:0}
.panes.stacked{grid-template-columns:1fr;grid-template-rows:1fr 1fr}
.pane{display:flex;flex-direction:column;min-height:0;min-width:0;background:#fff;
  border:1px solid #dde8f4;border-radius:11px;overflow:hidden;box-shadow:0 1px 6px rgba(15,23,42,.05)}
.pane.active{border-color:__BRAND__;box-shadow:0 0 0 2px rgba(0,94,172,.13)}
.phead{display:flex;align-items:center;gap:8px;padding:7px 11px;border-bottom:1px solid #eaf1f8;
  background:linear-gradient(180deg,#fff,#f9fcff);flex:0 0 auto}
.ptitle{display:flex;align-items:center;gap:7px;min-width:0;flex:1 1 auto;cursor:pointer}
.ext{font-size:9.5px;font-weight:800;letter-spacing:.05em;color:#fff;background:__BRAND__;
  border-radius:4px;padding:2px 6px;flex:0 0 auto}
.pname{font-size:12.5px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.pmeta{font-size:11px;color:#8ea2b8;white-space:nowrap;flex:0 0 auto}
.pzoom{display:flex;align-items:center;gap:3px;flex:0 0 auto}
.pzoom .ind{min-width:42px;font-size:11px}

.scroll{flex:1 1 auto;overflow:auto;background:#eef4fb;padding:12px;overscroll-behavior:contain}
.pg{width:calc(100% * var(--z,1));margin:0 auto 14px}
.pg:last-child{margin-bottom:4px}
.sheet{position:relative;background:#fff;border:1px solid #d8e4f2;border-radius:4px;overflow:hidden;
  box-shadow:0 2px 10px rgba(15,23,42,.1)}
.sheet img{display:block;width:100%;height:auto}
.pgcap{display:flex;align-items:center;gap:8px;justify-content:center;
  font-size:11px;color:#8ea2b8;margin-top:5px}
.badge{background:#fff7e6;border:1px solid #fcd34d;color:#92400e;border-radius:999px;
  padding:1px 7px;font-size:10px;font-weight:700}
.badge.none{background:#f1f5f9;border-color:#e8edf3;color:#aebcca}
.pg.empty{padding:28px;text-align:center;color:#94a3b8;font-size:13px}

/* ---------- DOCX sheets, styled to read like a printed page ----------
   Sizes are in em against a root that scales with --z, so zooming enlarges the
   type and the page together instead of just reflowing the text. */
.sheet.doc{padding:7.2% 8%;overflow:hidden}
.sheet.doc .dh{font-family:Georgia,"Times New Roman",serif;padding:0;color:#1a1a1a;
  font-size:calc(10.5px * var(--z,1));line-height:1.52}
.sheet.doc .dh p{margin:0 0 .62em;text-align:justify;hyphens:auto}
.sheet.doc .doc-h{font-size:1.16em;font-weight:700;margin:1.1em 0 .45em;color:#111;line-height:1.3}
.sheet.doc .dh > *:first-child{margin-top:0}
.sheet.doc .doc-tbl{font-size:.92em;margin:.8em 0;width:100%;border-collapse:collapse}
.sheet.doc .doc-cell{border:1px solid #d7dde5;padding:.3em .45em;vertical-align:top}
.sheet.doc .dh img{max-width:100%;height:auto}
.docsrc{display:none}

/* change anchors sit over the baked-in highlight; only the current one shows */
.chg{position:absolute;display:block;pointer-events:none;border-radius:2px}
.chg.cur{outline:2.5px solid #ea580c;outline-offset:2px;background:rgba(234,88,12,.16);
  animation:pulse 1.1s ease-out 1}
.dc.cur{outline:2px solid #ea580c;outline-offset:1px;border-radius:2px}
@keyframes pulse{0%{box-shadow:0 0 0 0 rgba(234,88,12,.55)}100%{box-shadow:0 0 0 14px rgba(234,88,12,0)}}

/* Only stack when genuinely too narrow for two columns. A higher breakpoint
   silently stacked the panes on 1280px screens and on 1366px ones at Windows
   125% scaling, which is the whole point of this view. */
@media (max-width:640px){.panes{grid-template-columns:1fr;grid-template-rows:1fr 1fr}}
"""


def build_viewer_html(
    doc1: dict[str, Any],
    doc2: dict[str, Any],
    summary: dict[str, Any],
) -> str:
    # A DOCX takes the page shape of the PDF beside it where there is one, so the
    # two panes line up; otherwise it falls back to A4 proportions.
    aspect1 = _page_aspect(doc1)
    aspect2 = _page_aspect(doc2)
    config = {
        "aspect1": aspect1 or aspect2 or A4_ASPECT,
        "aspect2": aspect2 or aspect1 or A4_ASPECT,
    }

    def head(index: int, doc: dict[str, Any], changed: int) -> str:
        return (
            f'<div class="ptitle" data-pane="{index}">'
            f'<span class="ext">{doc.get("extension", "doc").upper()}</span>'
            f'<span class="pname" title="{doc.get("name", "")}">{doc.get("name", "")}</span></div>'
            f'<span class="pmeta" id="meta{index}">{changed} changed</span>'
            f'<span class="pzoom">'
            f'<button class="icon" data-zoom="out" data-pane="{index}" title="Zoom out">&minus;</button>'
            f'<span class="ind" id="zoomInd{index}">100%</span>'
            f'<button class="icon" data-zoom="in" data-pane="{index}" title="Zoom in">+</button>'
            f'<button data-zoom="reset" data-pane="{index}" title="Reset zoom">Fit</button>'
            f"</span>"
        )

    css = _CSS.replace("__BRAND__", BRAND)

    return f"""<!doctype html><html><head><meta charset="utf-8"><style>{css}</style></head>
<body><div class="dcv">
  <div class="tbar">
    <div class="tgroup">
      <span class="tlabel">Sync</span>
      <button class="sw on" id="syncBtn" title="Toggle synchronised scrolling (S)"><span class="knob"></span></button>
      <div class="seg" id="modeSeg">
        <button data-mode="mouse" class="on" title="The wheel scrolls both panes together">Mouse</button>
        <button data-mode="page" title="Step both panes page by page">Page</button>
      </div>
    </div>
    <div class="tgroup sep">
      <span class="tlabel">Page</span>
      <button class="icon" id="prevPage" title="Previous page (left arrow)">&#9664;</button>
      <span class="ind" id="pageInd">1 / 1</span>
      <button class="icon" id="nextPage" title="Next page (right arrow)">&#9654;</button>
      <button id="alignBtn" title="Treat the pages now on screen as the same page">Align</button>
      <span class="offset" id="offInd" style="display:none"></span>
    </div>
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
      <span class="tlabel">Layout</span>
      <div class="seg" id="layoutSeg">
        <button data-layout="cols" class="on">Columns</button>
        <button data-layout="rows">Rows</button>
      </div>
    </div>
  </div>

  <div class="panes" id="panes">
    <section class="pane active" id="pane1">
      <header class="phead">{head(1, doc1, summary.get("diff_words1", 0))}</header>
      <div class="scroll" id="s1">{_pane_body(doc1)}</div>
    </section>
    <section class="pane" id="pane2">
      <header class="phead">{head(2, doc2, summary.get("diff_words2", 0))}</header>
      <div class="scroll" id="s2">{_pane_body(doc2)}</div>
    </section>
  </div>
</div>
<script>
(function(){{
  var CFG = {json.dumps(config)};
  var s1 = document.getElementById('s1'), s2 = document.getElementById('s2');
  var lock = false, sync = true, mode = 'mouse';
  var page = 0, offset = 0, activePane = 1, chgIdx = -1;
  var zoom = {{1: 1, 2: 1}};

  function pagesOf(el){{ return Array.prototype.slice.call(el.querySelectorAll('.pg')); }}
  function changesOf(el){{
    return Array.prototype.slice.call(
      el.querySelectorAll('.pg:not(.docsrc) .chg, .pg:not(.docsrc) .dc'));
  }}
  function clamp(v, lo, hi){{ return Math.max(lo, Math.min(v, hi)); }}
  function maxScroll(el){{ return Math.max(0, el.scrollHeight - el.clientHeight); }}
  function paneEl(n){{ return n === 1 ? s1 : s2; }}

  /* ---------- DOCX pagination ----------
     The DOCX arrives as one flow. Measure its top-level blocks against a page
     box and deal them onto sheets, breaking only between blocks so no line is
     ever cut in half. A block taller than a page keeps its own sheet and the
     sheet grows, so nothing is clipped. */
  var docSource = {{}};   /* pane -> detached .dh holding the original blocks */

  function captureSource(n){{
    var el = paneEl(n), src = el.querySelector('.docsrc .dh');
    if (src) docSource[n] = src.cloneNode(true);
  }}

  function paginateDoc(n){{
    if (!docSource[n]) return;
    var el = paneEl(n);
    var holder = el.querySelector('.docsrc');
    if (!holder) return;

    /* clear previously built pages, keep the hidden source holder */
    pagesOf(el).forEach(function(p){{ if (!p.classList.contains('docsrc')) p.remove(); }});

    var aspect = n === 1 ? CFG.aspect1 : CFG.aspect2;
    var blocks = Array.prototype.slice.call(docSource[n].cloneNode(true).children);

    function newPage(){{
      var pg = document.createElement('div');
      pg.className = 'pg';
      var sheet = document.createElement('div');
      sheet.className = 'sheet doc';
      var dh = document.createElement('div');
      dh.className = 'dh';
      sheet.appendChild(dh);
      pg.appendChild(sheet);
      var cap = document.createElement('div');
      cap.className = 'pgcap';
      pg.appendChild(cap);
      el.insertBefore(pg, holder);
      return {{pg: pg, sheet: sheet, dh: dh, cap: cap}};
    }}

    var cur = newPage();
    /* content height available on a page: the sheet is a page box minus its
       padding, which is a percentage of the width, so derive both from width */
    var pageW = cur.pg.clientWidth || el.clientWidth;
    var padY = pageW * 0.072 * 2;
    var limit = Math.max(pageW * aspect - padY, 120);

    blocks.forEach(function(block){{
      cur.dh.appendChild(block);
      if (cur.dh.scrollHeight > limit && cur.dh.children.length > 1) {{
        cur.dh.removeChild(block);          /* this block starts the next page */
        cur = newPage();
        cur.dh.appendChild(block);
      }}
    }});

    /* number the pages and badge their change counts */
    var built = pagesOf(el).filter(function(p){{ return !p.classList.contains('docsrc'); }});
    built.forEach(function(pg, i){{
      pg.setAttribute('data-page', i + 1);
      var sheet = pg.querySelector('.sheet');
      if (sheet) sheet.style.minHeight = (limit + padY) + 'px';
      var c = pg.querySelectorAll('.dc').length;
      pg.querySelector('.pgcap').innerHTML =
        'Page ' + (i + 1) +
        (c ? '<span class="badge">' + c + ' change' + (c === 1 ? '' : 's') + '</span>'
           : '<span class="badge none">no changes</span>');
    }});
  }}

  function repaginateAll(){{
    [1, 2].forEach(function(n){{ if (docSource[n]) paginateDoc(n); }});
    renderPage(); renderChg(); renderMeta();
  }}

  function renderMeta(){{
    [1, 2].forEach(function(n){{
      var el = paneEl(n), m = document.getElementById('meta' + n);
      if (!m) return;
      var count = pagesOf(el).filter(function(p){{
        return !p.classList.contains('docsrc'); }}).length;
      var changed = changesOf(el).length;
      m.textContent = count + ' page' + (count === 1 ? '' : 's') + ' \\u00b7 ' + changed + ' changed';
    }});
  }}

  /* ---------- page indicator ---------- */
  function realPages(el){{
    return pagesOf(el).filter(function(p){{ return !p.classList.contains('docsrc'); }});
  }}
  function pageIndexOf(el){{
    var ps = realPages(el), sel = 0;
    for (var i = 0; i < ps.length; i++) {{
      if (ps[i].offsetTop <= el.scrollTop + 24) sel = i; else break;
    }}
    return sel;
  }}
  function totalPages(){{ return Math.max(realPages(s1).length, realPages(s2).length, 1); }}
  function renderPage(){{
    var total = totalPages();
    page = clamp(page, 0, total - 1);
    document.getElementById('pageInd').textContent = (page + 1) + ' / ' + total;
    document.getElementById('prevPage').disabled = page <= 0;
    document.getElementById('nextPage').disabled = page >= total - 1;
  }}
  function goToPage(idx){{
    page = clamp(idx, 0, totalPages() - 1);
    lock = true;
    var p1 = realPages(s1), p2 = realPages(s2);
    if (p1[page]) s1.scrollTop = p1[page].offsetTop - 8;
    if (sync) {{
      var t = clamp(page + offset, 0, p2.length - 1);
      if (p2[t]) s2.scrollTop = p2[t].offsetTop - 8;
    }}
    requestAnimationFrame(function(){{ lock = false; }});
    renderPage();
  }}

  /* ---------- page offset ----------
     Pages line up 1:1 by default. Align is the only thing that sets an offset,
     and it then shows in the toolbar, so it is never a silent surprise. */
  function captureOffset(){{ offset = pageIndexOf(s2) - pageIndexOf(s1); renderOffset(); }}
  function resetOffset(){{ offset = 0; renderOffset(); }}
  function renderOffset(){{
    var el = document.getElementById('offInd');
    el.textContent = offset ? ('offset ' + (offset > 0 ? '+' : '') + offset) : '';
    el.style.display = offset ? '' : 'none';
  }}

  /* ---------- change navigation ---------- */
  function renderChg(){{
    var list = changesOf(paneEl(activePane));
    var pos = !list.length ? '0' : (chgIdx < 0 ? '\\u2013' : (clamp(chgIdx, 0, list.length - 1) + 1));
    document.getElementById('chgInd').textContent = pos + ' / ' + list.length;
    document.getElementById('prevChg').disabled = !list.length;
    document.getElementById('nextChg').disabled = !list.length;
  }}
  function goToChange(step){{
    var el = paneEl(activePane), list = changesOf(el);
    if (!list.length) return;
    chgIdx = chgIdx < 0 ? (step > 0 ? 0 : list.length - 1) : clamp(chgIdx + step, 0, list.length - 1);
    list.forEach(function(n){{ n.classList.remove('cur'); }});
    var target = list[chgIdx];
    target.classList.add('cur');
    target.style.animation = 'none'; void target.offsetWidth; target.style.animation = '';

    var before = el.scrollTop, top = 0, node = target;
    while (node && node !== el) {{ top += node.offsetTop; node = node.offsetParent; }}
    var next = clamp(top - el.clientHeight / 2, 0, maxScroll(el));
    lock = true;
    el.scrollTop = next;
    if (sync) {{
      var other = el === s1 ? s2 : s1;
      other.scrollTop = clamp(other.scrollTop + (next - before), 0, maxScroll(other));
    }}
    requestAnimationFrame(function(){{ lock = false; }});
    renderChg();
  }}

  /* ---------- wheel sync ---------- */
  function wheel(e){{
    if (!sync || mode !== 'mouse') return;      /* let the pane scroll on its own */
    e.preventDefault();
    if (lock) return;
    lock = true;
    var d = e.deltaMode === 1 ? e.deltaY * 16
          : e.deltaMode === 2 ? e.deltaY * s1.clientHeight : e.deltaY;
    s1.scrollTop = clamp(s1.scrollTop + d, 0, maxScroll(s1));
    s2.scrollTop = clamp(s2.scrollTop + d, 0, maxScroll(s2));
    requestAnimationFrame(function(){{ lock = false; }});
    page = pageIndexOf(s1); renderPage();
  }}
  s1.addEventListener('wheel', wheel, {{passive:false}});
  s2.addEventListener('wheel', wheel, {{passive:false}});
  s1.addEventListener('scroll', function(){{ if (!lock) {{ page = pageIndexOf(s1); renderPage(); }} }}, {{passive:true}});

  /* ---------- controls ---------- */
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
  Array.prototype.forEach.call(document.querySelectorAll('.ptitle'), function(h){{
    h.addEventListener('click', function(){{ setActive(parseInt(h.dataset.pane, 10)); }});
  }});

  document.getElementById('prevPage').addEventListener('click', function(){{ goToPage(page - 1); }});
  document.getElementById('nextPage').addEventListener('click', function(){{ goToPage(page + 1); }});
  document.getElementById('alignBtn').addEventListener('click', function(){{
    if (offset) resetOffset(); else captureOffset();
  }});
  document.getElementById('prevChg').addEventListener('click', function(){{ goToChange(-1); }});
  document.getElementById('nextChg').addEventListener('click', function(){{ goToChange(1); }});

  /* ---------- per-pane zoom ---------- */
  function setZoom(n, z){{
    zoom[n] = clamp(z, 0.5, 3);
    paneEl(n).style.setProperty('--z', zoom[n]);
    document.getElementById('zoomInd' + n).textContent = Math.round(zoom[n] * 100) + '%';
    if (docSource[n]) {{ paginateDoc(n); renderPage(); renderChg(); renderMeta(); }}
  }}
  document.querySelector('.dcv').addEventListener('click', function(e){{
    var b = e.target.closest('[data-zoom]'); if (!b) return;
    var n = parseInt(b.dataset.pane, 10);
    if (b.dataset.zoom === 'in') setZoom(n, zoom[n] + 0.15);
    else if (b.dataset.zoom === 'out') setZoom(n, zoom[n] - 0.15);
    else setZoom(n, 1);
  }});

  document.getElementById('layoutSeg').addEventListener('click', function(e){{
    var b = e.target.closest('button'); if (!b) return;
    document.getElementById('panes').classList.toggle('stacked', b.dataset.layout === 'rows');
    Array.prototype.forEach.call(this.children, function(c){{ c.classList.toggle('on', c === b); }});
    setTimeout(repaginateAll, 60);
  }});

  document.addEventListener('keydown', function(e){{
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
    var k = e.key.toLowerCase();
    if (k === 'arrowright') {{ e.preventDefault(); goToPage(page + 1); }}
    else if (k === 'arrowleft') {{ e.preventDefault(); goToPage(page - 1); }}
    else if (k === 'n') {{ e.preventDefault(); goToChange(1); }}
    else if (k === 'p') {{ e.preventDefault(); goToChange(-1); }}
    else if (k === 's') {{ e.preventDefault(); syncBtn.click(); }}
  }});

  var resizeTimer;
  window.addEventListener('resize', function(){{
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(repaginateAll, 180);
  }});

  [1, 2].forEach(function(n){{
    captureSource(n);
    paneEl(n).style.setProperty('--z', 1);
  }});
  repaginateAll();
}})();
</script></body></html>"""


def viewer_data_url(html: str) -> str:
    """Wrap the viewer document as a data URL for st.iframe."""
    return "data:text/html;base64," + base64.b64encode(html.encode("utf-8")).decode("ascii")
