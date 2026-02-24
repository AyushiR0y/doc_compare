import streamlit as st
import difflib
from io import BytesIO
import fitz  # PyMuPDF
from docx import Document
import re
import os
import tempfile
import base64
from PIL import Image
import json
import uuid
from datetime import datetime, timezone
from functools import lru_cache
from urllib import request as urllib_request
from urllib.error import URLError, HTTPError

# UPDATED: Sidebar expanded, layout wide
st.set_page_config(page_title="Document Diff Checker", layout="wide", initial_sidebar_state="expanded")

logo_b64 = ""
try:
    with open("logo.png", "rb") as logo_file:
        logo_b64 = base64.b64encode(logo_file.read()).decode("utf-8")
except OSError:
    logo_b64 = ""

if logo_b64:
    st.markdown(
        f"""
        <div class="brand-header">
            <img src="data:image/png;base64,{logo_b64}" class="brand-logo"/>
            <div class="brand-text">
                <h1>Document Diff Checker</h1>
                <p>Compare PDF and DOCX files with precision highlighting</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    st.markdown("<h1 class='app-title'>Document Diff Checker</h1>", unsafe_allow_html=True)
    st.markdown("<p class='app-subtitle'>Compare PDF and DOCX files with precision highlighting</p>", unsafe_allow_html=True)

st.markdown(
    """
    <style>
    div[data-testid="stRadio"] > label,
    div[data-testid="stRadio"] legend {
        font-size: 1.4rem !important;
        font-weight: 600 !important;
        line-height: 1.3;
    }
    </style>
    """,
    unsafe_allow_html=True
)

comparison_mode = st.radio(
    ":material/tune: Comparison Mode",
    ["Same Format (PDF vs PDF or Word vs Word)", "Mixed Format (PDF vs Word)"],
    horizontal=True
)
# Create two columns for file uploads
col1, col2 = st.columns(2)

with col1:
    st.subheader(":material/description: Document 1")
    if comparison_mode == "Same Format (PDF vs PDF or Word vs Word)":
        doc1_file = st.file_uploader("Select first document", type=['pdf', 'docx'], key="doc1")
    else:
        doc1_file = st.file_uploader("Select first document (PDF or Word)", type=['pdf', 'docx'], key="doc1_mixed")
    
with col2:
    st.subheader(":material/description: Document 2")
    if comparison_mode == "Same Format (PDF vs PDF or Word vs Word)":
        doc2_file = st.file_uploader("Select second document (same format as Document 1)", type=['pdf', 'docx'], key="doc2")
    else:
        doc2_file = st.file_uploader("Select second document (PDF or Word)", type=['pdf', 'docx'], key="doc2_mixed")

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

USAGE_LOG_FILE = "usage_logs.jsonl"


def _safe_get_headers():
    context = getattr(st, "context", None)
    if not context:
        return {}
    headers = getattr(context, "headers", None)
    return dict(headers) if headers else {}


def _extract_client_ip(headers):
    for key in ["x-forwarded-for", "x-real-ip", "cf-connecting-ip", "x-client-ip"]:
        value = headers.get(key) or headers.get(key.title())
        if value:
            return value.split(",")[0].strip()
    return "unknown"


@lru_cache(maxsize=256)
def _resolve_location_from_ip(ip_address):
    if not ip_address or ip_address == "unknown":
        return {
            "country": "unknown",
            "region": "unknown",
            "city": "unknown",
            "source": "none"
        }

    if ip_address.startswith("10.") or ip_address.startswith("192.168.") or ip_address.startswith("172."):
        return {
            "country": "private-network",
            "region": "private-network",
            "city": "private-network",
            "source": "private-ip"
        }

    url = f"https://ipapi.co/{ip_address}/json/"
    try:
        with urllib_request.urlopen(url, timeout=2) as response:
            payload = json.loads(response.read().decode("utf-8"))
        return {
            "country": payload.get("country_name", "unknown") or "unknown",
            "region": payload.get("region", "unknown") or "unknown",
            "city": payload.get("city", "unknown") or "unknown",
            "source": "ipapi.co"
        }
    except (URLError, HTTPError, TimeoutError, json.JSONDecodeError):
        return {
            "country": "unknown",
            "region": "unknown",
            "city": "unknown",
            "source": "lookup-failed"
        }


def _get_session_id():
    if "usage_session_id" not in st.session_state:
        st.session_state.usage_session_id = str(uuid.uuid4())
    return st.session_state.usage_session_id


def log_usage_event(doc1_name, doc2_name, ext1, ext2, comparison_mode):
    headers = _safe_get_headers()
    ip_address = _extract_client_ip(headers)
    location = _resolve_location_from_ip(ip_address)

    event = {
        "event_id": str(uuid.uuid4()),
        "event_type": "comparison",
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "session_id": _get_session_id(),
        "upload_count": 2,
        "comparison_mode": comparison_mode,
        "doc1_name": doc1_name,
        "doc2_name": doc2_name,
        "doc1_type": ext1,
        "doc2_type": ext2,
        "client_ip": ip_address,
        "client_country": location["country"],
        "client_region": location["region"],
        "client_city": location["city"],
        "location_source": location["source"]
    }

    try:
        with open(USAGE_LOG_FILE, "a", encoding="utf-8") as file:
            file.write(json.dumps(event, ensure_ascii=False) + "\n")
    except OSError:
        pass

def truncate_filename(filename, max_length=30):
    """Truncate filename if it exceeds max_length, preserving extension"""
    if len(filename) <= max_length:
        return filename
    
    # Split name and extension
    name_parts = filename.rsplit('.', 1)
    if len(name_parts) == 2:
        name, ext = name_parts
        # Reserve space for extension and ellipsis
        available = max_length - len(ext) - 4  # -4 for "..." and "."
        if available > 0:
            return f"{name[:available]}...{ext}"
        else:
            return f"{filename[:max_length-3]}..."
    else:
        return f"{filename[:max_length-3]}..."

# ============================================================================
# EXTRACTION - FROM CODE 1 FOR PDF, CODE 2 FOR WORD
# ============================================================================

def extract_words_from_word(docx_file):
    """
    FROM CODE 2: Extracts text including <xxxx> tags and handles tables properly
    Returns: (plain_text_string, word_objects, doc_obj)
    """
    try:
        docx_file.seek(0)
        doc = Document(docx_file)
        
        word_objects = []
        text_segments = []
        
        from docx.oxml.text.paragraph import CT_P
        from docx.oxml.table import CT_Tbl
        from docx.table import _Cell, Table
        from docx.text.paragraph import Paragraph
        
        def process_paragraph(para):
            """Helper to process a paragraph - FIXED to include all text"""
            para_text = para.text.strip()
            if not para_text:
                return
            
            # Detect Heading
            is_heading_para = para.style.name.startswith('Heading')
            
            # Get ALL text including special characters like < and >
            full_para_text = para.text
            
            # Split by whitespace - this preserves <xxxx> tags
            tokens = full_para_text.split()
            
            for token in tokens:
                if token.strip():
                    # Determine formatting from runs
                    is_bold = False
                    is_italic = False
                    
                    for run in para.runs:
                        if token in run.text:
                            if run.bold:
                                is_bold = True
                            if run.italic:
                                is_italic = True
                            break
                    
                    word_objects.append({
                        'text': token,
                        'type': 'word',
                        'is_bold': is_bold,
                        'is_italic': is_italic,
                        'is_heading': is_heading_para,
                        'in_table': False
                    })
                    text_segments.append(token)
            
            # FIXED: Only add ONE newline per paragraph
            word_objects.append({'type': 'newline', 'text': '\n'})
        
        def process_table(table):
            """Helper to process a table - IMPROVED with proper separation"""
            word_objects.append({'type': 'table_start', 'text': ''})
            
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    if cell_text:
                        # Split cell text into tokens - preserves <xxxx>
                        tokens = cell_text.split()
                        for token in tokens:
                            if token.strip():
                                word_objects.append({
                                    'text': token,
                                    'type': 'word',
                                    'is_bold': False,
                                    'is_italic': False,
                                    'is_heading': False,
                                    'in_table': True
                                })
                                text_segments.append(token)
                    
                    # Add cell separator (except for last cell in row)
                    if cell_idx < len(row.cells) - 1:
                        word_objects.append({'type': 'separator', 'text': '|'})
                
                # Add row separator
                word_objects.append({'type': 'row_end', 'text': '\n'})
            
            word_objects.append({'type': 'table_end', 'text': ''})
            # FIXED: Only add ONE newline after table
            word_objects.append({'type': 'newline', 'text': '\n'})
        
        # Iterate through body elements in order
        for element in doc.element.body:
            if isinstance(element, CT_P):
                # It's a paragraph
                para = Paragraph(element, doc)
                process_paragraph(para)
                
            elif isinstance(element, CT_Tbl):
                # It's a table
                table = Table(element, doc)
                process_table(table)

        extracted_text = ' '.join(text_segments)
        return extracted_text, word_objects, doc
        
    except Exception as e:
        st.error(f"Error reading Word document: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None, None, None

def extract_words_from_pdf(pdf_file):
    """
    FROM CODE 1: Original PDF extraction logic
    Returns: (plain_text_string, word_objects, highlight_data, pdf_doc)
    """
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        word_objects = []
        highlight_data = []
        text_segments = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Tables
            tables = page.find_tables()
            table_rects = []
            if tables:
                for table in tables:
                    table_rects.append(table.bbox)
            
            # Blocks for Heading Heuristic
            blocks = page.get_text("blocks")
            block_fonts = {}
            for b_idx, b in enumerate(blocks):
                max_size = 0
                if "lines" in b:
                    for line in b["lines"]:
                        for span in line["spans"]:
                            if span["size"] > max_size:
                                max_size = span["size"]
                block_fonts[b_idx] = max_size

            # Words
            words = page.get_text("words")
            prev_block = -1
            
            for word_info in words:
                x0, y0, x1, y1, word_text = word_info[:5]
                block_no = word_info[5]
                
                # Detect new blocks (paragraphs)
                if block_no != prev_block:
                    if prev_block != -1:
                        word_objects.append({'type': 'newline', 'text': '\n'})
                    prev_block = block_no
                
                if word_text.strip():
                    # Table check
                    in_table = False
                    for table_rect in table_rects:
                        if (x0 >= table_rect[0] and x1 <= table_rect[2] and 
                            y0 >= table_rect[1] and y1 <= table_rect[3]):
                            in_table = True
                            break
                    
                    # Heading check
                    is_heading = False
                    if block_no in block_fonts and block_fonts[block_no] > 14 and len(word_text) < 20:
                        is_heading = True

                    word_objects.append({
                        'text': word_text.strip(),
                        'type': 'word',
                        'is_bold': False,
                        'is_italic': False,
                        'is_heading': is_heading
                    })
                    
                    highlight_data.append({
                        'text': word_text.strip(),
                        'bbox': [x0, y0, x1, y1],
                        'page': page_num,
                        'in_table': in_table
                    })
                    
                    text_segments.append(word_text.strip())
            
            word_objects.append({'type': 'newline', 'text': '\n'}) # Page break
            
        extracted_text = ' '.join(text_segments)
        return extracted_text, word_objects, highlight_data, doc
        
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None, None, None, None

# ============================================================================
# DIFF LOGIC - FROM CODE 1
# ============================================================================

def normalize_word(word):
    import string
    translator = str.maketrans('', '', string.punctuation)
    normalized = word.translate(translator).lower()
    normalized = normalized.replace('"', '').replace('"', '').replace(''', '').replace(''', '')
    return normalized

def run_diff(text1, text2):
    words1 = text1.split()
    words2 = text2.split()
    
    norm_words1 = [normalize_word(w) for w in words1 if normalize_word(w)]
    norm_words2 = [normalize_word(w) for w in words2 if normalize_word(w)]
    
    matcher = difflib.SequenceMatcher(None, norm_words1, norm_words2, autojunk=False)
    opcodes = matcher.get_opcodes()
    
    diff_indices1 = set()
    diff_indices2 = set()
    
    # Map normalized indices back to original
    orig_idx1 = 0
    norm_idx1 = 0
    for w in words1:
        if normalize_word(w):
            for tag, i1, i2, j1, j2 in opcodes:
                if tag == 'replace' and i1 <= norm_idx1 < i2:
                    diff_indices1.add(orig_idx1)
                elif tag == 'delete' and i1 <= norm_idx1 < i2:
                    diff_indices1.add(orig_idx1)
            norm_idx1 += 1
        orig_idx1 += 1
        
    orig_idx2 = 0
    norm_idx2 = 0
    for w in words2:
        if normalize_word(w):
            for tag, i1, i2, j1, j2 in opcodes:
                if tag == 'replace' and j1 <= norm_idx2 < j2:
                    diff_indices2.add(orig_idx2)
                elif tag == 'insert' and j1 <= norm_idx2 < j2:
                    diff_indices2.add(orig_idx2)
            norm_idx2 += 1
        orig_idx2 += 1
        
    total_matching = sum(i2 - i1 for tag, i1, i2, _, _ in opcodes if tag == 'equal')
    
    info = {
        'total_matching': total_matching,
        'total_words1': len(words1),
        'total_words2': len(words2),
        'diff_words1': len(diff_indices1),
        'diff_words2': len(diff_indices2)
    }
    
    return diff_indices1, diff_indices2, info

# ============================================================================
# PREVIEW GENERATION - CODE 2 FOR WORD, CODE 1 FOR PDF
# ============================================================================

def create_html_preview(word_objects, diff_indices):
    """
    FIXED: Simplified HTML generation with correct highlighting
    """
    html_parts = []
    html_parts.append('<div style="padding: 20px; line-height: 1.8;">')
    
    text_idx = 0
    obj_idx = 0
    in_table = False
    
    while obj_idx < len(word_objects):
        obj = word_objects[obj_idx]
        
        if obj['type'] == 'table_start':
            html_parts.append('<div style="margin: 15px 0; overflow-x: auto;">')
            html_parts.append('<table style="width: 100%; border-collapse: collapse; border: 1px solid #ddd;"><tr>')
            in_table = True
            html_parts.append('<td style="padding: 8px; border: 1px solid #ddd; vertical-align: top;">')
            
        elif obj['type'] == 'table_end':
            html_parts.append('</td></tr></table></div>')
            in_table = False
            
        elif obj['type'] == 'word':
            word_html = []
            
            # Apply formatting
            if obj.get('is_heading'):
                word_html.append('<span style="font-size: 1.2em; font-weight: bold; color: #2c3e50;">')
            elif obj.get('is_bold'):
                word_html.append('<strong>')
            
            if obj.get('is_italic'):
                word_html.append('<em>')
            
            # Add the word with or without highlight (escape HTML)
            import html
            escaped_text = html.escape(obj["text"])
            
            # FIXED: Only highlight if this word index is in diff_indices
            if text_idx in diff_indices:
                word_html.append(f'<span class="highlight">{escaped_text}</span>')
            else:
                word_html.append(escaped_text)
            
            # Close tags
            if obj.get('is_italic'):
                word_html.append('</em>')
            
            if obj.get('is_heading'):
                word_html.append('</span>')
            elif obj.get('is_bold'):
                word_html.append('</strong>')
            
            html_parts.append(''.join(word_html))
            html_parts.append(' ')
            
            text_idx += 1
            
        elif obj['type'] == 'row_end':
            if in_table:
                html_parts.append('</td></tr><tr><td style="padding: 8px; border: 1px solid #ddd; vertical-align: top;">')
                
        elif obj['type'] == 'newline':
            if not in_table:
                html_parts.append('<br><br>')
            
        elif obj['type'] == 'separator':
            if in_table:
                html_parts.append('</td><td style="padding: 8px; border: 1px solid #ddd; vertical-align: top;">')
            
        obj_idx += 1
    
    html_parts.append('</div>')
    return "".join(html_parts)

def render_pdf_pages_with_highlights(doc, word_data, diff_indices, max_pages=None):
    """
    FROM CODE 1: Renders PDF pages as PNG images with highlight rectangles
    Returns list of PIL Image objects.
    """
    page_images = []
    num_pages = len(doc) if max_pages is None else min(max_pages, len(doc))
    
    for page_num in range(num_pages):
        page = doc[page_num]
        
        # Render page to pixmap (image) at 2x resolution for better quality
        mat = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=mat)
        
        # Convert to PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Draw highlights on the image
        from PIL import ImageDraw
        draw = ImageDraw.Draw(img, 'RGBA')
        
        for word_idx in diff_indices:
            if word_idx < len(word_data):
                word_info = word_data[word_idx]
                if word_info['page'] == page_num:
                    bbox = word_info['bbox']
                    # Scale coordinates by 2x (same as matrix)
                    x0, y0, x1, y1 = bbox[0]*2, bbox[1]*2, bbox[2]*2, bbox[3]*2
                    
                    # Draw semi-transparent yellow rectangle
                    if word_info.get('in_table', False):
                        fill_color = (255, 237, 224, 100)  # Light orange for tables
                    else:
                        fill_color = (255, 255, 0, 100)  # Yellow
                    
                    draw.rectangle([x0, y0, x1, y1], fill=fill_color, outline=(255, 200, 0, 200), width=2)
        
        page_images.append({
            'page_num': page_num + 1,
            'image': img  # Store PIL Image directly
        })
    
    return page_images

# ============================================================================
# HIGHLIGHTING FOR DOWNLOAD - FROM CODE 1
# ============================================================================

def highlight_pdf_words(doc, word_data, diff_indices):
    """FROM CODE 1: Highlight PDF using annotations"""
    highlighted_doc = fitz.open()
    for page_num in range(len(doc)):
        page = doc[page_num]
        new_page = highlighted_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.show_pdf_page(new_page.rect, doc, page_num)
        
        for word_idx in diff_indices:
            if word_idx < len(word_data):
                word_info = word_data[word_idx]
                if word_info['page'] == page_num:
                    bbox = word_info['bbox']
                    rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                    try:
                        highlight = new_page.add_highlight_annot(rect)
                        if word_info.get('in_table', False):
                            highlight.set_colors(stroke=[1.0, 0.93, 0.88])
                        else:
                            highlight.set_colors(stroke=fitz.utils.getColor("yellow"))
                        highlight.update()
                    except: pass
    return highlighted_doc

def highlight_word_doc(doc, word_objects, diff_indices):
    """FROM CODE 2: Highlight Word doc with proper table handling"""
    from docx.enum.text import WD_COLOR_INDEX
    
    text_idx = 0
    obj_idx = 0
    obj_indices_to_highlight = set()
    
    while obj_idx < len(word_objects):
        obj = word_objects[obj_idx]
        if obj['type'] == 'word':
            if text_idx in diff_indices:
                obj_indices_to_highlight.add(obj_idx)
            text_idx += 1
        obj_idx += 1
    
    current_obj_idx = 0
    
    # Process paragraphs and tables IN ORDER using document body
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    
    for element in doc.element.body:
        if isinstance(element, CT_P):
            # It's a paragraph
            para = Paragraph(element, doc)
            if not para.text.strip():
                continue
            
            for run in para.runs:
                if not run.text:
                    continue
                
                run_words = run.text.split()
                highlight_run = False
                
                for _ in run_words:
                    while current_obj_idx < len(word_objects) and word_objects[current_obj_idx]['type'] != 'word':
                        current_obj_idx += 1
                    
                    if current_obj_idx < len(word_objects) and current_obj_idx in obj_indices_to_highlight:
                        highlight_run = True
                    current_obj_idx += 1
                
                if highlight_run:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    
        elif isinstance(element, CT_Tbl):
            # It's a table
            table = Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if not run.text:
                                continue
                            highlight_run = False
                            run_words = run.text.split()
                            for _ in run_words:
                                 while current_obj_idx < len(word_objects) and word_objects[current_obj_idx]['type'] != 'word':
                                    current_obj_idx += 1
                                 if current_obj_idx < len(word_objects) and current_obj_idx in obj_indices_to_highlight:
                                    highlight_run = True
                                 current_obj_idx += 1
                            if highlight_run:
                                run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ============================================================================
# ORCHESTRATION & MAIN LOOP
# ============================================================================

def run_comparison(d1, d2):
    progress_bar = st.progress(0, text="Initializing...")
    
    is_pdf1 = d1.name.endswith('.pdf')
    is_pdf2 = d2.name.endswith('.pdf')
    
    # Extract original filenames (FROM CODE 2)
    doc1_name = os.path.splitext(d1.name)[0]
    doc2_name = os.path.splitext(d2.name)[0]
    
    # Get extensions
    ext1 = 'pdf' if is_pdf1 else 'docx'
    ext2 = 'pdf' if is_pdf2 else 'docx'
    
    # Create display names (truncated for preview)
    doc1_display = truncate_filename(f"{doc1_name}.{ext1}")
    doc2_display = truncate_filename(f"{doc2_name}.{ext2}")
    
    # 1. Extract
    progress_bar.progress(10, text="Extracting text from documents...")
    
    if is_pdf1:
        text1, w_objs1, high_data1, pdf_doc1 = extract_words_from_pdf(d1)
        docx_doc1 = None
    else:
        text1, w_objs1, docx_doc1 = extract_words_from_word(d1)
        high_data1 = None
        pdf_doc1 = None
        
    if is_pdf2:
        text2, w_objs2, high_data2, pdf_doc2 = extract_words_from_pdf(d2)
        docx_doc2 = None
    else:
        text2, w_objs2, docx_doc2 = extract_words_from_word(d2)
        high_data2 = None
        pdf_doc2 = None
        
    if not text1 or not text2: 
        progress_bar.empty()
        return None
    
    # 2. Diff
    progress_bar.progress(50, text="Analyzing differences...")
    diffs1, diffs2, info = run_diff(text1, text2)
    
    # Debug output
    st.write(f"DEBUG - Doc1: {len(w_objs1)} objects, {info['diff_words1']} differences")
    st.write(f"DEBUG - Doc2: {len(w_objs2)} objects, {info['diff_words2']} differences")
    st.write(f"DEBUG - Sample diff indices Doc1: {sorted(list(diffs1))[:10] if diffs1 else 'None'}")
    st.write(f"DEBUG - Sample diff indices Doc2: {sorted(list(diffs2))[:10] if diffs2 else 'None'}")
    
    # 3. Generate Previews & Downloads
    progress_bar.progress(70, text="Generating highlighted documents...")
    
    # -- Doc 1 --
    if is_pdf1:
        hl_doc1 = highlight_pdf_words(pdf_doc1, high_data1, diffs1)
        pdf1_bytes = BytesIO()
        hl_doc1.save(pdf1_bytes)
        pdf1_bytes.seek(0)
        hl_doc1.close()
        
        # Generate image preview
        progress_bar.progress(75, text="Rendering PDF preview for Doc 1...")
        preview1_images = render_pdf_pages_with_highlights(pdf_doc1, high_data1, diffs1)
        pdf_doc1.close()
        
        preview1_type = 'pdf_images'
        preview1_data = preview1_images
            
    else:
        pdf1_bytes = highlight_word_doc(docx_doc1, w_objs1, diffs1)
        preview1_type = 'html'
        preview1_data = create_html_preview(w_objs1, diffs1)
        
    # -- Doc 2 --
    progress_bar.progress(85, text="Rendering preview for Doc 2...")
    if is_pdf2:
        hl_doc2 = highlight_pdf_words(pdf_doc2, high_data2, diffs2)
        pdf2_bytes = BytesIO()
        hl_doc2.save(pdf2_bytes)
        pdf2_bytes.seek(0)
        hl_doc2.close()
        
        # Generate image preview
        preview2_images = render_pdf_pages_with_highlights(pdf_doc2, high_data2, diffs2)
        pdf_doc2.close()
        
        preview2_type = 'pdf_images'
        preview2_data = preview2_images
            
    else:
        pdf2_bytes = highlight_word_doc(docx_doc2, w_objs2, diffs2)
        preview2_type = 'html'
        preview2_data = create_html_preview(w_objs2, diffs2)
    
    progress_bar.progress(100, text="Complete!")
    # Sleep briefly to show 100%
    import time
    time.sleep(0.5)
    progress_bar.empty()
    
    return {
        'p1_type': preview1_type, 'p1_data': preview1_data, 'd1_bytes': pdf1_bytes, 'ext1': ext1,
        'p2_type': preview2_type, 'p2_data': preview2_data, 'd2_bytes': pdf2_bytes, 'ext2': ext2,
        'info': info,
        'doc1_name': doc1_name,
        'doc2_name': doc2_name,
        'doc1_display': doc1_display,
        'doc2_display': doc2_display
    }

# CSS - FROM CODE 1 (Original highlight styling) + improvements
st.markdown("""
<style>
    :root {
        --brand: #005eac;
        --brand-soft: #e8f2fb;
        --text-strong: #0f172a;
        --text-muted: #475569;
        --panel: #ffffff;
        --line: #dbe7f4;
    }

    html, body, [class*="css"] {
        font-family: "Inter", "Segoe UI", Arial, sans-serif;
    }

    .stApp {
        background: linear-gradient(180deg, #f5f9ff 0%, #f8fbff 50%, #ffffff 100%);
    }

    .main .block-container {
        padding-top: 1.25rem;
        max-width: 100%;
    }

    .brand-header {
        display: flex;
        align-items: center;
        gap: 14px;
        padding: 10px 14px;
        margin-bottom: 12px;
        border: 1px solid var(--line);
        border-radius: 14px;
        background: var(--panel);
        box-shadow: 0 8px 22px rgba(0, 94, 172, 0.08);
    }

    .brand-logo {
        width: 54px;
        height: 54px;
        border-radius: 10px;
        object-fit: contain;
        border: 1px solid var(--line);
        background: #ffffff;
        padding: 6px;
    }

    .brand-text h1,
    .app-title {
        color: var(--brand);
        margin: 0;
        font-size: 1.85rem;
        letter-spacing: 0.02em;
        font-weight: 700;
    }

    .brand-text p,
    .app-subtitle {
        color: var(--text-muted);
        margin: 2px 0 0 0;
        font-size: 0.95rem;
    }

    .stMarkdown h3 {
        color: var(--brand);
        letter-spacing: 0.01em;
    }

    div[data-testid="stFileUploader"] > section {
        border: 1px dashed #91bde0 !important;
        border-radius: 12px !important;
        background: #fbfdff;
    }

    div[data-testid="stFileUploader"] small {
        color: var(--text-muted);
    }

    div[data-testid="stRadio"] label {
        color: var(--text-strong);
    }

    div[data-baseweb="radio"] > div {
        background: #ffffff;
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 6px;
    }

    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 12px;
        box-shadow: 0 6px 16px rgba(15, 23, 42, 0.05);
    }

    div[data-testid="metric-container"] {
        background-color: #ffffff;
        padding: 10px;
        border-radius: 10px;
        border: 1px solid var(--line);
    }

    button[kind="primary"] {
        background: linear-gradient(135deg, #005eac 0%, #1479ce 100%) !important;
        border: none !important;
        color: #ffffff !important;
        border-radius: 10px !important;
        box-shadow: 0 8px 18px rgba(0, 94, 172, 0.25) !important;
    }

    button[kind="secondary"] {
        border-radius: 10px !important;
        border: 1px solid #b8d4ea !important;
        color: #0b3d66 !important;
    }

    div[data-testid="stDownloadButton"] button {
        width: 100%;
        border-radius: 10px !important;
    }
    
    /* Stats metrics */
    div[data-testid="metric-container"] {
        background-color: #f8f9fa;
        padding: 10px;
        border-radius: 5px;
    }
    
    /* Preview containers */
    .preview-wrapper {
        width: 100%;
        height: 85vh;
        display: flex;
        flex-direction: column;
        border: 1px solid var(--line);
        border-radius: 12px;
        background-color: #ffffff;
        overflow: hidden;
        box-shadow: 0 10px 22px rgba(0, 94, 172, 0.08);
    }

    .preview-header {
        background-color: #f2f8ff;
        padding: 10px 15px;
        border-bottom: 1px solid var(--line);
        font-weight: bold;
        font-size: 16px;
        color: var(--brand);
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    .preview-content {
        flex: 1;
        position: relative;
        overflow-y: auto;
        overflow-x: hidden;
        background: #f7fbff;
    }

    .diff-container {
        font-family: "Inter", "Segoe UI", sans-serif;
        font-size: 14px;
        line-height: 1.6;
        padding: 20px;
        overflow-y: auto;
        height: 100%;
        background-color: #ffffff;
        color: var(--text-strong);
        max-width: 100%;
        word-wrap: break-word;
    }

    .diff-container h3 {
        color: var(--brand);
        border-bottom: 2px solid #d7e7f6;
        margin-top: 10px;
        font-size: 18px;
    }

    /* CRITICAL: Highlight style */
    .highlight {
        background-color: #fff3a3 !important;
        padding: 2px 4px !important;
        font-weight: bold !important;
        border-radius: 2px;
        outline: 1px solid #ffd84a;
    }
    
    /* Table styling */
    .diff-container table {
        width: 100%;
        border-collapse: collapse;
        margin: 10px 0;
    }
    
    .diff-container td {
        padding: 8px;
        border: 1px solid #dbe7f4;
        vertical-align: top;
    }
    
    /* Image styling for PDF pages */
    img {
        border-radius: 8px;
        box-shadow: 0 6px 14px rgba(15, 23, 42, 0.1);
        max-width: 100%;
        height: auto;
    }

    [data-testid="stSidebar"] {
        background: #f6faff;
        border-right: 1px solid var(--line);
    }
</style>
""", unsafe_allow_html=True)

# Session State for Caching
if 'results' not in st.session_state: st.session_state.results = None
if 'last_files' not in st.session_state: st.session_state.last_files = None
if 'last_mode' not in st.session_state: st.session_state.last_mode = None
if 'last_logged_key' not in st.session_state: st.session_state.last_logged_key = None

if doc1_file and doc2_file:
    current_files = (doc1_file.name, doc2_file.name)
    
    if st.session_state.last_files != current_files or st.session_state.last_mode != comparison_mode:
        st.session_state.results = None 
        
    if not st.session_state.results:
        res = run_comparison(doc1_file, doc2_file)
        if res:
            st.session_state.results = res
            st.session_state.last_files = current_files
            st.session_state.last_mode = comparison_mode

            log_key = (doc1_file.name, doc2_file.name, comparison_mode)
            if st.session_state.last_logged_key != log_key:
                log_usage_event(
                    doc1_name=res['doc1_name'],
                    doc2_name=res['doc2_name'],
                    ext1=res['ext1'],
                    ext2=res['ext2'],
                    comparison_mode=comparison_mode
                )
                st.session_state.last_logged_key = log_key
        else:
            st.error("Comparison failed.")

    if st.session_state.results:
        r = st.session_state.results
        st.success(":material/task_alt: Comparison complete")
        i = r['info']
        
        # Stats
        col_s1, col_s2, col_s3 = st.columns(3)
        col_s1.metric("Total Words (Doc 1)", i['total_words1'])
        col_s2.metric("Total Words (Doc 2)", i['total_words2'])
        match_pct = (i['total_matching'] / max(i['total_words1'], i['total_words2'])) * 100
        col_s3.metric("Match Rate", f"{match_pct:.1f}%")
        
        st.markdown("---")
        
        # Downloads - FROM CODE 2 (Use original filenames with "highlighted_" prefix)
        st.markdown("### :material/download: Download Highlighted Documents")
        dl_c1, dl_c2 = st.columns(2)
        
        with dl_c1:
            st.download_button(
                ":material/file_download: Download Document 1", 
                r['d1_bytes'].getvalue(), 
                file_name=f"highlighted_{r['doc1_name']}.{r['ext1']}", 
                mime="application/pdf" if r['ext1']=='pdf' else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl1"
            )
        with dl_c2:
            st.download_button(
                ":material/file_download: Download Document 2", 
                r['d2_bytes'].getvalue(), 
                file_name=f"highlighted_{r['doc2_name']}.{r['ext2']}", 
                mime="application/pdf" if r['ext2']=='pdf' else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl2"
            )
            
        st.markdown("### :material/preview: Document Preview")
        
        # Preview Columns
        c1, c2 = st.columns(2)
        
        def show_preview(col, display_name, p_type, p_data):
            with col:
                if p_type == 'pdf_images':
                    # Show PDF as rendered images using Streamlit's native image display
                    st.markdown(f"#### {display_name}")
                    st.caption(f":material/article: {len(p_data)} pages with highlighted differences")
                    
                    # Create a scrollable container
                    with st.container(height=700):
                        for page_info in p_data:
                            st.image(
                                page_info['image'],  # Direct PIL Image object
                                caption=f"Page {page_info['page_num']}",
                                use_container_width=True
                            )
                            
                            # Only add separator if not the last page
                            if page_info['page_num'] < len(p_data):
                                st.divider()
                    
                else:
                    # HTML preview for Word docs - SIMPLIFIED
                    st.markdown(f"#### {display_name}")
                    with st.container(height=700):
                        st.markdown(p_data, unsafe_allow_html=True)

        show_preview(c1, r['doc1_display'], r['p1_type'], r['p1_data'])
        show_preview(c2, r['doc2_display'], r['p2_type'], r['p2_data'])

else:
    st.info(":material/upload_file: Upload both documents to begin comparison")