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

# UPDATED: Sidebar collapsed, layout wide
st.set_page_config(page_title="Document Diff Checker", layout="wide", initial_sidebar_state="collapsed")

st.title("📄 Document Diff Checker")
st.markdown("Upload two documents (PDF or Word) to compare and highlight their differences")

# Radio button for comparison mode
comparison_mode = st.radio(
    "Select comparison mode:",
    ["Same Format (PDF vs PDF or Word vs Word)", "Mixed Format (PDF vs Word)"],
    horizontal=True
)

# Create two columns for file uploads
col1, col2 = st.columns(2)

with col1:
    st.subheader("Document 1")
    if comparison_mode == "Same Format (PDF vs PDF or Word vs Word)":
        doc1_file = st.file_uploader("Upload first document", type=['pdf', 'docx'], key="doc1")
    else:
        doc1_file = st.file_uploader("Upload first document (PDF or Word)", type=['pdf', 'docx'], key="doc1_mixed")
    
with col2:
    st.subheader("Document 2")
    if comparison_mode == "Same Format (PDF vs PDF or Word vs Word)":
        doc2_file = st.file_uploader("Upload second document (same format as Doc 1)", type=['pdf', 'docx'], key="doc2")
    else:
        doc2_file = st.file_uploader("Upload second document (PDF or Word)", type=['pdf', 'docx'], key="doc2_mixed")

# ============================================================================
# UNIFIED EXTRACTION - FIXED TO PRESERVE TABLE ORDER
# ============================================================================

def extract_words_from_word(docx_file):
    """
    FIXED: Extracts text with tables in their correct position in document flow
    Returns: (plain_text_string, word_objects, doc_obj)
    """
    try:
        docx_file.seek(0)
        doc = Document(docx_file)
        
        word_objects = []
        text_segments = []
        
        # CRITICAL FIX: Process document elements in order (paragraphs AND tables)
        # We need to iterate through the document body to get elements in sequence
        
        from docx.oxml.text.paragraph import CT_P
        from docx.oxml.table import CT_Tbl
        from docx.table import _Cell, Table
        from docx.text.paragraph import Paragraph
        
        def process_paragraph(para):
            """Helper to process a paragraph"""
            para_text = para.text.strip()
            if not para_text:
                return
            
            # Detect Heading
            is_heading_para = para.style.name.startswith('Heading')
            
            # Extract text preserving exact character-level formatting
            full_para_text = para.text
            
            # Split by actual whitespace to get tokens
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
            
            # Add paragraph break
            word_objects.append({'type': 'newline', 'text': '\n'})
        
        def process_table(table):
            """Helper to process a table"""
            word_objects.append({'type': 'table_start', 'text': ''})
            
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    if cell_text:
                        # Split cell text into tokens
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
                word_objects.append({'type': 'newline', 'text': '\n'})
            
            word_objects.append({'type': 'table_end', 'text': ''})
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
    IMPROVED: Extracts text with better spacing preservation
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
            avg_font_size = 0
            font_count = 0
            
            # Calculate average font size
            for b_idx, b in enumerate(blocks):
                max_size = 0
                text_dict = page.get_text("dict")
                for block in text_dict.get("blocks", []):
                    if "lines" in block:
                        for line in block["lines"]:
                            for span in line["spans"]:
                                avg_font_size += span["size"]
                                font_count += 1
                                if span["size"] > max_size:
                                    max_size = span["size"]
                block_fonts[b_idx] = max_size
            
            avg_font_size = avg_font_size / font_count if font_count > 0 else 12

            # Get text with structure preservation
            text_dict = page.get_text("dict")
            
            prev_block = -1
            
            for block in text_dict.get("blocks", []):
                if block.get("type") == 0:  # Text block
                    block_no = block.get("number", -1)
                    
                    # Check if new block
                    if block_no != prev_block and prev_block != -1:
                        word_objects.append({'type': 'newline', 'text': '\n'})
                    prev_block = block_no
                    
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            span_text = span.get("text", "").strip()
                            
                            if span_text:
                                # Split while preserving punctuation context
                                tokens = span_text.split()
                                
                                for token in tokens:
                                    if token.strip():
                                        # Get bounding box for the span
                                        bbox = span.get("bbox", [0, 0, 0, 0])
                                        
                                        # Check if in table
                                        in_table = False
                                        for table_rect in table_rects:
                                            if (bbox[0] >= table_rect[0] and bbox[2] <= table_rect[2] and 
                                                bbox[1] >= table_rect[1] and bbox[3] <= table_rect[3]):
                                                in_table = True
                                                break
                                        
                                        # Heading check
                                        is_heading = False
                                        font_size = span.get("size", 12)
                                        if font_size > (avg_font_size * 1.3) and len(token) < 30:
                                            is_heading = True

                                        word_objects.append({
                                            'text': token,
                                            'type': 'word',
                                            'is_bold': False,
                                            'is_italic': False,
                                            'is_heading': is_heading,
                                            'in_table': in_table
                                        })
                                        
                                        highlight_data.append({
                                            'text': token,
                                            'bbox': bbox,
                                            'page': page_num,
                                            'in_table': in_table
                                        })
                                        
                                        text_segments.append(token)
            
            word_objects.append({'type': 'newline', 'text': '\n'})
            
        extracted_text = ' '.join(text_segments)
        return extracted_text, word_objects, highlight_data, doc
        
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None, None, None, None

# ============================================================================
# DIFF LOGIC - IMPROVED
# ============================================================================

def normalize_word(word):
    """
    IMPROVED: Better normalization that handles spacing and punctuation issues
    """
    import string
    
    # Remove smart quotes and apostrophes
    word = word.replace('"', '').replace('"', '').replace(''', '').replace(''', '')
    word = word.replace('`', '').replace('´', '')
    
    # Remove all whitespace (to handle extraction differences)
    word = ''.join(word.split())
    
    # Strip leading/trailing punctuation but preserve internal ones
    word = word.strip(string.punctuation + string.whitespace)
    
    # Convert to lowercase
    normalized = word.lower()
    
    return normalized

def run_diff(text1, text2):
    """
    IMPROVED: Better diff with custom word comparison
    """
    words1 = text1.split()
    words2 = text2.split()
    
    norm_words1 = [normalize_word(w) for w in words1 if normalize_word(w)]
    norm_words2 = [normalize_word(w) for w in words2 if normalize_word(w)]
    
    matcher = difflib.SequenceMatcher(None, norm_words1, norm_words2, autojunk=False)
    opcodes = matcher.get_opcodes()
    
    diff_indices1 = set()
    diff_indices2 = set()
    
    # Build mapping of normalized to original indices
    norm_to_orig1 = {}
    norm_to_orig2 = {}
    
    norm_idx = 0
    for orig_idx, w in enumerate(words1):
        if normalize_word(w):
            norm_to_orig1[norm_idx] = orig_idx
            norm_idx += 1
    
    norm_idx = 0
    for orig_idx, w in enumerate(words2):
        if normalize_word(w):
            norm_to_orig2[norm_idx] = orig_idx
            norm_idx += 1
    
    # Process opcodes
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == 'replace':
            for i in range(i1, i2):
                if i in norm_to_orig1:
                    diff_indices1.add(norm_to_orig1[i])
            for j in range(j1, j2):
                if j in norm_to_orig2:
                    diff_indices2.add(norm_to_orig2[j])
                    
        elif tag == 'delete':
            for i in range(i1, i2):
                if i in norm_to_orig1:
                    diff_indices1.add(norm_to_orig1[i])
                    
        elif tag == 'insert':
            for j in range(j1, j2):
                if j in norm_to_orig2:
                    diff_indices2.add(norm_to_orig2[j])
    
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
# PREVIEW GENERATION - IMPROVED
# ============================================================================

def create_html_preview(word_objects, diff_indices):
    """
    IMPROVED: Better HTML generation with table support
    """
    html_parts = []
    
    text_idx = 0
    obj_idx = 0
    in_paragraph = False
    in_table = False
    in_table_row = False
    
    while obj_idx < len(word_objects):
        obj = word_objects[obj_idx]
        
        if obj['type'] == 'table_start':
            # Close any open paragraph
            if in_paragraph:
                html_parts.append('</p>')
                in_paragraph = False
            html_parts.append('<div style="margin: 15px 0; padding: 10px; background: #f9f9f9; border-left: 3px solid #ddd;">')
            html_parts.append('<table style="width: 100%; border-collapse: collapse;"><tr>')
            in_table = True
            in_table_row = True
            html_parts.append('<td style="padding: 8px; border: 1px solid #ddd;">')
            
        elif obj['type'] == 'table_end':
            if in_table_row:
                html_parts.append('</td></tr>')
                in_table_row = False
            html_parts.append('</table>')
            html_parts.append('</div>')
            in_table = False
            
        elif obj['type'] == 'word':
            # Start paragraph if not in table and not already in one
            if not in_paragraph and not in_table:
                html_parts.append('<p>')
                in_paragraph = True
            
            word_html = []
            
            # Apply heading style inline
            if obj.get('is_heading'):
                word_html.append('<span style="font-size: 1.2em; font-weight: bold; color: #2c3e50;">')
            
            if obj.get('is_bold') and not obj.get('is_heading'):
                word_html.append('<strong>')
            
            if obj.get('is_italic'):
                word_html.append('<em>')
            
            # Add the word with or without highlight
            if text_idx in diff_indices:
                word_html.append(f'<span class="highlight">{obj["text"]}</span>')
            else:
                word_html.append(obj["text"])
            
            # Close tags in reverse order
            if obj.get('is_italic'):
                word_html.append('</em>')
            
            if obj.get('is_bold') and not obj.get('is_heading'):
                word_html.append('</strong>')
            
            if obj.get('is_heading'):
                word_html.append('</span>')
            
            html_parts.append(''.join(word_html))
            html_parts.append(' ')
            
            text_idx += 1
            
        elif obj['type'] == 'newline':
            if in_table:
                # In table, newline means end of row, start new row
                html_parts.append('</td></tr><tr><td style="padding: 8px; border: 1px solid #ddd;">')
            else:
                # Close paragraph if open
                if in_paragraph:
                    html_parts.append('</p>')
                    in_paragraph = False
                html_parts.append('<br>')
            
        elif obj['type'] == 'separator':
            if in_table:
                # Cell separator in table
                html_parts.append('</td><td style="padding: 8px; border: 1px solid #ddd;">')
            else:
                html_parts.append(' <span style="color:#ccc">|</span> ')
            
        obj_idx += 1
    
    # Close final paragraph if open
    if in_paragraph:
        html_parts.append('</p>')
    
    if in_table:
        if in_table_row:
            html_parts.append('</td></tr>')
        html_parts.append('</table></div>')
        
    return "".join(html_parts)

def render_pdf_pages_with_highlights(doc, word_data, diff_indices, max_pages=None):
    """
    Renders PDF pages as PNG images with highlight rectangles drawn on them.
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
            'image': img
        })
    
    return page_images

# ============================================================================
# HIGHLIGHTING FOR DOWNLOAD
# ============================================================================

def highlight_pdf_words(doc, word_data, diff_indices):
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
        'p1_type': preview1_type, 'p1_data': preview1_data, 'd1_bytes': pdf1_bytes, 'ext1': 'pdf' if is_pdf1 else 'docx',
        'p2_type': preview2_type, 'p2_data': preview2_data, 'd2_bytes': pdf2_bytes, 'ext2': 'pdf' if is_pdf2 else 'docx',
        'info': info
    }

# CSS - IMPROVED with table styling
st.markdown("""
<style>
    /* Global tweaks */
    .main .block-container {
        padding-top: 2rem;
        max-width: 100%;
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
        border: 1px solid #ddd;
        border-radius: 5px;
        background-color: #ffffff;
        overflow: hidden;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }

    .preview-header {
        background-color: #f8f9fa;
        padding: 10px 15px;
        border-bottom: 1px solid #ddd;
        font-weight: bold;
        font-size: 16px;
        color: #333;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    .preview-content {
        flex: 1;
        position: relative;
        overflow-y: auto;
        overflow-x: hidden;
        background: #f5f5f5;
    }

    /* HTML preview styling */
    .diff-container {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-size: 14px;
        line-height: 1.8;
        padding: 20px;
        overflow-y: auto;
        height: 100%;
        background-color: #ffffff;
        color: #333;
    }

    .diff-container p {
        margin: 0 0 1em 0;
        text-align: justify;
    }

    .diff-container br {
        display: block;
        content: "";
        margin: 0.5em 0;
    }
    
    /* Table styling */
    .diff-container table {
        width: 100%;
        border-collapse: collapse;
        margin: 15px 0;
    }
    
    .diff-container td {
        padding: 8px;
        border: 1px solid #ddd;
        vertical-align: top;
    }
    
    .diff-container tr:nth-child(even) {
        background-color: #f9f9f9;
    }

    .highlight {
        background-color: #ffff00;
        padding: 2px 0;
        font-weight: 500;
    }
    
    /* Image styling for PDF pages */
    img {
        border-radius: 4px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        max-width: 100%;
        height: auto;
    }
</style>
""", unsafe_allow_html=True)

# Session State for Caching
if 'results' not in st.session_state: st.session_state.results = None
if 'last_files' not in st.session_state: st.session_state.last_files = None
if 'last_mode' not in st.session_state: st.session_state.last_mode = None

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
        else:
            st.error("Comparison failed.")

    if st.session_state.results:
        r = st.session_state.results
        st.success("✅ Comparison Complete!")
        i = r['info']
        
        # Stats
        col_s1, col_s2, col_s3 = st.columns(3)
        col_s1.metric("Total Words (Doc 1)", i['total_words1'])
        col_s2.metric("Total Words (Doc 2)", i['total_words2'])
        match_pct = (i['total_matching'] / max(i['total_words1'], i['total_words2'])) * 100
        col_s3.metric("Match Rate", f"{match_pct:.1f}%")
        
        st.markdown("---")
        
        # Downloads
        st.markdown("### Download Highlighted Documents")
        dl_c1, dl_c2 = st.columns(2)
        
        with dl_c1:
            st.download_button(
                "⬇️ Download Doc 1", 
                r['d1_bytes'].getvalue(), 
                file_name=f"doc1_highlighted.{r['ext1']}", 
                mime="application/pdf" if r['ext1']=='pdf' else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl1"
            )
        with dl_c2:
            st.download_button(
                "⬇️ Download Doc 2", 
                r['d2_bytes'].getvalue(), 
                file_name=f"doc2_highlighted.{r['ext2']}", 
                mime="application/pdf" if r['ext2']=='pdf' else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl2"
            )
            
        st.markdown("### Document Preview (Highlighted)")
        
        # Preview Columns
        c1, c2 = st.columns(2)
        
        def show_preview(col, title, p_type, p_data, ext):
            with col:
                if p_type == 'pdf_images':
                    # Show PDF as rendered images
                    st.markdown(f"#### {title}")
                    
                    with st.container(height=700):
                        for page_info in p_data:
                            st.image(
                                page_info['image'],
                                caption=f"Page {page_info['page_num']}",
                                use_container_width=True
                            )
                            
                            if page_info['page_num'] < len(p_data):
                                st.divider()
                    
                else:
                    # HTML preview for Word docs
                    st.markdown(f"#### {title}")
                    with st.container(height=700):
                        st.markdown(f'<div class="diff-container">{p_data}</div>', unsafe_allow_html=True)

        show_preview(c1, "Document 1", r['p1_type'], r['p1_data'], r['ext1'])
        show_preview(c2, "Document 2", r['p2_type'], r['p2_data'], r['ext2'])

else:
    st.info("👆 Please upload both documents to begin comparison.")