import streamlit as st
import difflib
from io import BytesIO
import fitz  # PyMuPDF
from docx import Document
import re
import os
import tempfile
import base64

# Sidebar collapsed, layout wide
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
# UNIFIED EXTRACTION (Source of Truth)
# ============================================================================

def extract_words_from_word(docx_file):
    """
    Extracts a unified list of tokens from Word.
    Returns: (plain_text_string, word_objects, doc_obj)
    """
    try:
        docx_file.seek(0)
        doc = Document(docx_file)
        
        word_objects = []
        text_segments = []
        
        for para in doc.paragraphs:
            para_text = para.text.strip()
            if not para_text:
                continue
            
            # Detect Heading
            is_heading_para = "Heading" in para.style.name
            
            # Iterate runs
            for run in para.runs:
                if not run.text:
                    continue
                
                words_in_run = run.text.split()
                for w in words_in_run:
                    word_objects.append({
                        'text': w,
                        'type': 'word',
                        'is_bold': run.bold,
                        'is_italic': run.italic,
                        'is_heading': is_heading_para
                    })
                    text_segments.append(w)
            
            # Add paragraph break marker
            word_objects.append({'type': 'newline', 'text': '\n'})
        
        # Handle Tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        para_text = para.text.strip()
                        if para_text:
                            words = para.text.split()
                            for w in words:
                                word_objects.append({
                                    'text': w,
                                    'type': 'word',
                                    'is_bold': False,
                                    'is_italic': False,
                                    'is_heading': False
                                })
                                text_segments.append(w)
                    word_objects.append({'type': 'separator', 'text': '|'}) # Cell separator
                word_objects.append({'type': 'newline', 'text': '\n'}) # Row separator
            
            word_objects.append({'type': 'newline', 'text': '\n'}) # Table separator

        extracted_text = ' '.join(text_segments)
        return extracted_text, word_objects, doc
        
    except Exception as e:
        st.error(f"Error reading Word document: {str(e)}")
        return None, None, None

def extract_words_from_pdf(pdf_file):
    """
    Extracts unified tokens from PDF.
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
# DIFF LOGIC
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
# PREVIEW GENERATION
# ============================================================================

def create_html_preview(word_objects, diff_indices):
    """
    Generates HTML using word_objects and aligned diff_indices.
    """
    html_parts = []
    
    text_idx = 0
    obj_idx = 0
    
    while obj_idx < len(word_objects):
        obj = word_objects[obj_idx]
        
        if obj['type'] == 'word':
            if text_idx in diff_indices:
                if obj.get('is_heading'):
                    html_parts.append('<h3>')
                if obj.get('is_bold'):
                    html_parts.append('<b>')
                
                html_parts.append(f'<span class="highlight">{obj["text"]}</span>')
                
                if obj.get('is_bold'):
                    html_parts.append('</b>')
                if obj.get('is_heading'):
                    html_parts.append('</h3>')
                html_parts.append(' ')
            else:
                if obj.get('is_heading'):
                    html_parts.append('<h3>')
                if obj.get('is_bold'):
                    html_parts.append('<b>')
                
                html_parts.append(obj["text"])
                
                if obj.get('is_bold'):
                    html_parts.append('</b>')
                if obj.get('is_heading'):
                    html_parts.append('</h3>')
                html_parts.append(' ')
                
            text_idx += 1
            
        elif obj['type'] == 'newline':
            html_parts.append('<br>')
        elif obj['type'] == 'separator':
            html_parts.append(' <span style="color:#ccc">|</span> ')
            
        obj_idx += 1
        
    return "".join(html_parts)

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
    
    w_idx = 0
    current_obj_idx = 0
    
    for para in doc.paragraphs:
        if not para.text.strip(): continue
        
        for run in para.runs:
            if not run.text: continue
            
            run_words = run.text.split()
            highlight_run = False
            
            for _ in run_words:
                while current_obj_idx < len(word_objects) and word_objects[current_obj_idx]['type'] != 'word':
                    current_obj_idx += 1
                
                if current_obj_idx in obj_indices_to_highlight:
                    highlight_run = True
                current_obj_idx += 1
            
            if highlight_run:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if not run.text: continue
                        highlight_run = False
                        run_words = run.text.split()
                        for _ in run_words:
                             while current_obj_idx < len(word_objects) and word_objects[current_obj_idx]['type'] != 'word':
                                current_obj_idx += 1
                             if current_obj_idx in obj_indices_to_highlight:
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
    is_pdf1 = d1.name.endswith('.pdf')
    is_pdf2 = d2.name.endswith('.pdf')
    
    # 1. Extract
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
        
    if not text1 or not text2: return None
    
    # 2. Diff
    diffs1, diffs2, info = run_diff(text1, text2)
    
    # 3. Generate Previews & Downloads
    
    # -- Doc 1 --
    if is_pdf1:
        hl_doc1 = highlight_pdf_words(pdf_doc1, high_data1, diffs1)
        pdf1_bytes = BytesIO()
        hl_doc1.save(pdf1_bytes)
        pdf1_bytes.seek(0)
        hl_doc1.close()
        pdf_doc1.close()
        preview1_type = 'pdf'
        preview1_data = pdf1_bytes.getvalue()
    else:
        pdf1_bytes = highlight_word_doc(docx_doc1, w_objs1, diffs1)
        preview1_type = 'html'
        preview1_data = create_html_preview(w_objs1, diffs1)
        
    # -- Doc 2 --
    if is_pdf2:
        hl_doc2 = highlight_pdf_words(pdf_doc2, high_data2, diffs2)
        pdf2_bytes = BytesIO()
        hl_doc2.save(pdf2_bytes)
        pdf2_bytes.seek(0)
        hl_doc2.close()
        pdf_doc2.close()
        preview2_type = 'pdf'
        preview2_data = pdf2_bytes.getvalue()
    else:
        pdf2_bytes = highlight_word_doc(docx_doc2, w_objs2, diffs2)
        preview2_type = 'html'
        preview2_data = create_html_preview(w_objs2, diffs2)
    
    return {
        'p1_type': preview1_type, 'p1_data': preview1_data, 'd1_bytes': pdf1_bytes, 'ext1': 'pdf' if is_pdf1 else 'docx',
        'p2_type': preview2_type, 'p2_data': preview2_data, 'd2_bytes': pdf2_bytes, 'ext2': 'pdf' if is_pdf2 else 'docx',
        'info': info
    }

# CSS
st.markdown("""
<style>
    /* Container for the entire preview pane */
    .preview-wrapper {
        width: 100%;
        height: 85vh; /* Fixed height for the pane */
        display: flex;
        flex-direction: column;
        border: 1px solid #ddd;
        border-radius: 5px;
        background-color: #ffffff;
        overflow: hidden;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }

    /* Header inside the pane */
    .preview-header {
        background-color: #f8f9fa;
        padding: 10px 15px;
        border-bottom: 1px solid #ddd;
        font-weight: bold;
        font-size: 16px;
        color: #333;
    }

    /* Content area */
    .preview-content {
        flex: 1;
        position: relative;
        overflow: hidden;
    }

    /* PDF Iframe styling */
    .pdf-frame {
        width: 100%;
        height: 100%;
        border: none;
        display: block;
    }

    /* HTML Preview styling */
    .diff-container {
        font-family: 'Segoe UI', sans-serif;
        font-size: 14px;
        line-height: 1.6;
        padding: 20px;
        overflow-y: auto;
        height: 100%;
        background-color: #ffffff;
    }

    .diff-container h3 {
        color: #2c3e50;
        border-bottom: 2px solid #eee;
        margin-top: 10px;
        font-size: 18px;
    }

    .highlight {
        background-color: #ffff00;
        padding: 2px 0;
        font-weight: bold;
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
        
        def show_preview(col, title, p_type, p_data):
            with col:
                # Construct the entire HTML block as a single string to prevent layout breaking
                if p_type == 'pdf':
                    b64 = base64.b64encode(p_data).decode('utf-8')
                    # Note: Zoom is controlled by browser settings. 
                    html_block = f"""
                    <div class="preview-wrapper">
                        <div class="preview-header">{title}</div>
                        <div class="preview-content">
                            <iframe src="data:application/pdf;base64,{b64}" class="pdf-frame"></iframe>
                        </div>
                    </div>
                    """
                    st.markdown(html_block, unsafe_allow_html=True)
                else:
                    html_block = f"""
                    <div class="preview-wrapper">
                        <div class="preview-header">{title}</div>
                        <div class="preview-content">
                            <div class="diff-container">{p_data}</div>
                        </div>
                    </div>
                    """
                    st.markdown(html_block, unsafe_allow_html=True)

        show_preview(c1, "Document 1", r['p1_type'], r['p1_data'])
        show_preview(c2, "Document 2", r['p2_type'], r['p2_data'])

else:
    st.info("👆 Please upload both documents to begin comparison.")