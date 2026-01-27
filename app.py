import streamlit as st
import difflib
from io import BytesIO
import fitz  # PyMuPDF
from docx import Document
import re
import os
import tempfile

st.set_page_config(page_title="Document Diff Checker", layout="wide")

st.title("📄 Document Diff Checker")
st.markdown("Upload two documents (PDF or Word) to compare and highlight their differences")

# Create two columns for file uploads
col1, col2 = st.columns(2)

with col1:
    st.subheader("Document 1")
    doc1_file = st.file_uploader("Upload first document", type=['pdf', 'docx'], key="doc1")
    
with col2:
    st.subheader("Document 2")
    doc2_file = st.file_uploader("Upload second document", type=['pdf', 'docx'], key="doc2")

def extract_text_from_word(docx_file):
    """Extract text from Word document maintaining structure"""
    try:
        docx_file.seek(0)
        doc = Document(docx_file)
        
        text_segments = []
        
        # Extract from paragraphs
        for para in doc.paragraphs:
            para_text = para.text.strip()
            if para_text:
                text_segments.append(para_text)
        
        # Extract from tables
        for table in doc.tables:
            for row in table.rows:
                row_texts = []
                for cell in row.cells:
                    cell_text = ' '.join(p.text.strip() for p in cell.paragraphs if p.text.strip())
                    if cell_text:
                        row_texts.append(cell_text)
                if row_texts:
                    text_segments.append(' | '.join(row_texts))
        
        # Join segments with double newline to preserve structure
        extracted_text = '\n\n'.join(text_segments)
        
        return extracted_text
    except Exception as e:
        st.error(f"Error reading Word document: {str(e)}")
        return None

def extract_text_from_pdf(pdf_file):
    """Extract text from PDF with enhanced table detection"""
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        all_words = []
        full_text_parts = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Try to detect tables
            tables = page.find_tables()
            table_rects = []
            
            if tables:
                for table in tables:
                    table_rects.append(table.bbox)
            
            # Get word-level data with coordinates
            words_data = page.get_text("words")
            
            for word_info in words_data:
                x0, y0, x1, y1, word_text = word_info[:5]
                
                if word_text.strip():
                    # Check if word is in a table
                    in_table = False
                    for table_rect in table_rects:
                        if (x0 >= table_rect[0] and x1 <= table_rect[2] and 
                            y0 >= table_rect[1] and y1 <= table_rect[3]):
                            in_table = True
                            break
                    
                    all_words.append({
                        'text': word_text.strip(),
                        'bbox': [x0, y0, x1, y1],
                        'page': page_num,
                        'in_table': in_table
                    })
                    full_text_parts.append(word_text.strip())
        
        # Join with spaces
        full_text = ' '.join(full_text_parts)
        
        return full_text, all_words, doc
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None, None, None

def clean_text_for_comparison(text):
    """
    Clean text to remove formatting differences that shouldn't matter:
    - Remove bullet points (a., b., 1., 2., etc.)
    - Normalize whitespace (including line breaks within words)
    - Remove all punctuation
    - Convert to lowercase
    """
    import string
    
    # CRITICAL FIX: Replace slashes and dashes with spaces BEFORE processing
    # This prevents "whom/institution" from becoming "whominstitution"
    text = text.replace('/', ' ').replace('-', ' ').replace('–', ' ').replace('—', ' ')
    
    # Split into words
    words = text.split()
    
    cleaned_words = []
    for word in words:
        # Remove common bullet point patterns
        # Match patterns like: a., b., 1., 2., i., ii., (a), (1), etc.
        if re.match(r'^[a-z]\.$|^[0-9]+\.$|^[ivxlcdm]+\.$|^\([a-z0-9]+\)$', word.lower()):
            continue  # Skip bullet points
        
        # Remove ALL punctuation including periods at the end
        translator = str.maketrans('', '', string.punctuation)
        cleaned = word.translate(translator).lower()
        
        # Handle unicode characters
        cleaned = cleaned.replace('"', '').replace('"', '').replace(''', '').replace(''', '')
        
        # Remove any remaining whitespace
        cleaned = cleaned.strip()
        
        if cleaned:  # Only keep non-empty words
            cleaned_words.append(cleaned)
    
    return ' '.join(cleaned_words)

def find_word_differences_robust(text1, text2):
    """
    Robust comparison that handles:
    - Line breaks within words
    - Bullet points
    - Punctuation differences
    - Whitespace variations
    """
    
    # Split into words BEFORE cleaning (to maintain original structure)
    words1_original = text1.split()
    words2_original = text2.split()
    
    # Clean the full texts for comparison
    cleaned_text1 = clean_text_for_comparison(text1)
    cleaned_text2 = clean_text_for_comparison(text2)
    
    # If cleaned texts are identical, no differences
    if cleaned_text1 == cleaned_text2:
        return set(), set(), words1_original, words2_original, {
            'total_matching': len(cleaned_text1.split()),
            'total_words1': len(words1_original),
            'total_words2': len(words2_original),
            'equal_blocks': 1,
            'diff_words1': 0,
            'diff_words2': 0
        }
    
    # Split cleaned texts into words
    cleaned_words1 = cleaned_text1.split()
    cleaned_words2 = cleaned_text2.split()
    
    # Use SequenceMatcher on cleaned words
    matcher = difflib.SequenceMatcher(None, cleaned_words1, cleaned_words2, autojunk=False)
    opcodes = matcher.get_opcodes()
    
    # Build mapping from cleaned index to original index
    # This is tricky because bullet points are removed from cleaned but present in original
    cleaned_to_orig1 = []
    orig_idx = 0
    for word in words1_original:
        # Check if this word is a bullet point
        if re.match(r'^[a-z]\.$|^[0-9]+\.$|^[ivxlcdm]+\.$|^\([a-z0-9]+\)$', word.lower()):
            orig_idx += 1
            continue  # Don't map bullet points
        
        # Check if this word becomes empty after cleaning
        import string
        # Apply same cleaning as clean_text_for_comparison
        temp_word = word.replace('/', ' ').replace('-', ' ').replace('–', ' ').replace('—', ' ')
        translator = str.maketrans('', '', string.punctuation)
        cleaned = temp_word.translate(translator).lower()
        cleaned = cleaned.replace('"', '').replace('"', '').replace(''', '').replace(''', '')
        cleaned = cleaned.strip()
        
        if cleaned:
            cleaned_to_orig1.append(orig_idx)
        orig_idx += 1
    
    cleaned_to_orig2 = []
    orig_idx = 0
    for word in words2_original:
        if re.match(r'^[a-z]\.$|^[0-9]+\.$|^[ivxlcdm]+\.$|^\([a-z0-9]+\)$', word.lower()):
            orig_idx += 1
            continue
        
        import string
        # Apply same cleaning as clean_text_for_comparison
        temp_word = word.replace('/', ' ').replace('-', ' ').replace('–', ' ').replace('—', ' ')
        translator = str.maketrans('', '', string.punctuation)
        cleaned = temp_word.translate(translator).lower()
        cleaned = cleaned.replace('"', '').replace('"', '').replace(''', '').replace(''', '')
        cleaned = cleaned.strip()
        
        if cleaned:
            cleaned_to_orig2.append(orig_idx)
        orig_idx += 1
    
    # Identify differing regions in cleaned space
    diff_cleaned_indices1 = set()
    diff_cleaned_indices2 = set()
    
    for tag, i1, i2, j1, j2 in opcodes:
        if tag != 'equal':
            diff_cleaned_indices1.update(range(i1, i2))
            diff_cleaned_indices2.update(range(j1, j2))
    
    # Map back to original indices
    diff_indices1 = set()
    for cleaned_idx in diff_cleaned_indices1:
        if cleaned_idx < len(cleaned_to_orig1):
            diff_indices1.add(cleaned_to_orig1[cleaned_idx])
    
    diff_indices2 = set()
    for cleaned_idx in diff_cleaned_indices2:
        if cleaned_idx < len(cleaned_to_orig2):
            diff_indices2.add(cleaned_to_orig2[cleaned_idx])
    
    # Calculate statistics
    total_matching = sum(i2 - i1 for tag, i1, i2, _, _ in opcodes if tag == 'equal')
    total_equal_blocks = sum(1 for tag, _, _, _, _ in opcodes if tag == 'equal')
    
    sync_info = {
        'total_matching': total_matching,
        'total_words1': len(words1_original),
        'total_words2': len(words2_original),
        'equal_blocks': total_equal_blocks,
        'diff_words1': len(diff_indices1),
        'diff_words2': len(diff_indices2)
    }
    
    return diff_indices1, diff_indices2, words1_original, words2_original, sync_info

def create_html_diff(text1, text2, diff_indices1, diff_indices2):
    """Create HTML with highlighted differences"""
    def highlight_text(text, diff_indices):
        words = text.split()
        html_parts = []
        
        for i, word in enumerate(words):
            if i in diff_indices:
                html_parts.append(f'<span class="highlight">{word}</span>')
            else:
                html_parts.append(word)
        
        return ' '.join(html_parts)
    
    html1 = highlight_text(text1, diff_indices1)
    html2 = highlight_text(text2, diff_indices2)
    
    return html1, html2

def highlight_pdf_words(doc, word_data, diff_indices):
    """Highlight specific words in PDF based on indices with enhanced table support"""
    highlighted_doc = fitz.open()
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        new_page = highlighted_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.show_pdf_page(new_page.rect, doc, page_num)
        
        # Highlight words on this page
        for word_idx in diff_indices:
            if word_idx < len(word_data):
                word_info = word_data[word_idx]
                if word_info['page'] == page_num:
                    bbox = word_info['bbox']
                    rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                    
                    try:
                        # Use different color for table content
                        if word_info.get('in_table', False):
                            highlight = new_page.add_highlight_annot(rect)
                            # Light orange/peach color
                            highlight.set_colors(stroke=[1.0, 0.93, 0.88])
                        else:
                            highlight = new_page.add_highlight_annot(rect)
                            highlight.set_colors(stroke=fitz.utils.getColor("yellow"))
                        highlight.update()
                    except:
                        pass
    
    return highlighted_doc

def highlight_word_doc(docx_file, extracted_text, diff_indices):
    """Highlight words in Word document with precise run-level matching."""
    from docx.enum.text import WD_COLOR_INDEX
    
    docx_file.seek(0)
    doc = Document(docx_file)
    
    extracted_words = extracted_text.replace('\n\n', ' ').split()
    
    word_idx = 0
    
    # Process regular paragraphs
    for para in doc.paragraphs:
        para_text = para.text.strip()
        if not para_text:
            continue
        
        para_words = para_text.split()
        para_start_idx = word_idx
        
        # Try to align paragraph with extracted text
        matches = True
        for i, pword in enumerate(para_words):
            if word_idx + i >= len(extracted_words):
                matches = False
                break
            # Use cleaned comparison
            if clean_text_for_comparison(pword) != clean_text_for_comparison(extracted_words[word_idx + i]):
                matches = False
                break
        
        if not matches:
            # Search for alignment
            for search_offset in range(max(0, word_idx - 10), min(len(extracted_words), word_idx + 50)):
                temp_matches = True
                for i, pword in enumerate(para_words):
                    if search_offset + i >= len(extracted_words):
                        temp_matches = False
                        break
                    if clean_text_for_comparison(pword) != clean_text_for_comparison(extracted_words[search_offset + i]):
                        temp_matches = False
                        break
                if temp_matches:
                    word_idx = search_offset
                    para_start_idx = search_offset
                    break
        
        # Highlight runs
        run_word_position = para_start_idx
        
        for run in para.runs:
            run_text = run.text
            if not run_text:
                continue
            
            run_words = run_text.split()
            if not run_words:
                continue
            
            should_highlight = False
            for i in range(len(run_words)):
                if (run_word_position + i) in diff_indices:
                    should_highlight = True
                    break
            
            if should_highlight:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            
            run_word_position += len(run_words)
        
        word_idx += len(para_words)
    
    # Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para_text = para.text.strip()
                    if not para_text:
                        continue
                    
                    para_words = para_text.split()
                    cell_start_idx = word_idx
                    
                    matches = True
                    for i, cword in enumerate(para_words):
                        if word_idx + i >= len(extracted_words):
                            matches = False
                            break
                        if clean_text_for_comparison(cword) != clean_text_for_comparison(extracted_words[word_idx + i]):
                            matches = False
                            break
                    
                    run_word_position = cell_start_idx if matches else word_idx
                    
                    for run in para.runs:
                        run_text = run.text
                        if not run_text:
                            continue
                        
                        run_words = run_text.split()
                        if not run_words:
                            continue
                        
                        should_highlight = False
                        for i in range(len(run_words)):
                            if (run_word_position + i) in diff_indices:
                                should_highlight = True
                                break
                        
                        if should_highlight:
                            run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
                        
                        run_word_position += len(run_words)
                    
                    if matches:
                        word_idx += len(para_words)
                
                if word_idx < len(extracted_words) and cell != row.cells[-1]:
                    if word_idx + 1 < len(extracted_words) and extracted_words[word_idx] == '|':
                        word_idx += 1
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

def compare_pdf_vs_word(pdf_file, word_file, pdf_is_doc1=True):
    """Compare PDF vs Word document"""
    with st.spinner("Extracting text from documents..."):
        text_pdf, word_data_pdf, pdf_doc = extract_text_from_pdf(pdf_file)
        text_word = extract_text_from_word(word_file)
    
    if not text_pdf or not text_word:
        return None
    
    if pdf_is_doc1:
        text1, text2 = text_pdf, text_word
        word_data1 = word_data_pdf
        word_data2 = None
    else:
        text1, text2 = text_word, text_pdf
        word_data1 = None
        word_data2 = word_data_pdf
    
    with st.spinner("Finding differences (excluding punctuation & bullet points)..."):
        diff_indices1, diff_indices2, words1, words2, sync_info = find_word_differences_robust(text1, text2)
        html1, html2 = create_html_diff(text1, text2, diff_indices1, diff_indices2)
    
    with st.spinner("Generating highlighted documents..."):
        if pdf_is_doc1:
            highlighted_pdf = highlight_pdf_words(pdf_doc, word_data_pdf, diff_indices1)
            pdf_bytes = BytesIO()
            highlighted_pdf.save(pdf_bytes)
            pdf_bytes.seek(0)
            highlighted_pdf.close()
            
            word_bytes = highlight_word_doc(word_file, text_word, diff_indices2)
            
            doc1_bytes = pdf_bytes
            doc2_bytes = word_bytes
        else:
            word_bytes = highlight_word_doc(word_file, text_word, diff_indices1)
            
            highlighted_pdf = highlight_pdf_words(pdf_doc, word_data_pdf, diff_indices2)
            pdf_bytes = BytesIO()
            highlighted_pdf.save(pdf_bytes)
            pdf_bytes.seek(0)
            highlighted_pdf.close()
            
            doc1_bytes = word_bytes
            doc2_bytes = pdf_bytes
        
        pdf_doc.close()
    
    return {
        'text1': text1,
        'text2': text2,
        'diff_indices1': diff_indices1,
        'diff_indices2': diff_indices2,
        'words1': words1,
        'words2': words2,
        'html1': html1,
        'html2': html2,
        'pdf1_bytes': doc1_bytes,
        'pdf2_bytes': doc2_bytes,
        'is_pdf1': pdf_is_doc1,
        'is_pdf2': not pdf_is_doc1,
        'sync_info': sync_info,
        'is_mixed_format': True
    }

# CSS for styling
st.markdown("""
<style>
    .diff-container {
        font-family: 'Courier New', monospace;
        font-size: 13px;
        line-height: 1.8;
        padding: 20px;
        background-color: #ffffff;
        border-radius: 5px;
        max-height: 700px;
        overflow-y: auto;
        border: 1px solid #ddd;
        white-space: pre-wrap;
        word-wrap: break-word;
    }
    .highlight {
        background-color: #ffff00;
        padding: 1px 2px;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'comparison_done' not in st.session_state:
    st.session_state.comparison_done = False
if 'results' not in st.session_state:
    st.session_state.results = None

# Process and display differences
if doc1_file and doc2_file:
    current_files = (doc1_file.name, doc2_file.name)
    if 'last_files' not in st.session_state or st.session_state.last_files != current_files:
        st.session_state.comparison_done = False
        st.session_state.last_files = current_files
    
    if not st.session_state.comparison_done:
        is_pdf1 = doc1_file.name.endswith('.pdf')
        is_pdf2 = doc2_file.name.endswith('.pdf')
        
        is_mixed_format = (is_pdf1 and not is_pdf2) or (not is_pdf1 and is_pdf2)
        
        if is_mixed_format:
            st.info("🔄 Mixed format detected (PDF vs Word). Comparing with robust normalization...")
            
            if is_pdf1:
                results = compare_pdf_vs_word(doc1_file, doc2_file, pdf_is_doc1=True)
            else:
                results = compare_pdf_vs_word(doc2_file, doc1_file, pdf_is_doc1=False)
            
            if results:
                st.session_state.results = results
                st.session_state.comparison_done = True
        
        else:
            # Same-format comparison
            with st.spinner("Extracting text from documents..."):
                if is_pdf1:
                    text1, word_data1, pdf_doc1 = extract_text_from_pdf(doc1_file)
                else:
                    text1 = extract_text_from_word(doc1_file)
                    word_data1 = None
                    pdf_doc1 = None
                
                if is_pdf2:
                    text2, word_data2, pdf_doc2 = extract_text_from_pdf(doc2_file)
                else:
                    text2 = extract_text_from_word(doc2_file)
                    word_data2 = None
                    pdf_doc2 = None
            
            if text1 and text2:
                with st.spinner("Finding differences (excluding punctuation & bullet points)..."):
                    diff_indices1, diff_indices2, words1, words2, sync_info = find_word_differences_robust(text1, text2)
                    html1, html2 = create_html_diff(text1, text2, diff_indices1, diff_indices2)
                
                with st.spinner("Generating highlighted documents..."):
                    if is_pdf1:
                        highlighted_doc1 = highlight_pdf_words(pdf_doc1, word_data1, diff_indices1)
                        pdf1_bytes = BytesIO()
                        highlighted_doc1.save(pdf1_bytes)
                        pdf1_bytes.seek(0)
                        highlighted_doc1.close()
                        pdf_doc1.close()
                    else:
                        pdf1_bytes = highlight_word_doc(doc1_file, text1, diff_indices1)
                    
                    if is_pdf2:
                        highlighted_doc2 = highlight_pdf_words(pdf_doc2, word_data2, diff_indices2)
                        pdf2_bytes = BytesIO()
                        highlighted_doc2.save(pdf2_bytes)
                        pdf2_bytes.seek(0)
                        highlighted_doc2.close()
                        pdf_doc2.close()
                    else:
                        pdf2_bytes = highlight_word_doc(doc2_file, text2, diff_indices2)
                
                st.session_state.results = {
                    'text1': text1,
                    'text2': text2,
                    'diff_indices1': diff_indices1,
                    'diff_indices2': diff_indices2,
                    'words1': words1,
                    'words2': words2,
                    'html1': html1,
                    'html2': html2,
                    'pdf1_bytes': pdf1_bytes,
                    'pdf2_bytes': pdf2_bytes,
                    'is_pdf1': is_pdf1,
                    'is_pdf2': is_pdf2,
                    'sync_info': sync_info,
                    'is_mixed_format': False
                }
                st.session_state.comparison_done = True
    
    # Display results
    if st.session_state.results:
        results = st.session_state.results
        
        st.success("✅ Comparison complete!")
        
        sync_info = results['sync_info']
        match_percentage = (sync_info['total_matching'] / max(sync_info['total_words1'], sync_info['total_words2'])) * 100
        
        st.info(f"📊 **Match**: {sync_info['total_matching']} matching words out of {max(sync_info['total_words1'], sync_info['total_words2'])} ({match_percentage:.1f}%) • Punctuation & bullet points excluded")
        
        col_stat1, col_stat2, col_stat3 = st.columns(3)
        with col_stat1:
            st.metric("Words in Doc 1", sync_info['total_words1'])
            st.metric("Different in Doc 1", sync_info['diff_words1'])
        with col_stat2:
            st.metric("Words in Doc 2", sync_info['total_words2'])
            st.metric("Different in Doc 2", sync_info['diff_words2'])
        with col_stat3:
            similarity = (sync_info['total_matching'] / max(sync_info['total_words1'], sync_info['total_words2'])) * 100
            st.metric("Match Rate", f"{similarity:.1f}%")
            st.metric("Sync Blocks", sync_info['equal_blocks'])
        
        st.markdown("### Download Highlighted Documents")
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            file_ext1 = 'pdf' if results.get('is_pdf1', False) else 'docx'
            st.download_button(
                label=f"⬇️ Download Doc 1 (Highlighted .{file_ext1})",
                data=results['pdf1_bytes'].getvalue(),
                file_name=f"doc1_highlighted.{file_ext1}",
                mime="application/pdf" if results.get('is_pdf1', False) else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
        with col_dl2:
            file_ext2 = 'pdf' if results.get('is_pdf2', False) else 'docx'
            st.download_button(
                label=f"⬇️ Download Doc 2 (Highlighted .{file_ext2})",
                data=results['pdf2_bytes'].getvalue(),
                file_name=f"doc2_highlighted.{file_ext2}",
                mime="application/pdf" if results.get('is_pdf2', False) else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
        st.markdown("### Text Comparison Preview")
        st.markdown("🟡 **Yellow** = Regular text differences | 🟠 **Orange** (PDF) = Table differences | 🟢 **Green** (Word) = Table differences")
        
        col_diff1, col_diff2 = st.columns(2)
        
        with col_diff1:
            st.markdown("**Document 1**")
            st.markdown(f'<div class="diff-container">{results["html1"]}</div>', unsafe_allow_html=True)
        
        with col_diff2:
            st.markdown("**Document 2**")
            st.markdown(f'<div class="diff-container">{results["html2"]}</div>', unsafe_allow_html=True)
        
        with st.expander("📋 View Sample Differences (First 20)"):
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                st.markdown(f"**Different words in Doc 1: {len(results['diff_indices1'])} total**")
                sample_indices1 = sorted(list(results['diff_indices1']))[:20]
                for idx in sample_indices1:
                    if idx < len(results['words1']):
                        word1 = results['words1'][idx]
                        st.text(f"Pos {idx}: '{word1}'")
            with col_s2:
                st.markdown(f"**Different words in Doc 2: {len(results['diff_indices2'])} total**")
                sample_indices2 = sorted(list(results['diff_indices2']))[:20]
                for idx in sample_indices2:
                    if idx < len(results['words2']):
                        word2 = results['words2'][idx]
                        st.text(f"Pos {idx}: '{word2}'")

else:
    st.info("👆 Please upload both documents to begin comparison")

st.markdown("---")
st.markdown("💡 **Robust comparison** - Ignoimport streamlit as st
import difflib
from io import BytesIO
import fitz  # PyMuPDF
from docx import Document
import re
import os
import tempfile

st.set_page_config(page_title="Document Diff Checker", layout="wide")

st.title("📄 Document Diff Checker")
st.markdown("Upload two documents (PDF or Word) to compare and highlight their differences")

# Create two columns for file uploads
col1, col2 = st.columns(2)

with col1:
    st.subheader("Document 1")
    doc1_file = st.file_uploader("Upload first document", type=['pdf', 'docx'], key="doc1")
    
with col2:
    st.subheader("Document 2")
    doc2_file = st.file_uploader("Upload second document", type=['pdf', 'docx'], key="doc2")

def extract_text_from_word(docx_file):
    """Extract text from Word document maintaining structure"""
    try:
        docx_file.seek(0)
        doc = Document(docx_file)
        
        text_segments = []
        
        # Extract from paragraphs
        for para in doc.paragraphs:
            para_text = para.text.strip()
            if para_text:
                text_segments.append(para_text)
        
        # Extract from tables
        for table in doc.tables:
            for row in table.rows:
                row_texts = []
                for cell in row.cells:
                    cell_text = ' '.join(p.text.strip() for p in cell.paragraphs if p.text.strip())
                    if cell_text:
                        row_texts.append(cell_text)
                if row_texts:
                    text_segments.append(' | '.join(row_texts))
        
        # Join segments with double newline to preserve structure
        extracted_text = '\n\n'.join(text_segments)
        
        return extracted_text
    except Exception as e:
        st.error(f"Error reading Word document: {str(e)}")
        return None

def extract_text_from_pdf(pdf_file):
    """Extract text from PDF with enhanced table detection"""
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        all_words = []
        full_text_parts = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Try to detect tables
            tables = page.find_tables()
            table_rects = []
            
            if tables:
                for table in tables:
                    table_rects.append(table.bbox)
            
            # Get word-level data with coordinates
            words_data = page.get_text("words")
            
            for word_info in words_data:
                x0, y0, x1, y1, word_text = word_info[:5]
                
                if word_text.strip():
                    # Check if word is in a table
                    in_table = False
                    for table_rect in table_rects:
                        if (x0 >= table_rect[0] and x1 <= table_rect[2] and 
                            y0 >= table_rect[1] and y1 <= table_rect[3]):
                            in_table = True
                            break
                    
                    all_words.append({
                        'text': word_text.strip(),
                        'bbox': [x0, y0, x1, y1],
                        'page': page_num,
                        'in_table': in_table
                    })
                    full_text_parts.append(word_text.strip())
        
        # Join with spaces
        full_text = ' '.join(full_text_parts)
        
        return full_text, all_words, doc
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None, None, None

def normalize_word(word):
    """Normalize word by removing ALL punctuation and converting to lowercase for comparison"""
    # Remove ALL punctuation including brackets, quotes, dashes, periods, etc.
    # This ensures words are compared purely on alphanumeric content
    import string
    # Create a translation table that removes all punctuation
    translator = str.maketrans('', '', string.punctuation)
    # Remove all punctuation and convert to lowercase
    normalized = word.translate(translator).lower()
    # Also handle special unicode quotes and dashes
    normalized = normalized.replace('"', '').replace('"', '').replace(''', '').replace(''', '')
    normalized = normalized.replace('–', '').replace('—', '')
    return normalized

def normalize_for_comparison(text):
    """Normalize text for comparison - removes ALL punctuation differences"""
    # Split into words
    words = text.split()
    # Normalize each word and filter out empty strings (pure punctuation)
    normalized = []
    for w in words:
        norm_word = normalize_word(w)
        if norm_word:  # Only include if there's actual content left after removing punctuation
            normalized.append(norm_word)
    return normalized

def find_word_differences_optimized(text1, text2):
    """
    Optimized word-level difference detection with better alignment handling.
    Uses content-based matching to avoid drift in long documents.
    Excludes punctuation-only differences.
    """
    # Split into words
    words1 = text1.split()
    words2 = text2.split()
    
    # Create normalized versions for comparison (excluding punctuation)
    normalized1 = normalize_for_comparison(text1)
    normalized2 = normalize_for_comparison(text2)
    
    # Use SequenceMatcher with optimized settings
    matcher = difflib.SequenceMatcher(None, normalized1, normalized2, autojunk=False)
    
    # Get the opcodes
    opcodes = matcher.get_opcodes()
    
    # Sets to store indices of different words in NORMALIZED space
    diff_indices_norm1 = set()
    diff_indices_norm2 = set()
    
    # Create mapping from normalized to original indices
    norm_to_orig1 = []
    norm_to_orig2 = []
    
    for i, word in enumerate(words1):
        if normalize_word(word):
            norm_to_orig1.append(i)
    
    for i, word in enumerate(words2):
        if normalize_word(word):
            norm_to_orig2.append(i)
    
    # Process each operation
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == 'equal':
            # Words match - don't highlight
            continue
        elif tag == 'replace':
            # Check if this is a true content difference
            range1_words = normalized1[i1:i2]
            range2_words = normalized2[j1:j2]
            
            # Only mark as different if normalized content actually differs
            if range1_words != range2_words:
                diff_indices_norm1.update(range(i1, i2))
                diff_indices_norm2.update(range(j1, j2))
        elif tag == 'delete':
            # Words only in doc1
            diff_indices_norm1.update(range(i1, i2))
        elif tag == 'insert':
            # Words only in doc2
            diff_indices_norm2.update(range(j1, j2))
    
    # Map back to original word indices
    diff_indices1 = set()
    diff_indices2 = set()
    
    for norm_idx in diff_indices_norm1:
        if norm_idx < len(norm_to_orig1):
            diff_indices1.add(norm_to_orig1[norm_idx])
    
    for norm_idx in diff_indices_norm2:
        if norm_idx < len(norm_to_orig2):
            diff_indices2.add(norm_to_orig2[norm_idx])
    
    # CRITICAL: Remove false positives where both docs have same word at same position
    # This happens when alignment drifts
    validated_diff1 = set()
    validated_diff2 = set()
    
    # Create a reverse mapping to find normalized positions from original positions
    orig_to_norm1 = {}
    for norm_idx, orig_idx in enumerate(norm_to_orig1):
        orig_to_norm1[orig_idx] = norm_idx
    
    orig_to_norm2 = {}
    for norm_idx, orig_idx in enumerate(norm_to_orig2):
        orig_to_norm2[orig_idx] = norm_idx
    
    # Validate diff_indices1
    for orig_idx in diff_indices1:
        if orig_idx < len(words1):
            word1_norm = normalize_word(words1[orig_idx])
            
            # Find corresponding position in doc2
            if orig_idx in orig_to_norm1:
                norm_idx1 = orig_to_norm1[orig_idx]
                
                # Check if this normalized index has a corresponding position in doc2
                # by looking at the opcodes to find the mapping
                is_truly_different = True
                
                for tag, i1, i2, j1, j2 in opcodes:
                    if tag == 'equal' and i1 <= norm_idx1 < i2:
                        # This word is in an equal block, find its pair
                        offset = norm_idx1 - i1
                        norm_idx2 = j1 + offset
                        
                        if norm_idx2 < len(norm_to_orig2):
                            orig_idx2 = norm_to_orig2[norm_idx2]
                            if orig_idx2 < len(words2):
                                word2_norm = normalize_word(words2[orig_idx2])
                                # If words are actually the same, don't highlight
                                if word1_norm == word2_norm:
                                    is_truly_different = False
                                    break
                
                if is_truly_different:
                    validated_diff1.add(orig_idx)
            else:
                validated_diff1.add(orig_idx)
    
    # Validate diff_indices2
    for orig_idx in diff_indices2:
        if orig_idx < len(words2):
            word2_norm = normalize_word(words2[orig_idx])
            
            # Find corresponding position in doc1
            if orig_idx in orig_to_norm2:
                norm_idx2 = orig_to_norm2[orig_idx]
                
                is_truly_different = True
                
                for tag, i1, i2, j1, j2 in opcodes:
                    if tag == 'equal' and j1 <= norm_idx2 < j2:
                        # This word is in an equal block, find its pair
                        offset = norm_idx2 - j1
                        norm_idx1 = i1 + offset
                        
                        if norm_idx1 < len(norm_to_orig1):
                            orig_idx1 = norm_to_orig1[norm_idx1]
                            if orig_idx1 < len(words1):
                                word1_norm = normalize_word(words1[orig_idx1])
                                # If words are actually the same, don't highlight
                                if word1_norm == word2_norm:
                                    is_truly_different = False
                                    break
                
                if is_truly_different:
                    validated_diff2.add(orig_idx)
            else:
                validated_diff2.add(orig_idx)
    
    diff_indices1 = validated_diff1
    diff_indices2 = validated_diff2
    
    # Calculate statistics based on normalized words
    total_matching_words = sum(i2 - i1 for tag, i1, i2, _, _ in opcodes if tag == 'equal')
    total_equal_blocks = sum(1 for tag, _, _, _, _ in opcodes if tag == 'equal')
    
    sync_info = {
        'total_matching': total_matching_words,
        'total_words1': len(words1),
        'total_words2': len(words2),
        'equal_blocks': total_equal_blocks,
        'diff_words1': len(diff_indices1),
        'diff_words2': len(diff_indices2)
    }
    
    return diff_indices1, diff_indices2, words1, words2, sync_info

def find_sentence_differences(text1, text2):
    """
    Find differences at sentence level, excluding ALL punctuation differences.
    Returns sets of sentence indices that differ.
    """
    # Split into sentences - handle multiple sentence terminators
    sentences1 = re.split(r'[.!?]+(?:\s+|$)', text1)
    sentences2 = re.split(r'[.!?]+(?:\s+|$)', text2)
    
    # Clean up empty sentences
    sentences1 = [s.strip() for s in sentences1 if s.strip()]
    sentences2 = [s.strip() for s in sentences2 if s.strip()]
    
    # Normalize sentences for comparison - remove ALL punctuation
    normalized1 = []
    for s in sentences1:
        norm_words = normalize_for_comparison(s)
        normalized1.append(' '.join(norm_words))
    
    normalized2 = []
    for s in sentences2:
        norm_words = normalize_for_comparison(s)
        normalized2.append(' '.join(norm_words))
    
    # Use SequenceMatcher with higher threshold for better matching
    matcher = difflib.SequenceMatcher(None, normalized1, normalized2, autojunk=False)
    opcodes = matcher.get_opcodes()
    
    diff_indices1 = set()
    diff_indices2 = set()
    
    for tag, i1, i2, j1, j2 in opcodes:
        if tag != 'equal':
            # Double-check that sentences are actually different
            # (not just due to residual punctuation issues)
            for i in range(i1, i2):
                if i < len(normalized1):
                    # Check if this sentence has actual content differences
                    has_real_diff = True
                    if tag == 'replace' and (j1 + (i - i1)) < len(normalized2):
                        j = j1 + (i - i1)
                        if normalized1[i] == normalized2[j]:
                            has_real_diff = False
                    if has_real_diff:
                        diff_indices1.add(i)
            
            for j in range(j1, j2):
                if j < len(normalized2):
                    has_real_diff = True
                    if tag == 'replace' and (i1 + (j - j1)) < len(normalized1):
                        i = i1 + (j - j1)
                        if normalized1[i] == normalized2[j]:
                            has_real_diff = False
                    if has_real_diff:
                        diff_indices2.add(j)
    
    return diff_indices1, diff_indices2, sentences1, sentences2

def create_html_diff(text1, text2, diff_indices1, diff_indices2):
    """Create HTML with highlighted differences"""
    def highlight_text(text, diff_indices):
        words = text.split()
        html_parts = []
        
        for i, word in enumerate(words):
            if i in diff_indices:
                html_parts.append(f'<span class="highlight">{word}</span>')
            else:
                html_parts.append(word)
        
        return ' '.join(html_parts)
    
    html1 = highlight_text(text1, diff_indices1)
    html2 = highlight_text(text2, diff_indices2)
    
    return html1, html2

def highlight_pdf_words(doc, word_data, diff_indices):
    """Highlight specific words in PDF based on indices with enhanced table support"""
    highlighted_doc = fitz.open()
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        new_page = highlighted_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.show_pdf_page(new_page.rect, doc, page_num)
        
        # Highlight words on this page
        for word_idx in diff_indices:
            if word_idx < len(word_data):
                word_info = word_data[word_idx]
                if word_info['page'] == page_num:
                    bbox = word_info['bbox']
                    rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                    
                    try:
                        # Use different color for table content
                        if word_info.get('in_table', False):
                            highlight = new_page.add_highlight_annot(rect)
                            # Light orange/peach color (RGB: 255, 200, 150)
                            highlight.set_colors(stroke=[1.0, 0.93, 0.88])
                        else:
                            highlight = new_page.add_highlight_annot(rect)
                            highlight.set_colors(stroke=fitz.utils.getColor("yellow"))
                        highlight.update()
                    except:
                        pass
    
    return highlighted_doc

def highlight_word_doc(docx_file, extracted_text, diff_indices):
    """
    Highlight words in Word document with precise run-level matching.
    """
    from docx.enum.text import WD_COLOR_INDEX
    
    docx_file.seek(0)
    doc = Document(docx_file)
    
    # Get all words from extracted text
    extracted_words = extracted_text.replace('\n\n', ' ').split()
    
    # Build a map of word positions to track what we've processed
    word_idx = 0
    
    # Process regular paragraphs
    for para in doc.paragraphs:
        para_text = para.text.strip()
        if not para_text:
            continue
        
        para_words = para_text.split()
        
        # Match this paragraph's words to extracted words
        para_start_idx = word_idx
        
        # Check if current paragraph words match extracted words at current position
        matches = True
        for i, pword in enumerate(para_words):
            if word_idx + i >= len(extracted_words):
                matches = False
                break
            if normalize_word(pword) != normalize_word(extracted_words[word_idx + i]):
                matches = False
                break
        
        if not matches:
            # Try to find where this paragraph appears in extracted words
            # This handles cases where paragraphs might be out of sync
            for search_offset in range(max(0, word_idx - 10), min(len(extracted_words), word_idx + 50)):
                temp_matches = True
                for i, pword in enumerate(para_words):
                    if search_offset + i >= len(extracted_words):
                        temp_matches = False
                        break
                    if normalize_word(pword) != normalize_word(extracted_words[search_offset + i]):
                        temp_matches = False
                        break
                if temp_matches:
                    word_idx = search_offset
                    para_start_idx = search_offset
                    break
        
        # Now highlight runs in this paragraph
        run_word_position = para_start_idx
        
        for run in para.runs:
            run_text = run.text
            if not run_text:
                continue
            
            run_words = run_text.split()
            if not run_words:
                continue
            
            # Check if ANY word in this run should be highlighted
            should_highlight = False
            for i in range(len(run_words)):
                if (run_word_position + i) in diff_indices:
                    should_highlight = True
                    break
            
            if should_highlight:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            
            run_word_position += len(run_words)
        
        # Move global word index forward
        word_idx += len(para_words)
    
    # Process tables with improved logic
    for table in doc.tables:
        for row in table.rows:
            row_cells = []
            
            # First pass: collect all cell texts for this row
            for cell in row.cells:
                cell_text = ' '.join(p.text.strip() for p in cell.paragraphs if p.text.strip())
                if cell_text:
                    row_cells.append(cell_text)
            
            # Check if we can find this row in extracted text
            if not row_cells:
                continue
            
            # The row might be represented as "cell1 | cell2 | cell3" in extracted text
            expected_row_text = ' | '.join(row_cells)
            
            # Try to find this in extracted text starting from current position
            row_found = False
            search_range = ' '.join(extracted_words[word_idx:min(word_idx + 200, len(extracted_words))])
            
            if expected_row_text in search_range:
                row_found = True
            
            # Process each cell
            for cell in row.cells:
                for para in cell.paragraphs:
                    para_text = para.text.strip()
                    if not para_text:
                        continue
                    
                    para_words = para_text.split()
                    cell_start_idx = word_idx
                    
                    # Verify alignment
                    matches = True
                    for i, cword in enumerate(para_words):
                        if word_idx + i >= len(extracted_words):
                            matches = False
                            break
                        if normalize_word(cword) != normalize_word(extracted_words[word_idx + i]):
                            matches = False
                            break
                    
                    # Highlight runs in this cell paragraph
                    run_word_position = cell_start_idx if matches else word_idx
                    
                    for run in para.runs:
                        run_text = run.text
                        if not run_text:
                            continue
                        
                        run_words = run_text.split()
                        if not run_words:
                            continue
                        
                        # Check if ANY word in this run should be highlighted
                        should_highlight = False
                        for i in range(len(run_words)):
                            if (run_word_position + i) in diff_indices:
                                should_highlight = True
                                break
                        
                        if should_highlight:
                            # Use bright green for table differences
                            run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
                        
                        run_word_position += len(run_words)
                    
                    if matches:
                        word_idx += len(para_words)
                
                # Account for cell separator in extracted text (the " | " part)
                if word_idx < len(extracted_words) and cell != row.cells[-1]:
                    # Skip the separator marker if present
                    if word_idx + 1 < len(extracted_words) and extracted_words[word_idx] == '|':
                        word_idx += 1
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

def compare_pdf_vs_word(pdf_file, word_file, pdf_is_doc1=True):
    """
    Compare PDF vs Word document without conversion.
    Extract text from both, compare, and highlight differences in original formats.
    """
    with st.spinner("Extracting text from documents..."):
        # Extract from PDF
        text_pdf, word_data_pdf, pdf_doc = extract_text_from_pdf(pdf_file)
        
        # Extract from Word
        text_word = extract_text_from_word(word_file)
    
    if not text_pdf or not text_word:
        return None
    
    # Assign texts based on which is doc1
    if pdf_is_doc1:
        text1, text2 = text_pdf, text_word
        word_data1 = word_data_pdf
        word_data2 = None
    else:
        text1, text2 = text_word, text_pdf
        word_data1 = None
        word_data2 = word_data_pdf
    
    with st.spinner("Finding differences (excluding punctuation)..."):
        diff_indices1, diff_indices2, words1, words2, sync_info = find_word_differences_optimized(text1, text2)
        html1, html2 = create_html_diff(text1, text2, diff_indices1, diff_indices2)
    
    with st.spinner("Generating highlighted documents..."):
        # Highlight PDF
        if pdf_is_doc1:
            highlighted_pdf = highlight_pdf_words(pdf_doc, word_data_pdf, diff_indices1)
            pdf_bytes = BytesIO()
            highlighted_pdf.save(pdf_bytes)
            pdf_bytes.seek(0)
            highlighted_pdf.close()
            
            # Highlight Word
            word_bytes = highlight_word_doc(word_file, text_word, diff_indices2)
            
            doc1_bytes = pdf_bytes
            doc2_bytes = word_bytes
        else:
            # Highlight Word
            word_bytes = highlight_word_doc(word_file, text_word, diff_indices1)
            
            # Highlight PDF
            highlighted_pdf = highlight_pdf_words(pdf_doc, word_data_pdf, diff_indices2)
            pdf_bytes = BytesIO()
            highlighted_pdf.save(pdf_bytes)
            pdf_bytes.seek(0)
            highlighted_pdf.close()
            
            doc1_bytes = word_bytes
            doc2_bytes = pdf_bytes
        
        pdf_doc.close()
    
    return {
        'text1': text1,
        'text2': text2,
        'diff_indices1': diff_indices1,
        'diff_indices2': diff_indices2,
        'words1': words1,
        'words2': words2,
        'html1': html1,
        'html2': html2,
        'pdf1_bytes': doc1_bytes,
        'pdf2_bytes': doc2_bytes,
        'is_pdf1': pdf_is_doc1,
        'is_pdf2': not pdf_is_doc1,
        'sync_info': sync_info,
        'is_mixed_format': True
    }

# CSS for styling
st.markdown("""
<style>
    .diff-container {
        font-family: 'Courier New', monospace;
        font-size: 13px;
        line-height: 1.8;
        padding: 20px;
        background-color: #ffffff;
        border-radius: 5px;
        max-height: 700px;
        overflow-y: auto;
        border: 1px solid #ddd;
        white-space: pre-wrap;
        word-wrap: break-word;
    }
    .highlight {
        background-color: #ffff00;
        padding: 1px 2px;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'comparison_done' not in st.session_state:
    st.session_state.comparison_done = False
if 'results' not in st.session_state:
    st.session_state.results = None

# Process and display differences
if doc1_file and doc2_file:
    # Check if files changed
    current_files = (doc1_file.name, doc2_file.name)
    if 'last_files' not in st.session_state or st.session_state.last_files != current_files:
        st.session_state.comparison_done = False
        st.session_state.last_files = current_files
    
    if not st.session_state.comparison_done:
        # Detect file types
        is_pdf1 = doc1_file.name.endswith('.pdf')
        is_pdf2 = doc2_file.name.endswith('.pdf')
        
        # Check if we have a mixed format comparison (PDF vs Word)
        is_mixed_format = (is_pdf1 and not is_pdf2) or (not is_pdf1 and is_pdf2)
        
        if is_mixed_format:
            st.info("🔄 Mixed format detected (PDF vs Word). Comparing text content directly...")
            
            if is_pdf1:
                # PDF is doc1, Word is doc2
                results = compare_pdf_vs_word(doc1_file, doc2_file, pdf_is_doc1=True)
            else:
                # Word is doc1, PDF is doc2
                results = compare_pdf_vs_word(doc2_file, doc1_file, pdf_is_doc1=False)
            
            if results:
                st.session_state.results = results
                st.session_state.comparison_done = True
        
        else:
            # Original logic for same-format comparison
            with st.spinner("Extracting text from documents..."):
                if is_pdf1:
                    text1, word_data1, pdf_doc1 = extract_text_from_pdf(doc1_file)
                else:
                    text1 = extract_text_from_word(doc1_file)
                    word_data1 = None
                    pdf_doc1 = None
                
                if is_pdf2:
                    text2, word_data2, pdf_doc2 = extract_text_from_pdf(doc2_file)
                else:
                    text2 = extract_text_from_word(doc2_file)
                    word_data2 = None
                    pdf_doc2 = None
            
            if text1 and text2:
                with st.spinner("Finding differences (excluding punctuation)..."):
                    diff_indices1, diff_indices2, words1, words2, sync_info = find_word_differences_optimized(text1, text2)
                    html1, html2 = create_html_diff(text1, text2, diff_indices1, diff_indices2)
                
                with st.spinner("Generating highlighted documents..."):
                    # Generate highlighted versions
                    if is_pdf1:
                        highlighted_doc1 = highlight_pdf_words(pdf_doc1, word_data1, diff_indices1)
                        pdf1_bytes = BytesIO()
                        highlighted_doc1.save(pdf1_bytes)
                        pdf1_bytes.seek(0)
                        highlighted_doc1.close()
                        pdf_doc1.close()
                    else:
                        pdf1_bytes = highlight_word_doc(doc1_file, text1, diff_indices1)
                    
                    if is_pdf2:
                        highlighted_doc2 = highlight_pdf_words(pdf_doc2, word_data2, diff_indices2)
                        pdf2_bytes = BytesIO()
                        highlighted_doc2.save(pdf2_bytes)
                        pdf2_bytes.seek(0)
                        highlighted_doc2.close()
                        pdf_doc2.close()
                    else:
                        pdf2_bytes = highlight_word_doc(doc2_file, text2, diff_indices2)
                
                # Store results
                st.session_state.results = {
                    'text1': text1,
                    'text2': text2,
                    'diff_indices1': diff_indices1,
                    'diff_indices2': diff_indices2,
                    'words1': words1,
                    'words2': words2,
                    'html1': html1,
                    'html2': html2,
                    'pdf1_bytes': pdf1_bytes,
                    'pdf2_bytes': pdf2_bytes,
                    'is_pdf1': is_pdf1,
                    'is_pdf2': is_pdf2,
                    'sync_info': sync_info,
                    'is_mixed_format': False
                }
                st.session_state.comparison_done = True
    
    # Display results
    if st.session_state.results:
        results = st.session_state.results
        
        st.success("✅ Comparison complete!")
        
        # Display statistics
        sync_info = results['sync_info']
        match_percentage = (sync_info['total_matching'] / max(sync_info['total_words1'], sync_info['total_words2'])) * 100
        
        st.info(f"📊 **Alignment**: {sync_info['total_matching']} matching words out of {max(sync_info['total_words1'], sync_info['total_words2'])} ({match_percentage:.1f}% match) • Punctuation differences excluded")
        
        col_stat1, col_stat2, col_stat3 = st.columns(3)
        with col_stat1:
            st.metric("Words in Doc 1", sync_info['total_words1'])
            st.metric("Different in Doc 1", sync_info['diff_words1'])
        with col_stat2:
            st.metric("Words in Doc 2", sync_info['total_words2'])
            st.metric("Different in Doc 2", sync_info['diff_words2'])
        with col_stat3:
            similarity = (sync_info['total_matching'] / max(sync_info['total_words1'], sync_info['total_words2'])) * 100
            st.metric("Match Rate", f"{similarity:.1f}%")
            st.metric("Sync Blocks", sync_info['equal_blocks'])
        
        # Download buttons
        st.markdown("### Download Highlighted Documents")
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            file_ext1 = 'pdf' if results.get('is_pdf1', False) else 'docx'
            st.download_button(
                label=f"⬇️ Download Doc 1 (Highlighted .{file_ext1})",
                data=results['pdf1_bytes'].getvalue(),
                file_name=f"doc1_highlighted.{file_ext1}",
                mime="application/pdf" if results.get('is_pdf1', False) else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
        with col_dl2:
            file_ext2 = 'pdf' if results.get('is_pdf2', False) else 'docx'
            st.download_button(
                label=f"⬇️ Download Doc 2 (Highlighted .{file_ext2})",
                data=results['pdf2_bytes'].getvalue(),
                file_name=f"doc2_highlighted.{file_ext2}",
                mime="application/pdf" if results.get('is_pdf2', False) else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
        # Display text comparison
        st.markdown("### Text Comparison Preview")
        st.markdown("🟡 **Yellow highlight** = Regular text differences | 🟠 **Light Orange highlight** (PDF) = Table differences | 🟢 **Green highlight** (Word) = Table differences | Punctuation differences excluded")
        
        col_diff1, col_diff2 = st.columns(2)
        
        with col_diff1:
            st.markdown("**Document 1**")
            st.markdown(f'<div class="diff-container">{results["html1"]}</div>', unsafe_allow_html=True)
        
        with col_diff2:
            st.markdown("**Document 2**")
            st.markdown(f'<div class="diff-container">{results["html2"]}</div>', unsafe_allow_html=True)
        
        # Sample differences
        with st.expander("📋 View Sample Differences (First 20)"):
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                st.markdown(f"**Different words in Doc 1: {len(results['diff_indices1'])} total**")
                sample_indices1 = sorted(list(results['diff_indices1']))[:20]
                for idx in sample_indices1:
                    if idx < len(results['words1']):
                        word1 = results['words1'][idx]
                        word2 = results['words2'][idx] if idx < len(results['words2']) else "N/A"
                        st.text(f"Pos {idx}: '{word1}' vs '{word2}'")
            with col_s2:
                st.markdown(f"**Different words in Doc 2: {len(results['diff_indices2'])} total**")
                sample_indices2 = sorted(list(results['diff_indices2']))[:20]
                for idx in sample_indices2:
                    if idx < len(results['words2']):
                        word2 = results['words2'][idx]
                        word1 = results['words1'][idx] if idx < len(results['words1']) else "N/A"
                        st.text(f"Pos {idx}: '{word2}' vs '{word1}'")

else:
    st.info("👆 Please upload both documents to begin comparison")

# Footer
st.markdown("---")
st.markdown("💡 **Precise word-by-word comparison** - Highlights only the words that actually differ between documents (punctuation differences excluded)")
st.markdown("🔸 For PDFs: Yellow = regular text differences, Light Orange = table differences")
st.markdown("🔸 For Word docs: Yellow = regular text differences, Green = table differences")
st.markdown("🔸 For PDF vs Word: Compares text content directly without format conversion")res punctuation, bullet points, and line break differences")
st.markdown("🔸 Bullet points like 'a.', '1.', 'i.', '(a)' are automatically excluded")
st.markdown("🔸 Line breaks within words (common in PDFs) are normalized")