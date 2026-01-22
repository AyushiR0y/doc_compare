import streamlit as st
import difflib
from io import BytesIO
import fitz  # PyMuPDF
from docx import Document
import re
import subprocess
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

def convert_word_to_pdf(docx_file):
    """Convert Word document to PDF using LibreOffice"""
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            # Save the docx to temp file
            docx_path = os.path.join(tmpdir, "input.docx")
            with open(docx_path, 'wb') as f:
                docx_file.seek(0)
                f.write(docx_file.read())
            
            # Convert using LibreOffice
            result = subprocess.run(
                ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, docx_path],
                capture_output=True,
                timeout=30
            )
            
            if result.returncode != 0:
                st.error(f"Conversion failed: {result.stderr.decode()}")
                return None
            
            # Read the generated PDF
            pdf_path = os.path.join(tmpdir, "input.pdf")
            if os.path.exists(pdf_path):
                with open(pdf_path, 'rb') as f:
                    pdf_bytes = BytesIO(f.read())
                return pdf_bytes
            else:
                st.error("PDF file was not generated")
                return None
                
    except Exception as e:
        st.error(f"Error converting Word to PDF: {str(e)}")
        return None

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
    """Normalize word by removing punctuation and converting to lowercase for comparison"""
    # Remove common punctuation but preserve the word
    word = word.strip('.,;:!?"\'-()[]{}')
    return word.lower()

def normalize_for_comparison(text):
    """Normalize text for comparison - removes punctuation differences"""
    # Split into words
    words = text.split()
    # Normalize each word
    normalized = [normalize_word(w) for w in words if normalize_word(w)]
    return normalized

def find_word_differences_optimized(text1, text2):
    """
    Optimized word-level difference detection with better alignment handling.
    Uses a more sophisticated approach to handle insertions and deletions.
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
    
    # Sets to store indices of different words
    diff_indices1 = set()
    diff_indices2 = set()
    
    # Create mapping from normalized to original indices
    # This is needed because normalized list may be shorter (punctuation removed)
    norm_to_orig1 = []
    norm_to_orig2 = []
    
    for i, word in enumerate(words1):
        if normalize_word(word):
            norm_to_orig1.append(i)
    
    for i, word in enumerate(words2):
        if normalize_word(word):
            norm_to_orig2.append(i)
    
    # Process each operation with refined logic
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
                # Map back to original indices
                for idx in range(i1, i2):
                    if idx < len(norm_to_orig1):
                        diff_indices1.add(norm_to_orig1[idx])
                for idx in range(j1, j2):
                    if idx < len(norm_to_orig2):
                        diff_indices2.add(norm_to_orig2[idx])
        elif tag == 'delete':
            # Words only in doc1
            for idx in range(i1, i2):
                if idx < len(norm_to_orig1):
                    diff_indices1.add(norm_to_orig1[idx])
        elif tag == 'insert':
            # Words only in doc2
            for idx in range(j1, j2):
                if idx < len(norm_to_orig2):
                    diff_indices2.add(norm_to_orig2[idx])
    
    # Calculate statistics
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
    Find differences at sentence level, excluding punctuation differences.
    Returns sets of sentence indices that differ.
    """
    # Split into sentences
    sentences1 = re.split(r'[.!?]+\s+', text1)
    sentences2 = re.split(r'[.!?]+\s+', text2)
    
    # Normalize sentences for comparison
    normalized1 = [' '.join(normalize_for_comparison(s)) for s in sentences1]
    normalized2 = [' '.join(normalize_for_comparison(s)) for s in sentences2]
    
    # Use SequenceMatcher
    matcher = difflib.SequenceMatcher(None, normalized1, normalized2, autojunk=False)
    opcodes = matcher.get_opcodes()
    
    diff_indices1 = set()
    diff_indices2 = set()
    
    for tag, i1, i2, j1, j2 in opcodes:
        if tag != 'equal':
            diff_indices1.update(range(i1, i2))
            diff_indices2.update(range(j1, j2))
    
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

def highlight_pdf_sentences(doc, word_data, sentence_diff_indices, sentences):
    """Highlight entire sentences in PDF"""
    highlighted_doc = fitz.open()
    
    # Build a word-to-sentence mapping
    word_to_sentence = []
    word_count = 0
    
    for sent_idx, sentence in enumerate(sentences):
        sent_words = sentence.split()
        for _ in sent_words:
            word_to_sentence.append(sent_idx)
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        new_page = highlighted_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.show_pdf_page(new_page.rect, doc, page_num)
        
        # Highlight words that belong to different sentences
        for word_idx, word_info in enumerate(word_data):
            if word_info['page'] == page_num:
                if word_idx < len(word_to_sentence):
                    sent_idx = word_to_sentence[word_idx]
                    if sent_idx in sentence_diff_indices:
                        bbox = word_info['bbox']
                        rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                        
                        try:
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
    Compare PDF vs Word document by converting Word to PDF first.
    Returns comparison results using the sync strategy.
    """
    with st.spinner("Converting Word document to PDF..."):
        converted_pdf = convert_word_to_pdf(word_file)
        if not converted_pdf:
            return None
    
    with st.spinner("Extracting text from both PDFs..."):
        # Extract from original PDF
        if pdf_is_doc1:
            text1, word_data1, pdf_doc1 = extract_text_from_pdf(pdf_file)
            text2, word_data2, pdf_doc2 = extract_text_from_pdf(converted_pdf)
        else:
            text1, word_data1, pdf_doc1 = extract_text_from_pdf(converted_pdf)
            text2, word_data2, pdf_doc2 = extract_text_from_pdf(pdf_file)
    
    if not text1 or not text2:
        return None
    
    with st.spinner("Finding differences (excluding punctuation)..."):
        # Word-level comparison
        diff_indices1, diff_indices2, words1, words2, sync_info = find_word_differences_optimized(text1, text2)
        html1, html2 = create_html_diff(text1, text2, diff_indices1, diff_indices2)
        
        # Sentence-level comparison
        sent_diff1, sent_diff2, sentences1, sentences2 = find_sentence_differences(text1, text2)
    
    with st.spinner("Generating highlighted documents..."):
        # Generate word-level highlights
        highlighted_doc1_words = highlight_pdf_words(pdf_doc1, word_data1, diff_indices1)
        highlighted_doc2_words = highlight_pdf_words(pdf_doc2, word_data2, diff_indices2)
        
        # Generate sentence-level highlights
        highlighted_doc1_sents = highlight_pdf_sentences(pdf_doc1, word_data1, sent_diff1, sentences1)
        highlighted_doc2_sents = highlight_pdf_sentences(pdf_doc2, word_data2, sent_diff2, sentences2)
        
        # Save to BytesIO
        pdf1_words_bytes = BytesIO()
        highlighted_doc1_words.save(pdf1_words_bytes)
        pdf1_words_bytes.seek(0)
        
        pdf2_words_bytes = BytesIO()
        highlighted_doc2_words.save(pdf2_words_bytes)
        pdf2_words_bytes.seek(0)
        
        pdf1_sents_bytes = BytesIO()
        highlighted_doc1_sents.save(pdf1_sents_bytes)
        pdf1_sents_bytes.seek(0)
        
        pdf2_sents_bytes = BytesIO()
        highlighted_doc2_sents.save(pdf2_sents_bytes)
        pdf2_sents_bytes.seek(0)
        
        # Cleanup
        highlighted_doc1_words.close()
        highlighted_doc2_words.close()
        highlighted_doc1_sents.close()
        highlighted_doc2_sents.close()
        pdf_doc1.close()
        pdf_doc2.close()
    
    return {
        'text1': text1,
        'text2': text2,
        'diff_indices1': diff_indices1,
        'diff_indices2': diff_indices2,
        'words1': words1,
        'words2': words2,
        'html1': html1,
        'html2': html2,
        'pdf1_words_bytes': pdf1_words_bytes,
        'pdf2_words_bytes': pdf2_words_bytes,
        'pdf1_sents_bytes': pdf1_sents_bytes,
        'pdf2_sents_bytes': pdf2_sents_bytes,
        'sync_info': sync_info,
        'is_mixed_format': True,
        'pdf_is_doc1': pdf_is_doc1
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
            st.info("🔄 Mixed format detected (PDF vs Word). Converting Word to PDF for comparison...")
            
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
        
        if results.get('is_mixed_format', False):
            # Mixed format comparison - provide both word and sentence level highlights
            st.markdown("**Word-level highlights:**")
            col_dl1, col_dl2 = st.columns(2)
            
            with col_dl1:
                st.download_button(
                    label=f"⬇️ Download Doc 1 - Words (Highlighted PDF)",
                    data=results['pdf1_words_bytes'].getvalue(),
                    file_name=f"doc1_highlighted_words.pdf",
                    mime="application/pdf"
                )
            
            with col_dl2:
                st.download_button(
                    label=f"⬇️ Download Doc 2 - Words (Highlighted PDF)",
                    data=results['pdf2_words_bytes'].getvalue(),
                    file_name=f"doc2_highlighted_words.pdf",
                    mime="application/pdf"
                )
            
            st.markdown("**Sentence-level highlights:**")
            col_dl3, col_dl4 = st.columns(2)
            
            with col_dl3:
                st.download_button(
                    label=f"⬇️ Download Doc 1 - Sentences (Highlighted PDF)",
                    data=results['pdf1_sents_bytes'].getvalue(),
                    file_name=f"doc1_highlighted_sentences.pdf",
                    mime="application/pdf"
                )
            
            with col_dl4:
                st.download_button(
                    label=f"⬇️ Download Doc 2 - Sentences (Highlighted PDF)",
                    data=results['pdf2_sents_bytes'].getvalue(),
                    file_name=f"doc2_highlighted_sentences.pdf",
                    mime="application/pdf"
                )
        else:
            # Same format comparison - original download buttons
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
        if results.get('is_mixed_format', False):
            st.markdown("🟡 **Yellow highlight** = Differences (punctuation differences excluded) | Mixed format: Word document converted to PDF for comparison")
        else:
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
st.markdown("🔸 For PDF vs Word: Automatically converts Word to PDF and provides both word-level and sentence-level highlights")