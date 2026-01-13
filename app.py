import streamlit as st
import difflib
from io import BytesIO
import fitz  # PyMuPDF
from docx import Document
import re

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
    """Extract text from Word document including tables, headers, footers"""
    try:
        docx_file.seek(0)
        doc = Document(docx_file)
        
        all_elements = []
        
        # Track document structure to identify tables
        for element in doc.element.body:
            if element.tag.endswith('p'):
                for para in doc.paragraphs:
                    if para._element == element:
                        text = para.text.strip()
                        if text:
                            all_elements.append({
                                'type': 'paragraph',
                                'text': text
                            })
                        break
            elif element.tag.endswith('tbl'):
                for table in doc.tables:
                    if table._element == element:
                        table_rows = []
                        for row in table.rows:
                            row_text = []
                            for cell in row.cells:
                                cell_text = cell.text.strip()
                                if cell_text:
                                    row_text.append(cell_text)
                            if row_text:
                                table_rows.append(' | '.join(row_text))
                        if table_rows:
                            all_elements.append({
                                'type': 'table',
                                'text': '\n'.join(table_rows)
                            })
                        break
        
        # If structure extraction didn't work, fallback
        if len(all_elements) == 0:
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    all_elements.append({'type': 'paragraph', 'text': text})
            
            for table in doc.tables:
                table_rows = []
                for row in table.rows:
                    row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if row_text:
                        table_rows.append(' | '.join(row_text))
                if table_rows:
                    all_elements.append({'type': 'table', 'text': '\n'.join(table_rows)})
        
        return all_elements
    except Exception as e:
        st.error(f"Error reading Word document: {str(e)}")
        return None

def extract_text_from_pdf(pdf_file):
    """Extract text from PDF with structure information"""
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        all_elements = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Get blocks to identify tables vs text
            blocks = page.get_text("dict")["blocks"]
            
            for block in blocks:
                if "lines" in block:
                    block_text = []
                    words_info = []
                    
                    for line in block["lines"]:
                        line_text = []
                        for span in line["spans"]:
                            text = span["text"].strip()
                            if text:
                                line_text.append(text)
                                # Store word positions
                                for word in text.split():
                                    words_info.append({
                                        'word': word,
                                        'bbox': span["bbox"],
                                        'page': page_num
                                    })
                        if line_text:
                            block_text.append(' '.join(line_text))
                    
                    if block_text:
                        full_text = '\n'.join(block_text)
                        # Heuristic: if block has | or lots of numbers, likely a table
                        is_table = '|' in full_text or (len(re.findall(r'\d+', full_text)) / max(len(full_text.split()), 1)) > 0.3
                        
                        all_elements.append({
                            'type': 'table' if is_table else 'paragraph',
                            'text': full_text,
                            'words': words_info,
                            'page': page_num
                        })
        
        return all_elements, doc
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None, None

def find_differences_smart(elements1, elements2):
    """Find differences with context awareness - sentence-level for prose, word-level for tables"""
    
    # Flatten to get all text
    text1 = '\n'.join([el['text'] for el in elements1])
    text2 = '\n'.join([el['text'] for el in elements2])
    
    # Build word-to-element mapping
    word_to_element1 = []
    word_to_element2 = []
    
    for i, el in enumerate(elements1):
        words = re.findall(r'\S+', el['text'])
        word_to_element1.extend([i] * len(words))
    
    for i, el in enumerate(elements2):
        words = re.findall(r'\S+', el['text'])
        word_to_element2.extend([i] * len(words))
    
    words1 = re.findall(r'\S+', text1)
    words2 = re.findall(r'\S+', text2)
    
    # First pass: identify ONLY the actually different words
    changed_words1 = set()
    changed_words2 = set()
    
    matcher = difflib.SequenceMatcher(None, words1, words2, autojunk=False)
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'replace':
            # Words differ in both documents
            changed_words1.update(range(i1, i2))
            changed_words2.update(range(j1, j2))
        elif tag == 'delete':
            # Words only in doc1
            changed_words1.update(range(i1, i2))
        elif tag == 'insert':
            # Words only in doc2
            changed_words2.update(range(j1, j2))
        # 'equal' means the words match - don't mark anything
    
    # Second pass: expand to sentences ONLY for changed words in paragraphs
    diff_indices1 = set()
    diff_indices2 = set()
    
    for idx in changed_words1:
        if idx < len(word_to_element1):
            elem_idx = word_to_element1[idx]
            if elements1[elem_idx]['type'] == 'table':
                # Table: mark just this word
                diff_indices1.add(idx)
            else:
                # Paragraph: mark entire sentence containing this changed word
                sentence_indices = find_sentence_indices(words1, idx, word_to_element1, elem_idx)
                diff_indices1.update(sentence_indices)
    
    for idx in changed_words2:
        if idx < len(word_to_element2):
            elem_idx = word_to_element2[idx]
            if elements2[elem_idx]['type'] == 'table':
                # Table: mark just this word
                diff_indices2.add(idx)
            else:
                # Paragraph: mark entire sentence containing this changed word
                sentence_indices = find_sentence_indices(words2, idx, word_to_element2, elem_idx)
                diff_indices2.update(sentence_indices)
    
    return diff_indices1, diff_indices2, words1, words2

def find_sentence_indices(words, changed_idx, word_to_element, elem_idx):
    """Find all word indices in the sentence containing the changed word"""
    # Sentence boundaries: look for . ! ? followed by space/end, or element boundary
    start_idx = changed_idx
    end_idx = changed_idx
    
    # Search backwards for sentence start (or beginning of element)
    for i in range(changed_idx - 1, -1, -1):
        # Stop if we're in a different element
        if i >= len(word_to_element) or word_to_element[i] != elem_idx:
            start_idx = i + 1
            break
        # Stop if previous word ends with sentence terminator
        if i > 0 and i - 1 < len(words):
            prev_word = words[i - 1]
            if re.search(r'[.!?]$', prev_word):
                start_idx = i
                break
        start_idx = i
    
    # Search forwards for sentence end
    for i in range(changed_idx, len(words)):
        # Stop if we're in a different element
        if i >= len(word_to_element) or word_to_element[i] != elem_idx:
            end_idx = i - 1
            break
        end_idx = i
        # Stop after finding sentence terminator
        if re.search(r'[.!?]$', words[i]):
            break
    
    return set(range(start_idx, end_idx + 1))

def create_html_diff(text, diff_indices, words):
    """Create HTML with highlighting based on positions"""
    text_words = re.findall(r'\S+|\s+', text)
    html_parts = []
    word_index = 0
    
    for token in text_words:
        if token.strip():  # It's a word
            if word_index in diff_indices:
                html_parts.append(f'<span class="highlight">{token}</span>')
            else:
                html_parts.append(token)
            word_index += 1
        else:  # It's whitespace
            html_parts.append(token)
    
    return ''.join(html_parts)

def highlight_pdf_words(pdf_doc, elements, diff_indices, words_list):
    """Highlight specific word positions in PDF"""
    highlighted_doc = fitz.open()
    
    # Build word position mapping
    word_positions = []
    for element in elements:
        if 'words' in element:
            word_positions.extend(element['words'])
    
    for page_num in range(len(pdf_doc)):
        page = pdf_doc[page_num]
        new_page = highlighted_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.show_pdf_page(new_page.rect, pdf_doc, page_num)
        
        for word_idx in diff_indices:
            if word_idx < len(word_positions):
                word_pos = word_positions[word_idx]
                if word_pos['page'] == page_num:
                    bbox = word_pos['bbox']
                    rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                    
                    try:
                        highlight = new_page.add_highlight_annot(rect)
                        highlight.set_colors(stroke=fitz.utils.getColor("yellow"))
                        highlight.update()
                    except:
                        pass
    
    return highlighted_doc

def highlight_word_doc(docx_file, elements, diff_indices, words_list):
    """Highlight specific word positions in Word document"""
    from docx.shared import RGBColor
    from docx.enum.text import WD_COLOR_INDEX
    
    docx_file.seek(0)
    doc = Document(docx_file)
    
    # Map word indices to runs
    run_word_map = []
    current_word_idx = 0
    
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run_words = re.findall(r'\S+', run.text)
            start_idx = current_word_idx
            end_idx = current_word_idx + len(run_words)
            
            should_highlight = any(i in diff_indices for i in range(start_idx, end_idx))
            run_word_map.append({
                'run': run,
                'should_highlight': should_highlight
            })
            current_word_idx += len(run_words)
    
    # Apply highlighting
    for run_info in run_word_map:
        if run_info['should_highlight']:
            run_info['run'].font.highlight_color = WD_COLOR_INDEX.YELLOW
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

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
    .stDownloadButton button {
        pointer-events: auto !important;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'comparison_done' not in st.session_state:
    st.session_state.comparison_done = False
if 'results' not in st.session_state:
    st.session_state.results = None

# Process documents
if doc1_file and doc2_file:
    current_files = (doc1_file.name, doc2_file.name, doc1_file.size, doc2_file.size)
    if 'last_files' not in st.session_state or st.session_state.last_files != current_files:
        st.session_state.comparison_done = False
        st.session_state.last_files = current_files
    
    if not st.session_state.comparison_done:
        with st.spinner("Extracting text from documents..."):
            is_pdf1 = doc1_file.name.endswith('.pdf')
            is_pdf2 = doc2_file.name.endswith('.pdf')
            
            if is_pdf1:
                elements1, pdf_doc1 = extract_text_from_pdf(doc1_file)
            else:
                elements1 = extract_text_from_word(doc1_file)
                pdf_doc1 = None
            
            if is_pdf2:
                elements2, pdf_doc2 = extract_text_from_pdf(doc2_file)
            else:
                elements2 = extract_text_from_word(doc2_file)
                pdf_doc2 = None
        
        if elements1 and elements2:
            with st.spinner("Finding differences (smart mode: sentences for prose, words for tables)..."):
                diff_indices1, diff_indices2, words1, words2 = find_differences_smart(elements1, elements2)
                
                text1 = '\n'.join([el['text'] for el in elements1])
                text2 = '\n'.join([el['text'] for el in elements2])
                
                html1 = create_html_diff(text1, diff_indices1, words1)
                html2 = create_html_diff(text2, diff_indices2, words2)
            
            with st.spinner("Generating highlighted documents..."):
                if is_pdf1:
                    highlighted_doc1 = highlight_pdf_words(pdf_doc1, elements1, diff_indices1, words1)
                    pdf1_bytes = BytesIO()
                    highlighted_doc1.save(pdf1_bytes)
                    pdf1_bytes.seek(0)
                    highlighted_doc1.close()
                    if pdf_doc1:
                        pdf_doc1.close()
                else:
                    pdf1_bytes = highlight_word_doc(doc1_file, elements1, diff_indices1, words1)
                
                if is_pdf2:
                    highlighted_doc2 = highlight_pdf_words(pdf_doc2, elements2, diff_indices2, words2)
                    pdf2_bytes = BytesIO()
                    highlighted_doc2.save(pdf2_bytes)
                    pdf2_bytes.seek(0)
                    highlighted_doc2.close()
                    if pdf_doc2:
                        pdf_doc2.close()
                else:
                    pdf2_bytes = highlight_word_doc(doc2_file, elements2, diff_indices2, words2)
            
            st.session_state.results = {
                'text1': text1,
                'text2': text2,
                'html1': html1,
                'html2': html2,
                'pdf1_bytes': pdf1_bytes,
                'pdf2_bytes': pdf2_bytes,
                'is_pdf1': is_pdf1,
                'is_pdf2': is_pdf2,
                'diff_count1': len(diff_indices1),
                'diff_count2': len(diff_indices2)
            }
            st.session_state.comparison_done = True
    
    if st.session_state.results:
        results = st.session_state.results
        
        st.success("✅ Comparison complete!")
        
        # Statistics
        col_stat1, col_stat2, col_stat3 = st.columns(3)
        with col_stat1:
            st.metric("Words in Doc 1", len(re.findall(r'\S+', results['text1'])))
        with col_stat2:
            st.metric("Words in Doc 2", len(re.findall(r'\S+', results['text2'])))
        with col_stat3:
            similarity = difflib.SequenceMatcher(None, results['text1'], results['text2']).ratio()
            st.metric("Similarity", f"{similarity * 100:.1f}%")
        
        st.info(f"🔍 **Smart highlighting**: Sentences in prose, words in tables | {results['diff_count1']} highlights in Doc 1, {results['diff_count2']} in Doc 2")
        
        # Download buttons with unique keys to prevent refresh
        st.markdown("### Download Highlighted Documents")
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            file_ext1 = 'pdf' if results['is_pdf1'] else 'docx'
            st.download_button(
                label=f"⬇️ Download Doc 1 (Highlighted .{file_ext1})",
                data=results['pdf1_bytes'].getvalue(),
                file_name=f"doc1_highlighted.{file_ext1}",
                mime="application/pdf" if results['is_pdf1'] else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_doc1",
                use_container_width=True
            )
        
        with col_dl2:
            file_ext2 = 'pdf' if results['is_pdf2'] else 'docx'
            st.download_button(
                label=f"⬇️ Download Doc 2 (Highlighted .{file_ext2})",
                data=results['pdf2_bytes'].getvalue(),
                file_name=f"doc2_highlighted.{file_ext2}",
                mime="application/pdf" if results['is_pdf2'] else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_doc2",
                use_container_width=True
            )
        
        # Display comparison
        st.markdown("### Text Comparison Preview")
        st.markdown("🟡 **Yellow highlight** = Differences (sentences for prose, words for tables)")
        
        col_diff1, col_diff2 = st.columns(2)
        
        with col_diff1:
            st.markdown("**Document 1**")
            st.markdown(f'<div class="diff-container">{results["html1"]}</div>', unsafe_allow_html=True)
        
        with col_diff2:
            st.markdown("**Document 2**")
            st.markdown(f'<div class="diff-container">{results["html2"]}</div>', unsafe_allow_html=True)

else:
    st.info("👆 Please upload both documents to begin comparison")

st.markdown("---")
st.markdown("💡 **Smart mode**: Highlights entire sentences in prose, individual words in tables")