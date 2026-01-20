import streamlit as st
import difflib
from io import BytesIO
import fitz  # PyMuPDF
from docx import Document
import re
from collections import Counter

st.set_page_config(page_title="Document Diff Checker", layout="wide")

st.title("📄 Document Diff Checker")
st.markdown("Upload two documents (PDF or Word) to compare and highlight their differences")

# Configuration options
with st.sidebar:
    st.header("⚙️ Comparison Settings")
    
    ignore_case = st.checkbox("Ignore case differences", value=True)
    ignore_punctuation = st.checkbox("Ignore punctuation differences", value=False)
    min_word_length = st.slider("Minimum word length to compare", 1, 10, 3, 
                                help="Ignore very short words (like 'a', 'to', 'in') to reduce noise")
    context_window = st.slider("Context matching window", 3, 20, 8,
                              help="Number of surrounding words to consider for alignment")
    
    st.markdown("---")
    st.markdown("**Common words filter**")
    filter_common = st.checkbox("Filter out common formatting words", value=True,
                               help="Ignore page numbers, headers, footers, and common words")
    
    if filter_common:
        common_words_input = st.text_area(
            "Additional words to ignore (one per line)",
            value="page\nof\nthe\nand\na\nan\nin\nto\nfor\nis",
            height=150
        )
        custom_ignore_words = set(word.strip().lower() for word in common_words_input.split('\n') if word.strip())
    else:
        custom_ignore_words = set()

# Create two columns for file uploads
col1, col2 = st.columns(2)

with col1:
    st.subheader("Document 1")
    doc1_file = st.file_uploader("Upload first document", type=['pdf', 'docx'], key="doc1")
    
with col2:
    st.subheader("Document 2")
    doc2_file = st.file_uploader("Upload second document", type=['pdf', 'docx'], key="doc2")

def normalize_text(text, ignore_case=True, ignore_punct=False):
    """Normalize text for better comparison"""
    if ignore_case:
        text = text.lower()
    if ignore_punct:
        text = re.sub(r'[^\w\s]', '', text)
    return text

def is_likely_header_footer(text, position_ratio):
    """Detect if text is likely a header or footer based on position and content"""
    text_lower = text.lower().strip()
    
    # Position-based detection (top 10% or bottom 10% of page)
    if position_ratio < 0.1 or position_ratio > 0.9:
        # Common header/footer patterns
        if re.match(r'^\d+$', text_lower):  # Just a number
            return True
        if re.match(r'^page\s*\d+', text_lower):  # "Page X"
            return True
        if re.match(r'^\d+\s*of\s*\d+$', text_lower):  # "X of Y"
            return True
        if len(text_lower) < 50 and any(word in text_lower for word in ['page', 'chapter', 'section']):
            return True
    
    return False

def extract_text_from_word(docx_file, filter_metadata=False):
    """Extract text from Word document with metadata filtering"""
    try:
        docx_file.seek(0)
        doc = Document(docx_file)
        
        all_text = []
        word_metadata = []
        
        # Extract from paragraphs
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                all_text.append(text)
                word_metadata.append({'type': 'paragraph', 'text': text})
        
        # Extract from tables
        for table in doc.tables:
            for row in table.rows:
                row_text = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text:
                        row_text.append(cell_text)
                if row_text:
                    combined = ' '.join(row_text)
                    all_text.append(combined)
                    word_metadata.append({'type': 'table', 'text': combined})
        
        full_text = '\n'.join(all_text)
        
        return full_text, word_metadata
    except Exception as e:
        st.error(f"Error reading Word document: {str(e)}")
        return None, None

def extract_text_from_pdf(pdf_file, filter_metadata=False):
    """Extract text from PDF with position-based metadata filtering"""
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        pages_data = []
        full_text = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            page_height = page.rect.height
            
            # Get text blocks with positions
            blocks = page.get_text("dict")["blocks"]
            
            page_content = []
            page_words = []
            
            for block in blocks:
                if "lines" in block:
                    # Calculate relative position on page
                    block_y = block['bbox'][1]
                    position_ratio = block_y / page_height
                    
                    for line in block["lines"]:
                        line_text = ""
                        for span in line["spans"]:
                            text = span["text"].strip()
                            if text:
                                line_text += text + " "
                                
                                # Store word positions for highlighting
                                words = text.split()
                                bbox = span["bbox"]
                                char_width = (bbox[2] - bbox[0]) / len(text) if len(text) > 0 else 0
                                x_pos = bbox[0]
                                
                                for word in words:
                                    if word.strip():
                                        word_bbox = [x_pos, bbox[1], x_pos + len(word) * char_width, bbox[3]]
                                        
                                        # Check if likely header/footer
                                        is_metadata = filter_metadata and is_likely_header_footer(word, position_ratio)
                                        
                                        page_words.append({
                                            'text': word,
                                            'bbox': word_bbox,
                                            'page': page_num,
                                            'is_metadata': is_metadata
                                        })
                                        x_pos += (len(word) + 1) * char_width
                        
                        if line_text.strip():
                            is_metadata = filter_metadata and is_likely_header_footer(line_text, position_ratio)
                            if not is_metadata or not filter_metadata:
                                page_content.append(line_text.strip())
            
            pages_data.append({
                'page_num': page_num,
                'words': page_words,
                'content': page_content
            })
            
            full_text.append('\n'.join(page_content))
        
        return '\n\n'.join(full_text), pages_data, doc
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None, None, None

def clean_word_for_comparison(word, ignore_case, ignore_punct, min_length, ignore_words):
    """Clean and filter words for comparison"""
    # Normalize
    cleaned = normalize_text(word, ignore_case, ignore_punct)
    
    # Filter by length
    if len(cleaned) < min_length:
        return None
    
    # Filter common words
    if cleaned in ignore_words:
        return None
    
    # Filter numbers-only
    if cleaned.isdigit():
        return None
    
    return cleaned

def find_semantic_differences(text1, text2, ignore_case, ignore_punct, min_length, ignore_words, context_window):
    """Find meaningful differences using context-aware matching"""
    
    # Tokenize and clean
    raw_words1 = re.findall(r'\S+', text1)
    raw_words2 = re.findall(r'\S+', text2)
    
    # Create cleaned versions for comparison
    words1_cleaned = []
    words1_mapping = []  # Maps cleaned index to raw index
    
    for idx, word in enumerate(raw_words1):
        cleaned = clean_word_for_comparison(word, ignore_case, ignore_punct, min_length, ignore_words)
        if cleaned:
            words1_cleaned.append(cleaned)
            words1_mapping.append(idx)
    
    words2_cleaned = []
    words2_mapping = []
    
    for idx, word in enumerate(raw_words2):
        cleaned = clean_word_for_comparison(word, ignore_case, ignore_punct, min_length, ignore_words)
        if cleaned:
            words2_cleaned.append(cleaned)
            words2_mapping.append(idx)
    
    # Use SequenceMatcher on cleaned words
    matcher = difflib.SequenceMatcher(None, words1_cleaned, words2_cleaned, autojunk=False)
    
    # Find differences in cleaned word space
    diff_cleaned_idx1 = set()
    diff_cleaned_idx2 = set()
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag in ['replace', 'delete']:
            diff_cleaned_idx1.update(range(i1, i2))
        if tag in ['replace', 'insert']:
            diff_cleaned_idx2.update(range(j1, j2))
    
    # Map back to raw word indices
    diff_indices1 = set(words1_mapping[i] for i in diff_cleaned_idx1 if i < len(words1_mapping))
    diff_indices2 = set(words2_mapping[i] for i in diff_cleaned_idx2 if i < len(words2_mapping))
    
    # Calculate statistics
    total_compared1 = len(words1_cleaned)
    total_compared2 = len(words2_cleaned)
    diff_count1 = len(diff_cleaned_idx1)
    diff_count2 = len(diff_cleaned_idx2)
    
    return diff_indices1, diff_indices2, raw_words1, raw_words2, {
        'total_compared1': total_compared1,
        'total_compared2': total_compared2,
        'diff_count1': diff_count1,
        'diff_count2': diff_count2,
        'filtered_out1': len(raw_words1) - total_compared1,
        'filtered_out2': len(raw_words2) - total_compared2
    }

def create_html_diff(text1, text2, diff_indices1, diff_indices2):
    """Create HTML with highlighting"""
    def highlight_text(text, diff_indices):
        text_words = re.findall(r'\S+|\s+', text)
        html_parts = []
        word_index = 0
        
        for token in text_words:
            if token.strip():
                if word_index in diff_indices:
                    html_parts.append(f'<span class="highlight">{token}</span>')
                else:
                    html_parts.append(token)
                word_index += 1
            else:
                html_parts.append(token)
        
        return ''.join(html_parts)
    
    return highlight_text(text1, diff_indices1), highlight_text(text2, diff_indices2)

def highlight_pdf_words(doc, pages_data, diff_indices):
    """Highlight words in PDF"""
    highlighted_doc = fitz.open()
    
    word_positions = []
    for page_data in pages_data:
        for word_info in page_data['words']:
            word_positions.append(word_info)
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        new_page = highlighted_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.show_pdf_page(new_page.rect, doc, page_num)
        
        for word_idx in diff_indices:
            if word_idx < len(word_positions):
                word_pos = word_positions[word_idx]
                if word_pos['page'] == page_num and not word_pos.get('is_metadata', False):
                    bbox = word_pos['bbox']
                    rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                    
                    try:
                        highlight = new_page.add_highlight_annot(rect)
                        highlight.set_colors(stroke=fitz.utils.getColor("yellow"))
                        highlight.update()
                    except:
                        pass
    
    return highlighted_doc

def highlight_word_doc(docx_file, diff_indices, words_list):
    """Highlight words in Word document"""
    from docx.shared import RGBColor
    from docx.enum.text import WD_COLOR_INDEX
    
    docx_file.seek(0)
    doc = Document(docx_file)
    
    current_word_idx = 0
    
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run_words = re.findall(r'\S+', run.text)
            
            # Check if any word in this run should be highlighted
            should_highlight = any(
                current_word_idx + i in diff_indices 
                for i in range(len(run_words))
            )
            
            if should_highlight:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            
            current_word_idx += len(run_words)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# CSS
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

# Main processing
if doc1_file and doc2_file:
    if st.button("🔍 Compare Documents", type="primary"):
        with st.spinner("Extracting and comparing documents..."):
            # Extract text
            is_pdf1 = doc1_file.name.endswith('.pdf')
            is_pdf2 = doc2_file.name.endswith('.pdf')
            
            if is_pdf1:
                text1, pages_data1, pdf_doc1 = extract_text_from_pdf(doc1_file, filter_metadata=filter_common)
            else:
                text1, word_meta1 = extract_text_from_word(doc1_file, filter_metadata=filter_common)
                pages_data1 = None
                pdf_doc1 = None
            
            if is_pdf2:
                text2, pages_data2, pdf_doc2 = extract_text_from_pdf(doc2_file, filter_metadata=filter_common)
            else:
                text2, word_meta2 = extract_text_from_word(doc2_file, filter_metadata=filter_common)
                pages_data2 = None
                pdf_doc2 = None
            
            if text1 and text2:
                # Find differences
                diff_indices1, diff_indices2, words1, words2, stats = find_semantic_differences(
                    text1, text2, ignore_case, ignore_punctuation, 
                    min_word_length, custom_ignore_words, context_window
                )
                
                # Create HTML preview
                html1, html2 = create_html_diff(text1, text2, diff_indices1, diff_indices2)
                
                # Generate highlighted documents
                if is_pdf1:
                    highlighted_doc1 = highlight_pdf_words(pdf_doc1, pages_data1, diff_indices1)
                    pdf1_bytes = BytesIO()
                    highlighted_doc1.save(pdf1_bytes)
                    pdf1_bytes.seek(0)
                    highlighted_doc1.close()
                    pdf_doc1.close()
                else:
                    pdf1_bytes = highlight_word_doc(doc1_file, diff_indices1, words1)
                
                if is_pdf2:
                    highlighted_doc2 = highlight_pdf_words(pdf_doc2, pages_data2, diff_indices2)
                    pdf2_bytes = BytesIO()
                    highlighted_doc2.save(pdf2_bytes)
                    pdf2_bytes.seek(0)
                    highlighted_doc2.close()
                    pdf_doc2.close()
                else:
                    pdf2_bytes = highlight_word_doc(doc2_file, diff_indices2, words2)
                
                # Display results
                st.success("✅ Comparison complete!")
                
                # Statistics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Doc 1 Words", len(words1))
                with col2:
                    st.metric("Doc 2 Words", len(words2))
                with col3:
                    st.metric("Compared Words", f"{stats['total_compared1']} / {stats['total_compared2']}")
                with col4:
                    diff_pct = (stats['diff_count1'] / max(stats['total_compared1'], 1)) * 100
                    st.metric("Different", f"{diff_pct:.1f}%")
                
                st.info(f"📊 Filtered out {stats['filtered_out1']} words from Doc 1 and {stats['filtered_out2']} words from Doc 2 "
                       f"(short words, common words, numbers)")
                
                # Downloads
                st.markdown("### Download Highlighted Documents")
                col_dl1, col_dl2 = st.columns(2)
                
                with col_dl1:
                    ext1 = 'pdf' if is_pdf1 else 'docx'
                    st.download_button(
                        f"⬇️ Download Doc 1 (.{ext1})",
                        data=pdf1_bytes.getvalue(),
                        file_name=f"doc1_highlighted.{ext1}",
                        mime="application/pdf" if is_pdf1 else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                
                with col_dl2:
                    ext2 = 'pdf' if is_pdf2 else 'docx'
                    st.download_button(
                        f"⬇️ Download Doc 2 (.{ext2})",
                        data=pdf2_bytes.getvalue(),
                        file_name=f"doc2_highlighted.{ext2}",
                        mime="application/pdf" if is_pdf2 else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                
                # Preview
                st.markdown("### Text Comparison Preview")
                col_p1, col_p2 = st.columns(2)
                
                with col_p1:
                    st.markdown("**Document 1**")
                    st.markdown(f'<div class="diff-container">{html1}</div>', unsafe_allow_html=True)
                
                with col_p2:
                    st.markdown("**Document 2**")
                    st.markdown(f'<div class="diff-container">{html2}</div>', unsafe_allow_html=True)
            else:
                st.error("Could not extract text from one or both documents")
else:
    st.info("👆 Upload both documents and configure settings, then click Compare")