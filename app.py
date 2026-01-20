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
    """Extract text from Word document in the exact same order as we'll process for highlighting"""
    try:
        docx_file.seek(0)
        doc = Document(docx_file)
        
        all_text = []
        
        # Extract from paragraphs
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                all_text.append(text)
        
        # Extract from tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        text = para.text.strip()
                        if text:
                            all_text.append(text)
        
        extracted_text = '\n'.join(all_text)
        
        word_count = len(extracted_text.split())
        if len(all_text) == 0:
            st.warning("⚠️ No text extracted from Word document.")
        elif word_count < 10:
            st.warning(f"⚠️ Only {word_count} words extracted from Word document.")
        
        return extracted_text
    except Exception as e:
        st.error(f"Error reading Word document: {str(e)}")
        return None

def extract_text_from_pdf(pdf_file):
    """Extract text from PDF with word-level coordinates"""
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        all_words = []
        full_text_parts = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Get word-level data with coordinates
            words_data = page.get_text("words")  # Returns list of (x0, y0, x1, y1, "word", block_no, line_no, word_no)
            
            for word_info in words_data:
                x0, y0, x1, y1, word_text = word_info[:5]
                
                if word_text.strip():
                    all_words.append({
                        'text': word_text.strip(),
                        'bbox': [x0, y0, x1, y1],
                        'page': page_num
                    })
                    full_text_parts.append(word_text.strip())
        
        # Join with spaces to create full text
        full_text = ' '.join(full_text_parts)
        
        return full_text, all_words, doc
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None, None, None

def normalize_word(word):
    """Normalize word by removing punctuation and converting to lowercase for comparison"""
    # Remove common punctuation from edges
    word = word.strip('.,;:!?"\'-()[]{}')
    return word.lower()

def find_word_differences_with_sync(text1, text2):
    """
    Find word-level differences with proper syncing.
    Ignores case and punctuation differences.
    Returns sets of word indices that should be highlighted in each document.
    """
    # Split into words
    words1 = text1.split()
    words2 = text2.split()
    
    # Create normalized versions for comparison
    normalized1 = [normalize_word(w) for w in words1]
    normalized2 = [normalize_word(w) for w in words2]
    
    # Use SequenceMatcher on normalized words
    matcher = difflib.SequenceMatcher(None, normalized1, normalized2, autojunk=False)
    
    # Sets to store indices of different words
    diff_indices1 = set()
    diff_indices2 = set()
    
    # Process each operation
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            # Words match - don't highlight
            continue
        elif tag == 'replace':
            # Words are different - highlight all in this range
            diff_indices1.update(range(i1, i2))
            diff_indices2.update(range(j1, j2))
        elif tag == 'delete':
            # Words only in doc1 - highlight them
            diff_indices1.update(range(i1, i2))
        elif tag == 'insert':
            # Words only in doc2 - highlight them
            diff_indices2.update(range(j1, j2))
    
    # Calculate statistics
    total_equal_blocks = sum(1 for tag, _, _, _, _ in matcher.get_opcodes() if tag == 'equal')
    total_matching_words = sum(i2 - i1 for tag, i1, i2, _, _ in matcher.get_opcodes() if tag == 'equal')
    
    sync_info = {
        'total_matching': total_matching_words,
        'total_words1': len(words1),
        'total_words2': len(words2),
        'equal_blocks': total_equal_blocks,
        'diff_words1': len(diff_indices1),
        'diff_words2': len(diff_indices2)
    }
    
    return diff_indices1, diff_indices2, words1, words2, sync_info

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
    """Highlight specific words in PDF based on indices"""
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
                        highlight = new_page.add_highlight_annot(rect)
                        highlight.set_colors(stroke=fitz.utils.getColor("yellow"))
                        highlight.update()
                    except:
                        pass
    
    return highlighted_doc

def highlight_word_doc(docx_file, text, diff_indices):
    """Highlight specific words in Word document based on indices"""
    from docx.shared import RGBColor
    from docx.enum.text import WD_COLOR_INDEX
    
    docx_file.seek(0)
    doc = Document(docx_file)
    
    # Get list of all words from extracted text
    all_words = text.split()
    
    # Debug: let's see what we're working with
    st.write(f"DEBUG: Total words in extracted text: {len(all_words)}")
    st.write(f"DEBUG: Total diff indices: {len(diff_indices)}")
    if diff_indices:
        st.write(f"DEBUG: First 10 diff indices: {sorted(list(diff_indices))[:10]}")
        st.write(f"DEBUG: First 10 diff words: {[all_words[i] for i in sorted(list(diff_indices))[:10] if i < len(all_words)]}")
    
    # Track which word index we're at globally
    current_word_idx = 0
    
    # Process all paragraphs
    for para_num, paragraph in enumerate(doc.paragraphs):
        for run_num, run in enumerate(paragraph.runs):
            run_text = run.text
            if not run_text.strip():
                continue
            
            # Count words in this run
            run_words = run_text.split()
            num_words = len(run_words)
            
            if num_words == 0:
                continue
            
            # Check if ANY word in this run's range needs highlighting
            run_start_idx = current_word_idx
            run_end_idx = current_word_idx + num_words
            
            # Get the actual words from extracted text for this range
            extracted_words_in_range = all_words[run_start_idx:run_end_idx] if run_end_idx <= len(all_words) else []
            
            should_highlight = any(i in diff_indices for i in range(run_start_idx, run_end_idx))
            
            if should_highlight:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                st.write(f"DEBUG: Highlighting para {para_num}, run {run_num}, words {run_start_idx}-{run_end_idx}: '{run_text[:50]}'")
                st.write(f"       Extracted words: {extracted_words_in_range}")
            
            current_word_idx += num_words
    
    # Also process tables
    for table_num, table in enumerate(doc.tables):
        for row_num, row in enumerate(table.rows):
            for cell_num, cell in row.cells:
                for para_num, paragraph in enumerate(cell.paragraphs):
                    for run_num, run in enumerate(paragraph.runs):
                        run_text = run.text
                        if not run_text.strip():
                            continue
                        
                        run_words = run_text.split()
                        num_words = len(run_words)
                        
                        if num_words == 0:
                            continue
                        
                        run_start_idx = current_word_idx
                        run_end_idx = current_word_idx + num_words
                        
                        should_highlight = any(i in diff_indices for i in range(run_start_idx, run_end_idx))
                        
                        if should_highlight:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                            st.write(f"DEBUG: Highlighting table {table_num}, row {row_num}, cell {cell_num}, para {para_num}, run {run_num}")
                        
                        current_word_idx += num_words
    
    st.write(f"DEBUG: Final word count after processing: {current_word_idx}")
    st.write(f"DEBUG: Expected word count: {len(all_words)}")
    
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
        with st.spinner("Extracting text from documents..."):
            # Detect file types
            is_pdf1 = doc1_file.name.endswith('.pdf')
            is_pdf2 = doc2_file.name.endswith('.pdf')
            
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
            with st.spinner("Finding differences..."):
                diff_indices1, diff_indices2, words1, words2, sync_info = find_word_differences_with_sync(text1, text2)
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
                'sync_info': sync_info
            }
            st.session_state.comparison_done = True
    
    # Display results
    if st.session_state.results:
        results = st.session_state.results
        
        st.success("✅ Comparison complete!")
        
        # Display statistics
        sync_info = results['sync_info']
        match_percentage = (sync_info['total_matching'] / max(sync_info['total_words1'], sync_info['total_words2'])) * 100
        
        st.info(f"📊 **Alignment**: {sync_info['total_matching']} matching words out of {max(sync_info['total_words1'], sync_info['total_words2'])} ({match_percentage:.1f}% match)")
        
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
            file_ext1 = 'pdf' if results['is_pdf1'] else 'docx'
            st.download_button(
                label=f"⬇️ Download Doc 1 (Highlighted .{file_ext1})",
                data=results['pdf1_bytes'].getvalue(),
                file_name=f"doc1_highlighted.{file_ext1}",
                mime="application/pdf" if results['is_pdf1'] else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
        with col_dl2:
            file_ext2 = 'pdf' if results['is_pdf2'] else 'docx'
            st.download_button(
                label=f"⬇️ Download Doc 2 (Highlighted .{file_ext2})",
                data=results['pdf2_bytes'].getvalue(),
                file_name=f"doc2_highlighted.{file_ext2}",
                mime="application/pdf" if results['is_pdf2'] else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
        # Display text comparison
        st.markdown("### Text Comparison Preview")
        st.markdown("🟡 **Yellow highlight** = Words that differ between documents")
        
        col_diff1, col_diff2 = st.columns(2)
        
        with col_diff1:
            st.markdown("**Document 1**")
            st.markdown(f'<div class="diff-container">{results["html1"]}</div>', unsafe_allow_html=True)
        
        with col_diff2:
            st.markdown("**Document 2**")
            st.markdown(f'<div class="diff-container">{results["html2"]}</div>', unsafe_allow_html=True)
        
        # Sample differences
        with st.expander("📋 Sample Differences (First 50)"):
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                st.markdown(f"**Different words in Doc 1: {len(results['diff_indices1'])} total**")
                sample_indices1 = sorted(list(results['diff_indices1']))[:50]
                for idx in sample_indices1:
                    if idx < len(results['words1']):
                        st.text(f"Position {idx}: '{results['words1'][idx]}'")
            with col_s2:
                st.markdown(f"**Different words in Doc 2: {len(results['diff_indices2'])} total**")
                sample_indices2 = sorted(list(results['diff_indices2']))[:50]
                for idx in sample_indices2:
                    if idx < len(results['words2']):
                        st.text(f"Position {idx}: '{results['words2'][idx]}'")

else:
    st.info("👆 Please upload both documents to begin comparison")

# Footer
st.markdown("---")
st.markdown("💡 **Precise word-by-word comparison** - Highlights only the words that actually differ between documents")