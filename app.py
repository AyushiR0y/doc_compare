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
    """Extract text from Word document with position tracking"""
    try:
        docx_file.seek(0)
        doc = Document(docx_file)
        
        all_words = []
        word_positions = []  # Track where each word came from
        word_idx = 0
        
        # Extract from paragraphs - RUN BY RUN
        for para_idx, para in enumerate(doc.paragraphs):
            for run_idx, run in enumerate(para.runs):
                run_text = run.text
                if run_text.strip():
                    words = run_text.split()
                    for word in words:
                        all_words.append(word)
                        word_positions.append({
                            'index': word_idx,
                            'type': 'paragraph',
                            'para_idx': para_idx,
                            'run_idx': run_idx,
                            'word': word
                        })
                        word_idx += 1
        
        # Extract from tables - RUN BY RUN
        for table_idx, table in enumerate(doc.tables):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    for para_idx, para in enumerate(cell.paragraphs):
                        for run_idx, run in enumerate(para.runs):
                            run_text = run.text
                            if run_text.strip():
                                words = run_text.split()
                                for word in words:
                                    all_words.append(word)
                                    word_positions.append({
                                        'index': word_idx,
                                        'type': 'table',
                                        'table_idx': table_idx,
                                        'row_idx': row_idx,
                                        'cell_idx': cell_idx,
                                        'para_idx': para_idx,
                                        'run_idx': run_idx,
                                        'word': word
                                    })
                                    word_idx += 1
        
        # Join with spaces to create text
        extracted_text = ' '.join(all_words)
        
        if len(all_words) == 0:
            st.warning("⚠️ No text extracted from Word document.")
        elif len(all_words) < 10:
            st.warning(f"⚠️ Only {len(all_words)} words extracted from Word document.")
        
        return extracted_text, word_positions
    except Exception as e:
        st.error(f"Error reading Word document: {str(e)}")
        return None, None

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
            words_data = page.get_text("words")
            
            for word_info in words_data:
                x0, y0, x1, y1, word_text = word_info[:5]
                
                if word_text.strip():
                    all_words.append({
                        'text': word_text.strip(),
                        'bbox': [x0, y0, x1, y1],
                        'page': page_num
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
    # Remove all common punctuation characters
    word = re.sub(r'[.,;:!?\"\'\-\(\)\[\]\{\}/\\`~@#$%^&*+=<>|]', '', word)
    return word.lower().strip()

def find_word_differences_with_sync(text1, text2):
    """
    Find word-level differences with improved syncing.
    Only marks words as different if normalized content differs.
    """
    # Split into words
    words1 = text1.split()
    words2 = text2.split()
    
    # Create normalized versions for comparison
    normalized1 = [normalize_word(w) for w in words1]
    normalized2 = [normalize_word(w) for w in words2]
    
    # Use SequenceMatcher with better parameters for syncing
    matcher = difflib.SequenceMatcher(
        None, 
        normalized1, 
        normalized2, 
        autojunk=False
    )
    
    # Get matching blocks
    matching_blocks = matcher.get_matching_blocks()
    
    # Filter out very small matches that might be noise
    significant_matches = []
    for match in matching_blocks:
        if match.size >= 3 or match == matching_blocks[-1]:
            significant_matches.append(match)
    
    # If we filtered out too many matches, use original
    if len(significant_matches) < len(matching_blocks) * 0.3:
        significant_matches = matching_blocks
    
    # Rebuild matcher
    matcher.matching_blocks = significant_matches
    
    # Sets to store indices of different words
    diff_indices1 = set()
    diff_indices2 = set()
    
    # Process each operation
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            continue
        elif tag == 'replace':
            # Check each word individually in the replacement
            min_len = min(i2 - i1, j2 - j1)
            
            # Compare words that align
            for idx in range(min_len):
                norm1 = normalized1[i1 + idx]
                norm2 = normalized2[j1 + idx]
                # Only mark as different if normalized words differ
                if norm1 != norm2:
                    diff_indices1.add(i1 + idx)
                    diff_indices2.add(j1 + idx)
            
            # Add any remaining words from the longer sequence
            if i2 - i1 > min_len:
                diff_indices1.update(range(i1 + min_len, i2))
            if j2 - j1 > min_len:
                diff_indices2.update(range(j1 + min_len, j2))
                
        elif tag == 'delete':
            diff_indices1.update(range(i1, i2))
        elif tag == 'insert':
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
        'diff_words2': len(diff_indices2),
        'significant_matches': len(significant_matches)
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

def highlight_word_doc(docx_file, word_positions, diff_indices):
    """
    Highlight words in Word document using position tracking.
    """
    from docx.enum.text import WD_COLOR_INDEX
    
    docx_file.seek(0)
    doc = Document(docx_file)
    
    # Create a set of positions that need highlighting for faster lookup
    highlight_set = set(diff_indices)
    
    # Track which runs need highlighting
    runs_to_highlight = {}  # Key: (type, indices...), Value: True
    
    # Map diff indices to run positions
    for pos in word_positions:
        if pos['index'] in highlight_set:
            if pos['type'] == 'paragraph':
                key = ('para', pos['para_idx'], pos['run_idx'])
            else:  # table
                key = ('table', pos['table_idx'], pos['row_idx'], pos['cell_idx'], pos['para_idx'], pos['run_idx'])
            runs_to_highlight[key] = True
    
    # Apply highlighting to paragraphs
    for para_idx, para in enumerate(doc.paragraphs):
        for run_idx, run in enumerate(para.runs):
            key = ('para', para_idx, run_idx)
            if key in runs_to_highlight:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    
    # Apply highlighting to tables
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for para_idx, para in enumerate(cell.paragraphs):
                    for run_idx, run in enumerate(para.runs):
                        key = ('table', table_idx, row_idx, cell_idx, para_idx, run_idx)
                        if key in runs_to_highlight:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    
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
                word_positions1 = None
            else:
                result = extract_text_from_word(doc1_file)
                if result[0] is None:
                    st.stop()
                text1, word_positions1 = result
                word_data1 = None
                pdf_doc1 = None
            
            if is_pdf2:
                text2, word_data2, pdf_doc2 = extract_text_from_pdf(doc2_file)
                word_positions2 = None
            else:
                result = extract_text_from_word(doc2_file)
                if result[0] is None:
                    st.stop()
                text2, word_positions2 = result
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
                    pdf1_bytes = highlight_word_doc(doc1_file, word_positions1, diff_indices1)
                
                if is_pdf2:
                    highlighted_doc2 = highlight_pdf_words(pdf_doc2, word_data2, diff_indices2)
                    pdf2_bytes = BytesIO()
                    highlighted_doc2.save(pdf2_bytes)
                    pdf2_bytes.seek(0)
                    highlighted_doc2.close()
                    pdf_doc2.close()
                else:
                    pdf2_bytes = highlight_word_doc(doc2_file, word_positions2, diff_indices2)
            
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
                        word1 = results['words1'][idx]
                        word2 = results['words2'][idx] if idx < len(results['words2']) else "N/A"
                        norm1 = normalize_word(word1)
                        norm2 = normalize_word(word2) if idx < len(results['words2']) else "N/A"
                        st.text(f"Pos {idx}: '{word1}' (norm: '{norm1}') vs '{word2}' (norm: '{norm2}')")
            with col_s2:
                st.markdown(f"**Different words in Doc 2: {len(results['diff_indices2'])} total**")
                sample_indices2 = sorted(list(results['diff_indices2']))[:50]
                for idx in sample_indices2:
                    if idx < len(results['words2']):
                        word2 = results['words2'][idx]
                        word1 = results['words1'][idx] if idx < len(results['words1']) else "N/A"
                        norm2 = normalize_word(word2)
                        norm1 = normalize_word(word1) if idx < len(results['words1']) else "N/A"
                        st.text(f"Pos {idx}: '{word2}' (norm: '{norm2}') vs '{word1}' (norm: '{norm1}')")

else:
    st.info("👆 Please upload both documents to begin comparison")

# Footer
st.markdown("---")
st.markdown("💡 **Precise word-by-word comparison** - Highlights only the words that actually differ between documents")