import streamlit as st
import difflib
from io import BytesIO
import fitz  # PyMuPDF
from docx import Document
import re

st.set_page_config(page_title="Document Diff Checker", layout="wide")

st.title("📄 Document Diff Checker")
st.markdown("Upload two documents (PDF or Word) to compare and highlight their differences")

# Advanced settings
with st.expander("⚙️ Advanced Settings"):
    st.markdown("**Matching Sensitivity** - Adjust how strictly words must match")
    context_threshold = st.slider(
        "Context similarity threshold (lower = more lenient for layout differences)",
        min_value=0.3,
        max_value=0.9,
        value=0.6,
        step=0.1,
        help="When comparing documents with different layouts (1-col vs 2-col), lower values will be more forgiving"
    )
    context_window = st.slider(
        "Context window size (words before/after to consider)",
        min_value=3,
        max_value=20,
        value=8,
        step=1,
        help="Larger windows provide more context but may be slower"
    )
    st.session_state.context_threshold = context_threshold
    st.session_state.context_window = context_window

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
        
        all_text = []
        seen_text = set()
        
        # Method 1: Extract text from main document body (paragraphs and tables in order)
        for element in doc.element.body:
            # Check if it's a paragraph
            if element.tag.endswith('p'):
                for para in doc.paragraphs:
                    if para._element == element:
                        text = para.text.strip()
                        if text and text not in seen_text:
                            all_text.append(text)
                            seen_text.add(text)
                        break
            # Check if it's a table
            elif element.tag.endswith('tbl'):
                for table in doc.tables:
                    if table._element == element:
                        for row in table.rows:
                            row_text = []
                            for cell in row.cells:
                                cell_text = cell.text.strip()
                                if cell_text:
                                    row_text.append(cell_text)
                            if row_text:
                                combined_row = ' '.join(row_text)
                                if combined_row not in seen_text:
                                    all_text.append(combined_row)
                                    seen_text.add(combined_row)
                        break
        
        # If the above didn't work (fallback to simple extraction)
        if len(all_text) == 0:
            # Simple paragraph extraction
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    all_text.append(text)
            
            # Table extraction
            for table in doc.tables:
                for row in table.rows:
                    row_text = []
                    for cell in row.cells:
                        cell_text = cell.text.strip()
                        if cell_text:
                            row_text.append(cell_text)
                    if row_text:
                        all_text.append(' '.join(row_text))
        
        extracted_text = '\n'.join(all_text)
        
        # Debug: show what was extracted
        word_count = len(re.findall(r'\S+', extracted_text))
        if len(all_text) == 0:
            st.warning("⚠️ No text extracted from Word document. The document might be empty or use unsupported formatting.")
        elif word_count < 10:
            st.warning(f"⚠️ Only {word_count} words extracted from Word document. Check the debug section to verify extraction.")
        
        return extracted_text
    except Exception as e:
        st.error(f"Error reading Word document: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None

def extract_text_from_pdf(pdf_file):
    """Extract text from PDF with exact word coordinates"""
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        pages_data = []
        full_text = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Method 1: Simple text extraction for comparison (prevents duplicates)
            simple_text = page.get_text("text")
            
            # Method 2: Detailed extraction for highlighting
            blocks = page.get_text("dict")["blocks"]
            
            page_words = []
            
            for block in blocks:
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            text = span["text"].strip()
                            bbox = span["bbox"]
                            
                            if not text:
                                continue
                            
                            # Split into words and track each
                            words = text.split()
                            
                            char_width = (bbox[2] - bbox[0]) / len(text) if len(text) > 0 else 0
                            x_pos = bbox[0]
                            
                            for word in words:
                                if word.strip():  # Only add non-empty words
                                    word_bbox = [
                                        x_pos,
                                        bbox[1],
                                        x_pos + len(word) * char_width,
                                        bbox[3]
                                    ]
                                    page_words.append({
                                        'text': word,
                                        'bbox': word_bbox,
                                        'page': page_num
                                    })
                                    x_pos += (len(word) + 1) * char_width
            
            # Use simple text extraction for comparison to avoid duplicates
            pages_data.append({
                'page_num': page_num,
                'text': simple_text,
                'words': page_words
            })
            full_text.append(simple_text)
        
        # Join pages with double newline to preserve page breaks
        return '\n\n'.join(full_text), pages_data, doc
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None, None, None

def find_word_level_differences(text1, text2, context_threshold=0.6, context_window=8):
    """Find actual content differences using sentence-aware alignment"""
    
    # Split into sentences first, then words
    import re
    
    # Simple sentence splitter
    def split_sentences(text):
        # Split on sentence endings but keep the text
        sentences = re.split(r'([.!?]+\s+|\n+)', text)
        result = []
        current = ""
        for part in sentences:
            current += part
            if re.match(r'[.!?]+\s+|\n+', part) or part == sentences[-1]:
                if current.strip():
                    result.append(current.strip())
                current = ""
        return result if result else [text]
    
    sentences1 = split_sentences(text1)
    sentences2 = split_sentences(text2)
    
    # Use SequenceMatcher on sentences to find which sentences differ
    matcher = difflib.SequenceMatcher(None, sentences1, sentences2, autojunk=False)
    
    # Track which words come from different sentences
    diff_indices1 = set()
    diff_indices2 = set()
    
    # Convert to word lists for position tracking
    words1 = re.findall(r'\S+', text1)
    words2 = re.findall(r'\S+', text2)
    
    # Build sentence-to-word-index mapping
    word_to_sentence1 = {}
    word_to_sentence2 = {}
    
    word_idx1 = 0
    for sent_idx, sentence in enumerate(sentences1):
        sent_words = re.findall(r'\S+', sentence)
        for _ in sent_words:
            if word_idx1 < len(words1):
                word_to_sentence1[word_idx1] = sent_idx
                word_idx1 += 1
    
    word_idx2 = 0
    for sent_idx, sentence in enumerate(sentences2):
        sent_words = re.findall(r'\S+', sentence)
        for _ in sent_words:
            if word_idx2 < len(words2):
                word_to_sentence2[word_idx2] = sent_idx
                word_idx2 += 1
    
    # Find different sentences
    diff_sentences1 = set()
    diff_sentences2 = set()
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            # Sentences match, but let's check if they're actually identical
            for sent_idx in range(i1, i2):
                if sent_idx < len(sentences1):
                    s1 = sentences1[sent_idx].lower().strip()
                    corresponding_idx = j1 + (sent_idx - i1)
                    if corresponding_idx < len(sentences2):
                        s2 = sentences2[corresponding_idx].lower().strip()
                        # Do word-level diff within matching sentences
                        if s1 != s2:
                            words_s1 = re.findall(r'\S+', s1)
                            words_s2 = re.findall(r'\S+', s2)
                            word_matcher = difflib.SequenceMatcher(None, words_s1, words_s2)
                            
                            # Find word positions in original text
                            sent_word_idx1 = sum(len(re.findall(r'\S+', sentences1[k])) for k in range(sent_idx))
                            sent_word_idx2 = sum(len(re.findall(r'\S+', sentences2[k])) for k in range(corresponding_idx))
                            
                            for wtag, wi1, wi2, wj1, wj2 in word_matcher.get_opcodes():
                                if wtag != 'equal':
                                    for wi in range(wi1, wi2):
                                        if sent_word_idx1 + wi < len(words1):
                                            diff_indices1.add(sent_word_idx1 + wi)
                                    for wj in range(wj1, wj2):
                                        if sent_word_idx2 + wj < len(words2):
                                            diff_indices2.add(sent_word_idx2 + wj)
        
        elif tag == 'replace':
            # Sentences are different
            diff_sentences1.update(range(i1, i2))
            diff_sentences2.update(range(j1, j2))
        
        elif tag == 'delete':
            # Sentences only in doc1
            diff_sentences1.update(range(i1, i2))
        
        elif tag == 'insert':
            # Sentences only in doc2
            diff_sentences2.update(range(j1, j2))
    
    # Mark all words in different sentences
    for word_idx, sent_idx in word_to_sentence1.items():
        if sent_idx in diff_sentences1:
            diff_indices1.add(word_idx)
    
    for word_idx, sent_idx in word_to_sentence2.items():
        if sent_idx in diff_sentences2:
            diff_indices2.add(word_idx)
    
    # Calculate statistics
    total_sentences1 = len(sentences1)
    total_sentences2 = len(sentences2)
    matching_sentences = total_sentences1 - len(diff_sentences1)
    
    sync_info = {
        'sync_found': matching_sentences > 0,
        'sync_word1': None,
        'sync_idx1': None,
        'sync_idx2': None,
        'words_before_sync1': 0,
        'words_before_sync2': 0,
        'total_matching': matching_sentences,
        'total_words1': len(words1),
        'total_words2': len(words2),
        'match_rate1': (matching_sentences / total_sentences1 * 100) if total_sentences1 > 0 else 0,
        'match_rate2': (matching_sentences / total_sentences2 * 100) if total_sentences2 > 0 else 0,
        'unique_to_1': len(diff_sentences1),
        'unique_to_2': len(diff_sentences2),
        'total_sentences1': total_sentences1,
        'total_sentences2': total_sentences2
    }
    
    return diff_indices1, diff_indices2, words1, words2, sync_info

def create_html_diff(text1, text2, diff_indices1, diff_indices2, words1, words2):
    """Create HTML with word-level highlighting based on positions"""
    def highlight_text(text, diff_indices, words):
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
    
    html1 = highlight_text(text1, diff_indices1, words1)
    html2 = highlight_text(text2, diff_indices2, words2)
    
    return html1, html2

def highlight_pdf_words(doc, pages_data, diff_indices, words_list):
    """Highlight specific word positions in PDF"""
    highlighted_doc = fitz.open()
    
    # Build a mapping of word index to page and bbox
    word_positions = []
    for page_data in pages_data:
        for word_info in page_data['words']:
            word_positions.append({
                'text': word_info['text'],
                'bbox': word_info['bbox'],
                'page': word_info['page']
            })
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        new_page = highlighted_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.show_pdf_page(new_page.rect, doc, page_num)
        
        # Highlight words at the specified indices on this page
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

def highlight_word_doc(docx_file, diff_indices, words_list):
    """Highlight specific word positions in Word document"""
    from docx.shared import RGBColor
    from docx.enum.text import WD_COLOR_INDEX
    
    docx_file.seek(0)
    doc = Document(docx_file)
    
    # Extract all runs with their word positions
    run_word_map = []
    current_word_idx = 0
    
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run_text = run.text
            run_words = re.findall(r'\S+', run_text)
            
            # Map this run to its word indices
            start_idx = current_word_idx
            end_idx = current_word_idx + len(run_words)
            
            run_word_map.append({
                'run': run,
                'start_idx': start_idx,
                'end_idx': end_idx,
                'should_highlight': any(i in diff_indices for i in range(start_idx, end_idx))
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
</style>
""", unsafe_allow_html=True)

# Initialize session state for caching results
if 'comparison_done' not in st.session_state:
    st.session_state.comparison_done = False
if 'results' not in st.session_state:
    st.session_state.results = None

# Process and display differences
if doc1_file and doc2_file:
    # Check if we need to reprocess (files changed)
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
                text1, pages_data1, pdf_doc1 = extract_text_from_pdf(doc1_file)
            else:
                text1 = extract_text_from_word(doc1_file)
                pages_data1 = None
                pdf_doc1 = None
            
            if is_pdf2:
                text2, pages_data2, pdf_doc2 = extract_text_from_pdf(doc2_file)
            else:
                text2 = extract_text_from_word(doc2_file)
                pages_data2 = None
                pdf_doc2 = None
        
        if text1 and text2:
            with st.spinner("Finding word-level differences..."):
                # Get user settings or use defaults
                threshold = st.session_state.get('context_threshold', 0.6)
                window = st.session_state.get('context_window', 8)
                diff_indices1, diff_indices2, words1, words2, sync_info = find_word_level_differences(
                    text1, text2, context_threshold=threshold, context_window=window
                )
                html1, html2 = create_html_diff(text1, text2, diff_indices1, diff_indices2, words1, words2)
            
            with st.spinner("Generating highlighted documents..."):
                # Generate highlighted versions
                if is_pdf1:
                    highlighted_doc1 = highlight_pdf_words(pdf_doc1, pages_data1, diff_indices1, words1)
                    pdf1_bytes = BytesIO()
                    highlighted_doc1.save(pdf1_bytes)
                    pdf1_bytes.seek(0)
                    highlighted_doc1.close()
                    pdf_doc1.close()
                else:
                    pdf1_bytes = highlight_word_doc(doc1_file, diff_indices1, words1)
                
                if is_pdf2:
                    highlighted_doc2 = highlight_pdf_words(pdf_doc2, pages_data2, diff_indices2, words2)
                    pdf2_bytes = BytesIO()
                    highlighted_doc2.save(pdf2_bytes)
                    pdf2_bytes.seek(0)
                    highlighted_doc2.close()
                    pdf_doc2.close()
                else:
                    pdf2_bytes = highlight_word_doc(doc2_file, diff_indices2, words2)
            
            # Store results in session state
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
    
    # Display results from session state
    if st.session_state.results:
        results = st.session_state.results
        
        st.success("✅ Comparison complete!")
        
        # Display sync information
        sync_info = results.get('sync_info', {})
        if sync_info.get('sync_found'):
            st.info(f"📊 **Sentence Analysis**: {sync_info.get('total_sentences1', 0)} sentences in Doc 1, "
                   f"{sync_info.get('total_sentences2', 0)} sentences in Doc 2")
            col_info1, col_info2 = st.columns(2)
            with col_info1:
                st.metric("Different/Added sentences in Doc 1", sync_info.get('unique_to_1', 0))
            with col_info2:
                st.metric("Different/Added sentences in Doc 2", sync_info.get('unique_to_2', 0))
            
            match_rate = min(sync_info.get('match_rate1', 0), sync_info.get('match_rate2', 0))
            if match_rate < 70:
                st.warning(f"⚠️ Only {match_rate:.0f}% sentence match. Documents have significant differences.")
        else:
            st.warning("⚠️ No matching sentences found - documents appear completely different")
        
        # Display statistics
        col_stat1, col_stat2, col_stat3 = st.columns(3)
        with col_stat1:
            st.metric("Words in Doc 1", len(re.findall(r'\S+', results['text1'])))
        with col_stat2:
            st.metric("Words in Doc 2", len(re.findall(r'\S+', results['text2'])))
        with col_stat3:
            similarity = difflib.SequenceMatcher(None, results['text1'], results['text2']).ratio()
            st.metric("Similarity", f"{similarity * 100:.1f}%")
        
        # Show statistics
        st.info(f"🔍 Found **{len(results['diff_indices1'])}** different word positions in Doc 1 and **{len(results['diff_indices2'])}** different word positions in Doc 2")
        
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
        
        # Display the differences side by side
        st.markdown("### Text Comparison Preview")
        st.markdown("🟡 **Yellow highlight** = Words unique to this document")
        
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
                st.markdown(f"**Different words in Doc 1: {len(results['diff_indices1'])} positions**")
                # Get actual words at these positions
                sample_indices1 = sorted(list(results['diff_indices1']))[:50]
                for idx in sample_indices1:
                    if idx < len(results['words1']):
                        st.text(f"Position {idx}: '{results['words1'][idx]}'")
            with col_s2:
                st.markdown(f"**Different words in Doc 2: {len(results['diff_indices2'])} positions**")
                # Get actual words at these positions
                sample_indices2 = sorted(list(results['diff_indices2']))[:50]
                for idx in sample_indices2:
                    if idx < len(results['words2']):
                        st.text(f"Position {idx}: '{results['words2'][idx]}'")
        
        # Debug: Show first few lines of each document
        with st.expander("🔍 Debug: Document Analysis"):
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.markdown("**Document 1 Analysis:**")
                words1 = re.findall(r'\S+', results['text1'])
                st.text(f"Total words extracted: {len(words1)}")
                st.text(f"Total characters: {len(results['text1'])}")
                st.markdown("**First 20 words:**")
                st.text(' '.join(words1[:20]))
                st.markdown("**First 10 lines:**")
                lines1 = results['text1'].split('\n')[:10]
                for i, line in enumerate(lines1):
                    st.text(f"{i+1}: {line[:100]}")
            with col_d2:
                st.markdown("**Document 2 Analysis:**")
                words2 = re.findall(r'\S+', results['text2'])
                st.text(f"Total words extracted: {len(words2)}")
                st.text(f"Total characters: {len(results['text2'])}")
                st.markdown("**First 20 words:**")
                st.text(' '.join(words2[:20]))
                st.markdown("**First 10 lines:**")
                lines2 = results['text2'].split('\n')[:10]
                for i, line in enumerate(lines2):
                    st.text(f"{i+1}: {line[:100]}")

else:
    st.info("👆 Please upload both documents to begin comparison")

# Footer
st.markdown("---")
st.markdown("💡 **Word-level precision** - Only highlights the actual different words (e.g., 'XXXXX' vs 'Arya')")