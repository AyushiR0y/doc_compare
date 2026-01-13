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
    """Extract text from Word document"""
    try:
        docx_file.seek(0)
        doc = Document(docx_file)
        
        paragraphs = []
        for para in doc.paragraphs:
            if para.text.strip():
                paragraphs.append(para.text.strip())
        
        return '\n'.join(paragraphs)
    except Exception as e:
        st.error(f"Error reading Word document: {str(e)}")
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
            blocks = page.get_text("dict")["blocks"]
            
            page_words = []
            page_text_parts = []
            
            for block in blocks:
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            text = span["text"]
                            bbox = span["bbox"]
                            
                            # Split into words and track each
                            words = text.split()
                            if not words:
                                continue
                            
                            char_width = (bbox[2] - bbox[0]) / len(text) if text else 0
                            x_pos = bbox[0]
                            
                            for word in words:
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
                                page_text_parts.append(word)
                                x_pos += (len(word) + 1) * char_width
            
            page_text = ' '.join(page_text_parts)
            pages_data.append({
                'page_num': page_num,
                'text': page_text,
                'words': page_words
            })
            full_text.append(page_text)
        
        return '\n'.join(full_text), pages_data, doc
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None, None, None

def find_word_level_differences(text1, text2):
    """Find word-level differences by comparing line by line"""
    lines1 = text1.split('\n')
    lines2 = text2.split('\n')
    
    # Use difflib to find matching and non-matching lines
    matcher = difflib.SequenceMatcher(None, lines1, lines2)
    
    diff_words1 = set()
    diff_words2 = set()
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            # Lines are identical, skip
            continue
        elif tag == 'replace':
            # Lines differ - find word differences within these lines
            for line1, line2 in zip(lines1[i1:i2], lines2[j1:j2]):
                words1 = re.findall(r'\S+', line1)
                words2 = re.findall(r'\S+', line2)
                
                # Find word-level diffs within this line pair
                word_matcher = difflib.SequenceMatcher(None, words1, words2)
                for wtag, wi1, wi2, wj1, wj2 in word_matcher.get_opcodes():
                    if wtag in ['replace', 'delete']:
                        diff_words1.update(words1[wi1:wi2])
                    if wtag in ['replace', 'insert']:
                        diff_words2.update(words2[wj1:wj2])
        elif tag == 'delete':
            # Lines only in text1
            for line in lines1[i1:i2]:
                diff_words1.update(re.findall(r'\S+', line))
        elif tag == 'insert':
            # Lines only in text2
            for line in lines2[j1:j2]:
                diff_words2.update(re.findall(r'\S+', line))
    
    return diff_words1, diff_words2

def create_html_diff(text1, text2, diff_words1, diff_words2):
    """Create HTML with word-level highlighting"""
    def highlight_text(text, diff_words):
        words = re.findall(r'\S+|\s+', text)
        html_parts = []
        
        for word in words:
            if word.strip() in diff_words:
                html_parts.append(f'<span class="highlight">{word}</span>')
            else:
                html_parts.append(word)
        
        return ''.join(html_parts)
    
    html1 = highlight_text(text1, diff_words1)
    html2 = highlight_text(text2, diff_words2)
    
    return html1, html2

def highlight_pdf_words(doc, pages_data, diff_words):
    """Highlight specific words in PDF"""
    highlighted_doc = fitz.open()
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        new_page = highlighted_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.show_pdf_page(new_page.rect, doc, page_num)
        
        page_data = pages_data[page_num] if page_num < len(pages_data) else None
        if not page_data:
            continue
        
        page_words = page_data['words']
        
        for word_info in page_words:
            word_text = word_info['text']
            
            # Check if this exact word should be highlighted
            if word_text in diff_words:
                bbox = word_info['bbox']
                rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                
                try:
                    highlight = new_page.add_highlight_annot(rect)
                    highlight.set_colors(stroke=fitz.utils.getColor("yellow"))
                    highlight.update()
                except:
                    pass
    
    return highlighted_doc

def highlight_word_doc(docx_file, diff_words):
    """Highlight words in Word document"""
    from docx.shared import RGBColor
    from docx.enum.text import WD_COLOR_INDEX
    
    docx_file.seek(0)
    doc = Document(docx_file)
    
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            # Check if any diff word is in this run
            for diff_word in diff_words:
                if diff_word in run.text:
                    # Highlight the run
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    break
    
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
                diff_words1, diff_words2 = find_word_level_differences(text1, text2)
                html1, html2 = create_html_diff(text1, text2, diff_words1, diff_words2)
            
            with st.spinner("Generating highlighted documents..."):
                # Generate highlighted versions
                if is_pdf1:
                    highlighted_doc1 = highlight_pdf_words(pdf_doc1, pages_data1, diff_words1)
                    pdf1_bytes = BytesIO()
                    highlighted_doc1.save(pdf1_bytes)
                    pdf1_bytes.seek(0)
                    highlighted_doc1.close()
                    pdf_doc1.close()
                else:
                    pdf1_bytes = highlight_word_doc(doc1_file, diff_words1)
                
                if is_pdf2:
                    highlighted_doc2 = highlight_pdf_words(pdf_doc2, pages_data2, diff_words2)
                    pdf2_bytes = BytesIO()
                    highlighted_doc2.save(pdf2_bytes)
                    pdf2_bytes.seek(0)
                    highlighted_doc2.close()
                    pdf_doc2.close()
                else:
                    pdf2_bytes = highlight_word_doc(doc2_file, diff_words2)
            
            # Store results in session state
            st.session_state.results = {
                'text1': text1,
                'text2': text2,
                'diff_words1': diff_words1,
                'diff_words2': diff_words2,
                'html1': html1,
                'html2': html2,
                'pdf1_bytes': pdf1_bytes,
                'pdf2_bytes': pdf2_bytes,
                'is_pdf1': is_pdf1,
                'is_pdf2': is_pdf2
            }
            st.session_state.comparison_done = True
    
    # Display results from session state
    if st.session_state.results:
        results = st.session_state.results
        
        st.success("✅ Comparison complete!")
        
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
        st.info(f"🔍 Found **{len(results['diff_words1'])}** unique words in Doc 1 and **{len(results['diff_words2'])}** unique words in Doc 2")
        
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
        with st.expander("📋 Sample Differences"):
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                st.markdown("**Unique to Doc 1 (sample):**")
                sample1 = list(results['diff_words1'])[:20]
                st.write(", ".join(sample1))
            with col_s2:
                st.markdown("**Unique to Doc 2 (sample):**")
                sample2 = list(results['diff_words2'])[:20]
                st.write(", ".join(sample2))

else:
    st.info("👆 Please upload both documents to begin comparison")

# Footer
st.markdown("---")
st.markdown("💡 **Word-level precision** - Only highlights the actual different words (e.g., 'XXXXX' vs 'Arya')")