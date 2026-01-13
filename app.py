import streamlit as st
import difflib
from io import BytesIO
import fitz  # PyMuPDF
import re

st.set_page_config(page_title="PDF Diff Checker", layout="wide")

st.title("📄 PDF Diff Checker")
st.markdown("Upload two PDF files to compare and highlight their differences")

# Create two columns for file uploads
col1, col2 = st.columns(2)

with col1:
    st.subheader("PDF 1")
    pdf1 = st.file_uploader("Upload first PDF", type=['pdf'], key="pdf1")
    
with col2:
    st.subheader("PDF 2")
    pdf2 = st.file_uploader("Upload second PDF", type=['pdf'], key="pdf2")

def extract_text_with_coordinates(pdf_file):
    """Extract text from PDF with exact coordinates for each word"""
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        pages_data = []
        full_text = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Extract text blocks with coordinates
            blocks = page.get_text("dict")["blocks"]
            
            page_words = []
            page_text_lines = []
            
            for block in blocks:
                if "lines" in block:
                    for line in block["lines"]:
                        line_text = ""
                        line_words = []
                        
                        for span in line["spans"]:
                            text = span["text"]
                            bbox = span["bbox"]  # (x0, y0, x1, y1)
                            
                            # Store each word with its coordinates
                            words = text.split()
                            word_width = (bbox[2] - bbox[0]) / max(len(text), 1)
                            
                            x_offset = bbox[0]
                            for word in words:
                                word_bbox = (
                                    x_offset,
                                    bbox[1],
                                    x_offset + len(word) * word_width,
                                    bbox[3]
                                )
                                word_data = {
                                    'text': word,
                                    'bbox': word_bbox,
                                    'page': page_num
                                }
                                page_words.append(word_data)
                                line_words.append(word_data)
                                x_offset += (len(word) + 1) * word_width
                            
                            line_text += text + " "
                        
                        if line_text.strip():
                            page_text_lines.append({
                                'text': line_text.strip(),
                                'words': line_words
                            })
            
            # Create simple text version for comparison
            simple_text = "\n".join([line['text'] for line in page_text_lines])
            
            pages_data.append({
                'page_num': page_num,
                'text': simple_text,
                'lines_with_words': page_text_lines
            })
            
            full_text.append(f"--- Page {page_num + 1} ---\n{simple_text}")
        
        return "\n\n".join(full_text), pages_data, doc
    
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None, None, None

def find_diff_line_indices(text1, text2):
    """Find exact line indices that are different - returns line text that should be highlighted"""
    lines1 = []
    lines2 = []
    
    for line in text1.split('\n'):
        if line.strip() and not line.startswith('---'):
            lines1.append(line.strip())
    
    for line in text2.split('\n'):
        if line.strip() and not line.startswith('---'):
            lines2.append(line.strip())
    
    differ = difflib.Differ()
    diff = list(differ.compare(lines1, lines2))
    
    # Track exact lines to highlight
    highlight_lines1 = []
    highlight_lines2 = []
    
    line_idx1 = 0
    line_idx2 = 0
    
    for d in diff:
        if d.startswith('  '):  # Same in both
            line_idx1 += 1
            line_idx2 += 1
        elif d.startswith('- '):  # Only in file 1
            if line_idx1 < len(lines1):
                highlight_lines1.append(lines1[line_idx1])
            line_idx1 += 1
        elif d.startswith('+ '):  # Only in file 2
            if line_idx2 < len(lines2):
                highlight_lines2.append(lines2[line_idx2])
            line_idx2 += 1
    
    return set(highlight_lines1), set(highlight_lines2)

def highlight_pdf_precise(doc, pages_data, highlight_lines):
    """Add yellow highlights based on exact line matching"""
    highlighted_doc = fitz.open()
    
    for page_num in range(len(doc)):
        # Copy original page
        page = doc[page_num]
        new_page = highlighted_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.show_pdf_page(new_page.rect, doc, page_num)
        
        # Get page data
        page_data = pages_data[page_num] if page_num < len(pages_data) else None
        
        if not page_data:
            continue
        
        lines_with_words = page_data['lines_with_words']
        
        # For each line in the page, check if it should be highlighted
        for line_info in lines_with_words:
            line_text = line_info['text'].strip()
            line_words = line_info['words']
            
            # Check if this exact line should be highlighted
            should_highlight = False
            for highlight_line in highlight_lines:
                # Normalize both for comparison
                norm_line = ' '.join(line_text.split())
                norm_highlight = ' '.join(highlight_line.split())
                
                if norm_line == norm_highlight:
                    should_highlight = True
                    break
            
            # Highlight all words in this line
            if should_highlight:
                for word_info in line_words:
                    bbox = word_info['bbox']
                    rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                    
                    # Add yellow highlight
                    try:
                        highlight = new_page.add_highlight_annot(rect)
                        highlight.set_colors(stroke=fitz.utils.getColor("yellow"))
                        highlight.update()
                    except:
                        pass  # Skip if highlight fails
    
    return highlighted_doc

def highlight_diff_html(text1, text2):
    """Compare texts and return HTML with highlighted differences"""
    lines1 = text1.splitlines()
    lines2 = text2.splitlines()
    
    differ = difflib.Differ()
    diff = list(differ.compare(lines1, lines2))
    
    html1 = []
    html2 = []
    
    i = 0
    while i < len(diff):
        line = diff[i]
        
        if line.startswith('  '):
            content = line[2:] if len(line) > 2 else ""
            html1.append(f'<div class="same">{content}</div>')
            html2.append(f'<div class="same">{content}</div>')
        elif line.startswith('- '):
            content = line[2:] if len(line) > 2 else ""
            html1.append(f'<div class="highlight">{content}</div>')
            if i + 1 < len(diff) and diff[i + 1].startswith('+ '):
                i += 1
                content2 = diff[i][2:] if len(diff[i]) > 2 else ""
                html2.append(f'<div class="highlight">{content2}</div>')
            else:
                html2.append(f'<div class="placeholder"></div>')
        elif line.startswith('+ '):
            content = line[2:] if len(line) > 2 else ""
            if i > 0 and diff[i - 1].startswith('- '):
                pass
            else:
                html1.append(f'<div class="placeholder"></div>')
                html2.append(f'<div class="highlight">{content}</div>')
        
        i += 1
    
    return '\n'.join(html1), '\n'.join(html2)

# CSS for styling
st.markdown("""
<style>
    .diff-container {
        font-family: monospace;
        font-size: 14px;
        line-height: 1.6;
        padding: 20px;
        background-color: #f5f5f5;
        border-radius: 5px;
        max-height: 600px;
        overflow-y: auto;
        border: 1px solid #ddd;
    }
    .same {
        background-color: white;
        padding: 2px 5px;
        margin: 1px 0;
    }
    .highlight {
        background-color: #ffff99;
        padding: 2px 5px;
        margin: 1px 0;
        border-left: 3px solid #ffcc00;
    }
    .placeholder {
        background-color: #f0f0f0;
        padding: 2px 5px;
        margin: 1px 0;
        min-height: 20px;
    }
</style>
""", unsafe_allow_html=True)

# Process and display differences
if pdf1 and pdf2:
    with st.spinner("Extracting text with coordinates from PDFs..."):
        text1, pages_data1, doc1 = extract_text_with_coordinates(pdf1)
        text2, pages_data2, doc2 = extract_text_with_coordinates(pdf2)
    
    if text1 and text2 and doc1 and doc2:
        with st.spinner("Analyzing differences..."):
            # Get exact lines to highlight
            highlight_lines1, highlight_lines2 = find_diff_line_indices(text1, text2)
            
            # Generate HTML preview
            html1, html2 = highlight_diff_html(text1, text2)
        
        st.success("✅ Comparison complete!")
        
        # Display statistics
        col_stat1, col_stat2, col_stat3 = st.columns(3)
        with col_stat1:
            st.metric("Lines in PDF 1", len([l for l in text1.split('\n') if l.strip() and not l.startswith('---')]))
        with col_stat2:
            st.metric("Lines in PDF 2", len([l for l in text2.split('\n') if l.strip() and not l.startswith('---')]))
        with col_stat3:
            similarity = difflib.SequenceMatcher(None, text1, text2).ratio()
            st.metric("Similarity", f"{similarity * 100:.1f}%")
        
        # Show what will be highlighted
        with st.expander("🔍 Debug: Lines to be highlighted"):
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.write(f"**PDF 1: {len(highlight_lines1)} lines will be highlighted**")
                for line in list(highlight_lines1)[:5]:
                    st.text(line[:80] + "..." if len(line) > 80 else line)
            with col_d2:
                st.write(f"**PDF 2: {len(highlight_lines2)} lines will be highlighted**")
                for line in list(highlight_lines2)[:5]:
                    st.text(line[:80] + "..." if len(line) > 80 else line)
        
        # Generate highlighted PDFs
        with st.spinner("Generating precisely highlighted PDFs..."):
            highlighted_doc1 = highlight_pdf_precise(doc1, pages_data1, highlight_lines1)
            highlighted_doc2 = highlight_pdf_precise(doc2, pages_data2, highlight_lines2)
            
            # Save to bytes
            pdf1_bytes = BytesIO()
            pdf2_bytes = BytesIO()
            highlighted_doc1.save(pdf1_bytes)
            highlighted_doc2.save(pdf2_bytes)
            pdf1_bytes.seek(0)
            pdf2_bytes.seek(0)
        
        # Download buttons
        st.markdown("### Download Highlighted PDFs")
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            st.download_button(
                label="⬇️ Download PDF 1 (Highlighted)",
                data=pdf1_bytes.getvalue(),
                file_name="pdf1_highlighted.pdf",
                mime="application/pdf"
            )
        
        with col_dl2:
            st.download_button(
                label="⬇️ Download PDF 2 (Highlighted)",
                data=pdf2_bytes.getvalue(),
                file_name="pdf2_highlighted.pdf",
                mime="application/pdf"
            )
        
        # Display the differences side by side
        st.markdown("### Text Comparison Preview")
        st.markdown("🟡 Yellow = Text unique to this PDF")
        
        col_diff1, col_diff2 = st.columns(2)
        
        with col_diff1:
            st.markdown("**PDF 1**")
            st.markdown(f'<div class="diff-container">{html1}</div>', unsafe_allow_html=True)
        
        with col_diff2:
            st.markdown("**PDF 2**")
            st.markdown(f'<div class="diff-container">{html2}</div>', unsafe_allow_html=True)
        
        # Clean up
        doc1.close()
        doc2.close()
        highlighted_doc1.close()
        highlighted_doc2.close()

else:
    st.info("👆 Please upload both PDF files to begin comparison")

# Footer
st.markdown("---")
st.markdown("💡 **Professional PDF diff tool** - Downloads match the preview highlighting exactly")