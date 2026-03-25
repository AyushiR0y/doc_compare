from __future__ import annotations

import base64
import difflib
from io import BytesIO
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
from docx import Document
from PIL import Image, ImageDraw


ALLOWED_EXTENSIONS = {"pdf", "docx"}


def _normalize_word(word: str) -> str:
    import string

    translator = str.maketrans("", "", string.punctuation)
    return word.translate(translator).lower().strip()


def _extract_words_from_word(file_bytes: bytes) -> tuple[str, list[dict[str, Any]], Document]:
    doc = Document(BytesIO(file_bytes))

    word_objects: list[dict[str, Any]] = []
    text_segments: list[str] = []

    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    def process_paragraph(para: Paragraph) -> None:
        if not para.text.strip():
            return

        is_heading_para = para.style.name.startswith("Heading") if para.style else False
        tokens = para.text.split()

        for token in tokens:
            if not token.strip():
                continue

            is_bold = False
            is_italic = False
            for run in para.runs:
                if token in run.text:
                    if run.bold:
                        is_bold = True
                    if run.italic:
                        is_italic = True
                    break

            word_objects.append(
                {
                    "text": token,
                    "type": "word",
                    "is_bold": is_bold,
                    "is_italic": is_italic,
                    "is_heading": is_heading_para,
                    "in_table": False,
                }
            )
            text_segments.append(token)

        word_objects.append({"type": "newline", "text": "\n"})

    def process_table(table: Table) -> None:
        word_objects.append({"type": "table_start", "text": ""})

        for row in table.rows:
            for cell_idx, cell in enumerate(row.cells):
                cell_text = cell.text.strip()
                if cell_text:
                    for token in cell_text.split():
                        if not token.strip():
                            continue

                        word_objects.append(
                            {
                                "text": token,
                                "type": "word",
                                "is_bold": False,
                                "is_italic": False,
                                "is_heading": False,
                                "in_table": True,
                            }
                        )
                        text_segments.append(token)

                if cell_idx < len(row.cells) - 1:
                    word_objects.append({"type": "separator", "text": "|"})

            word_objects.append({"type": "row_end", "text": "\n"})

        word_objects.append({"type": "table_end", "text": ""})
        word_objects.append({"type": "newline", "text": "\n"})

    for element in doc.element.body:
        if isinstance(element, CT_P):
            process_paragraph(Paragraph(element, doc))
        elif isinstance(element, CT_Tbl):
            process_table(Table(element, doc))

    return " ".join(text_segments), word_objects, doc


def _extract_words_from_pdf(file_bytes: bytes) -> tuple[str, list[dict[str, Any]], list[dict[str, Any]], fitz.Document]:
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    word_objects: list[dict[str, Any]] = []
    highlight_data: list[dict[str, Any]] = []
    text_segments: list[str] = []

    for page_num in range(len(doc)):
        page = doc[page_num]

        table_rects: list[tuple[float, float, float, float]] = []
        try:
            tables = page.find_tables()
            if tables:
                for table in tables:
                    table_rects.append(table.bbox)
        except Exception:
            table_rects = []

        words = page.get_text("words")
        prev_block = -1

        for word_info in words:
            x0, y0, x1, y1, word_text = word_info[:5]
            block_no = word_info[5]

            if block_no != prev_block:
                if prev_block != -1:
                    word_objects.append({"type": "newline", "text": "\n"})
                prev_block = block_no

            if not word_text.strip():
                continue

            in_table = False
            for table_rect in table_rects:
                if x0 >= table_rect[0] and x1 <= table_rect[2] and y0 >= table_rect[1] and y1 <= table_rect[3]:
                    in_table = True
                    break

            word = word_text.strip()
            word_objects.append(
                {
                    "text": word,
                    "type": "word",
                    "is_bold": False,
                    "is_italic": False,
                    "is_heading": False,
                    "in_table": in_table,
                }
            )
            highlight_data.append({"text": word, "bbox": [x0, y0, x1, y1], "page": page_num, "in_table": in_table})
            text_segments.append(word)

        word_objects.append({"type": "newline", "text": "\n"})

    return " ".join(text_segments), word_objects, highlight_data, doc


def _run_diff(text1: str, text2: str) -> tuple[set[int], set[int], dict[str, int]]:
    words1 = text1.split()
    words2 = text2.split()

    norm_words1 = [_normalize_word(w) for w in words1 if _normalize_word(w)]
    norm_words2 = [_normalize_word(w) for w in words2 if _normalize_word(w)]

    matcher = difflib.SequenceMatcher(None, norm_words1, norm_words2, autojunk=False)
    opcodes = matcher.get_opcodes()

    diff_indices1: set[int] = set()
    diff_indices2: set[int] = set()

    orig_idx1 = 0
    norm_idx1 = 0
    for w in words1:
        if _normalize_word(w):
            for tag, i1, i2, _, _ in opcodes:
                if tag in {"replace", "delete"} and i1 <= norm_idx1 < i2:
                    diff_indices1.add(orig_idx1)
            norm_idx1 += 1
        orig_idx1 += 1

    orig_idx2 = 0
    norm_idx2 = 0
    for w in words2:
        if _normalize_word(w):
            for tag, _, _, j1, j2 in opcodes:
                if tag in {"replace", "insert"} and j1 <= norm_idx2 < j2:
                    diff_indices2.add(orig_idx2)
            norm_idx2 += 1
        orig_idx2 += 1

    total_matching = sum(i2 - i1 for tag, i1, i2, _, _ in opcodes if tag == "equal")
    info = {
        "total_matching": total_matching,
        "total_words1": len(words1),
        "total_words2": len(words2),
        "diff_words1": len(diff_indices1),
        "diff_words2": len(diff_indices2),
    }

    return diff_indices1, diff_indices2, info


def _highlight_pdf_words(doc: fitz.Document, word_data: list[dict[str, Any]], diff_indices: set[int]) -> BytesIO:
    highlighted_doc = fitz.open()
    try:
        for page_num in range(len(doc)):
            page = doc[page_num]
            new_page = highlighted_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.show_pdf_page(new_page.rect, doc, page_num)

            for word_idx in diff_indices:
                if word_idx >= len(word_data):
                    continue

                word_info = word_data[word_idx]
                if word_info["page"] != page_num:
                    continue

                bbox = word_info["bbox"]
                rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                try:
                    highlight = new_page.add_highlight_annot(rect)
                    if word_info.get("in_table", False):
                        highlight.set_colors(stroke=[1.0, 0.93, 0.88])
                    else:
                        highlight.set_colors(stroke=fitz.utils.getColor("yellow"))
                    highlight.update()
                except Exception:
                    continue

        output = BytesIO()
        highlighted_doc.save(output)
        output.seek(0)
        return output
    finally:
        highlighted_doc.close()


def _highlight_word_doc(doc: Document, word_objects: list[dict[str, Any]], diff_indices: set[int]) -> BytesIO:
    from docx.enum.text import WD_COLOR_INDEX
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    text_idx = 0
    obj_indices_to_highlight: set[int] = set()
    for obj_idx, obj in enumerate(word_objects):
        if obj["type"] == "word":
            if text_idx in diff_indices:
                obj_indices_to_highlight.add(obj_idx)
            text_idx += 1

    current_obj_idx = 0

    for element in doc.element.body:
        if isinstance(element, CT_P):
            para = Paragraph(element, doc)
            if not para.text.strip():
                continue

            for run in para.runs:
                if not run.text:
                    continue

                highlight_run = False
                for _ in run.text.split():
                    while current_obj_idx < len(word_objects) and word_objects[current_obj_idx]["type"] != "word":
                        current_obj_idx += 1

                    if current_obj_idx < len(word_objects) and current_obj_idx in obj_indices_to_highlight:
                        highlight_run = True
                    current_obj_idx += 1

                if highlight_run:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

        elif isinstance(element, CT_Tbl):
            table = Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if not run.text:
                                continue

                            highlight_run = False
                            for _ in run.text.split():
                                while current_obj_idx < len(word_objects) and word_objects[current_obj_idx]["type"] != "word":
                                    current_obj_idx += 1

                                if current_obj_idx < len(word_objects) and current_obj_idx in obj_indices_to_highlight:
                                    highlight_run = True
                                current_obj_idx += 1

                            if highlight_run:
                                run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def _create_html_preview(word_objects: list[dict[str, Any]], diff_indices: set[int]) -> str:
    html_parts: list[str] = []
    html_parts.append('<div style="padding: 20px; line-height: 1.8;">')

    text_idx = 0
    in_table = False

    for obj in word_objects:
        if obj["type"] == "table_start":
            html_parts.append('<div style="margin: 15px 0; overflow-x: auto;">')
            html_parts.append('<table style="width: 100%; border-collapse: collapse; border: 1px solid #ddd;"><tr>')
            in_table = True
            html_parts.append('<td style="padding: 8px; border: 1px solid #ddd; vertical-align: top;">')

        elif obj["type"] == "table_end":
            html_parts.append("</td></tr></table></div>")
            in_table = False

        elif obj["type"] == "word":
            import html

            escaped_text = html.escape(obj["text"])
            token_html = escaped_text
            if text_idx in diff_indices:
                token_html = f'<span style="background:#fff3a3; padding:2px 4px; border-radius:2px; font-weight:700;">{escaped_text}</span>'

            if obj.get("is_heading"):
                token_html = f'<span style="font-size:1.2em; font-weight:700; color:#2c3e50;">{token_html}</span>'
            if obj.get("is_bold"):
                token_html = f"<strong>{token_html}</strong>"
            if obj.get("is_italic"):
                token_html = f"<em>{token_html}</em>"

            html_parts.append(token_html)
            html_parts.append(" ")
            text_idx += 1

        elif obj["type"] == "row_end":
            if in_table:
                html_parts.append('</td></tr><tr><td style="padding: 8px; border: 1px solid #ddd; vertical-align: top;">')

        elif obj["type"] == "separator":
            if in_table:
                html_parts.append('</td><td style="padding: 8px; border: 1px solid #ddd; vertical-align: top;">')

        elif obj["type"] == "newline":
            if not in_table:
                html_parts.append("<br><br>")

    html_parts.append("</div>")
    return "".join(html_parts)


def _render_pdf_preview_base64(
    doc: fitz.Document,
    word_data: list[dict[str, Any]],
    diff_indices: set[int],
    max_pages: int,
    include_images: bool,
) -> list[dict[str, Any]]:
    previews: list[dict[str, Any]] = []
    total_pages = len(doc) if max_pages <= 0 else min(max_pages, len(doc))

    for page_num in range(total_pages):
        page = doc[page_num]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        draw = ImageDraw.Draw(image, "RGBA")

        for idx in diff_indices:
            if idx >= len(word_data):
                continue
            word_info = word_data[idx]
            if word_info["page"] != page_num:
                continue

            x0, y0, x1, y1 = word_info["bbox"]
            x0, y0, x1, y1 = x0 * 2, y0 * 2, x1 * 2, y1 * 2
            fill_color = (255, 237, 224, 110) if word_info.get("in_table", False) else (255, 255, 0, 110)
            draw.rectangle([x0, y0, x1, y1], fill=fill_color, outline=(255, 200, 0, 200), width=2)

        page_payload: dict[str, Any] = {"page": page_num + 1}
        if include_images:
            output = BytesIO()
            image.save(output, format="JPEG", quality=55, optimize=True)
            encoded = base64.b64encode(output.getvalue()).decode("utf-8")
            page_payload["image_base64"] = encoded
            page_payload["mime_type"] = "image/jpeg"

        previews.append(page_payload)

    return previews


def _safe_name(name: str) -> str:
    return Path(name).name


def _extension(filename: str) -> str:
    return Path(filename).suffix.lower().lstrip(".")


def compare_documents(doc1_name: str, doc1_bytes: bytes, doc2_name: str, doc2_bytes: bytes) -> dict[str, Any]:
    ext1 = _extension(doc1_name)
    ext2 = _extension(doc2_name)

    if ext1 not in ALLOWED_EXTENSIONS or ext2 not in ALLOWED_EXTENSIONS:
        raise ValueError("Only .pdf and .docx files are supported")

    is_pdf1 = ext1 == "pdf"
    is_pdf2 = ext2 == "pdf"

    pdf_doc1 = None
    pdf_doc2 = None

    if is_pdf1:
        text1, word_objs1, highlight_data1, pdf_doc1 = _extract_words_from_pdf(doc1_bytes)
        docx_doc1 = None
    else:
        text1, word_objs1, docx_doc1 = _extract_words_from_word(doc1_bytes)
        highlight_data1 = None

    if is_pdf2:
        text2, word_objs2, highlight_data2, pdf_doc2 = _extract_words_from_pdf(doc2_bytes)
        docx_doc2 = None
    else:
        text2, word_objs2, docx_doc2 = _extract_words_from_word(doc2_bytes)
        highlight_data2 = None

    if not text1 or not text2:
        if pdf_doc1:
            pdf_doc1.close()
        if pdf_doc2:
            pdf_doc2.close()
        raise ValueError("Could not extract text from one or both documents")

    diff1, diff2, info = _run_diff(text1, text2)

    try:
        if is_pdf1:
            highlighted_doc1 = _highlight_pdf_words(pdf_doc1, highlight_data1, diff1)
        else:
            highlighted_doc1 = _highlight_word_doc(docx_doc1, word_objs1, diff1)

        if is_pdf2:
            highlighted_doc2 = _highlight_pdf_words(pdf_doc2, highlight_data2, diff2)
        else:
            highlighted_doc2 = _highlight_word_doc(docx_doc2, word_objs2, diff2)
    finally:
        if pdf_doc1:
            pdf_doc1.close()
        if pdf_doc2:
            pdf_doc2.close()

    original_name1 = Path(_safe_name(doc1_name)).stem
    original_name2 = Path(_safe_name(doc2_name)).stem

    return {
        "doc1_output_name": f"highlighted_{original_name1}.{ext1}",
        "doc2_output_name": f"highlighted_{original_name2}.{ext2}",
        "doc1_bytes": highlighted_doc1.getvalue(),
        "doc2_bytes": highlighted_doc2.getvalue(),
        "doc1_ext": ext1,
        "doc2_ext": ext2,
        "summary": {
            **info,
            "match_rate": round((info["total_matching"] / max(info["total_words1"], info["total_words2"])) * 100, 2)
            if max(info["total_words1"], info["total_words2"]) > 0
            else 0,
        },
    }


def compare_documents_with_preview(
    doc1_name: str,
    doc1_bytes: bytes,
    doc2_name: str,
    doc2_bytes: bytes,
    max_pages: int = 0,
    include_images: bool = False,
) -> dict[str, Any]:
    ext1 = _extension(doc1_name)
    ext2 = _extension(doc2_name)

    if ext1 not in ALLOWED_EXTENSIONS or ext2 not in ALLOWED_EXTENSIONS:
        raise ValueError("Only .pdf and .docx files are supported")

    is_pdf1 = ext1 == "pdf"
    is_pdf2 = ext2 == "pdf"

    pdf_doc1 = None
    pdf_doc2 = None

    if is_pdf1:
        text1, word_objs1, highlight_data1, pdf_doc1 = _extract_words_from_pdf(doc1_bytes)
        docx_doc1 = None
    else:
        text1, word_objs1, docx_doc1 = _extract_words_from_word(doc1_bytes)
        highlight_data1 = None

    if is_pdf2:
        text2, word_objs2, highlight_data2, pdf_doc2 = _extract_words_from_pdf(doc2_bytes)
        docx_doc2 = None
    else:
        text2, word_objs2, docx_doc2 = _extract_words_from_word(doc2_bytes)
        highlight_data2 = None

    if not text1 or not text2:
        if pdf_doc1:
            pdf_doc1.close()
        if pdf_doc2:
            pdf_doc2.close()
        raise ValueError("Could not extract text from one or both documents")

    diff1, diff2, info = _run_diff(text1, text2)

    try:
        if is_pdf1:
            doc1_page_limit = len(pdf_doc1) if max_pages <= 0 else min(max_pages, len(pdf_doc1))
            preview1 = {
                "type": "pdf_images",
                "pages": _render_pdf_preview_base64(
                    pdf_doc1,
                    highlight_data1,
                    diff1,
                    max_pages=max_pages,
                    include_images=include_images,
                ),
                "images_included": include_images,
                "truncated": len(pdf_doc1) > doc1_page_limit,
                "total_pages": len(pdf_doc1),
            }
        else:
            preview1 = {
                "type": "html",
                "html": _create_html_preview(word_objs1, diff1),
                "truncated": False,
                "total_pages": None,
            }

        if is_pdf2:
            doc2_page_limit = len(pdf_doc2) if max_pages <= 0 else min(max_pages, len(pdf_doc2))
            preview2 = {
                "type": "pdf_images",
                "pages": _render_pdf_preview_base64(
                    pdf_doc2,
                    highlight_data2,
                    diff2,
                    max_pages=max_pages,
                    include_images=include_images,
                ),
                "images_included": include_images,
                "truncated": len(pdf_doc2) > doc2_page_limit,
                "total_pages": len(pdf_doc2),
            }
        else:
            preview2 = {
                "type": "html",
                "html": _create_html_preview(word_objs2, diff2),
                "truncated": False,
                "total_pages": None,
            }
    finally:
        if pdf_doc1:
            pdf_doc1.close()
        if pdf_doc2:
            pdf_doc2.close()

    return {
        "summary": {
            **info,
            "match_rate": round((info["total_matching"] / max(info["total_words1"], info["total_words2"])) * 100, 2)
            if max(info["total_words1"], info["total_words2"]) > 0
            else 0,
        },
        "doc1": {"name": Path(_safe_name(doc1_name)).name, "extension": ext1, "preview": preview1},
        "doc2": {"name": Path(_safe_name(doc2_name)).name, "extension": ext2, "preview": preview2},
    }
