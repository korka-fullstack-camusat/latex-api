import io
import re
import json
import zipfile
from typing import Optional

import anthropic
from docx import Document
from docx.oxml.ns import qn
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

app = FastAPI(
    title="Word to LaTeX Converter",
    description="Upload a Word document (.docx) and get a compilable LaTeX file + references.bib",
)

client = anthropic.Anthropic()


# ---------------------------------------------------------------------------
# Word document extraction
# ---------------------------------------------------------------------------

def extract_paragraph_info(para) -> Optional[dict]:
    """Return structured info for a paragraph, or None if empty."""
    style_name = para.style.name if para.style else "Normal"

    # Map style name → paragraph type
    para_type = "paragraph"
    level = 0

    if style_name.startswith("Heading"):
        para_type = "heading"
        try:
            level = int(style_name.split()[-1])
        except (ValueError, IndexError):
            level = 1
    elif style_name == "Title":
        para_type = "title"
    elif style_name == "Subtitle":
        para_type = "subtitle"
    elif style_name in ("Quote", "Block Text", "Intense Quote"):
        para_type = "quote"
    elif style_name == "Caption":
        para_type = "caption"
    elif "List" in style_name:
        para_type = "list_item"

    # Detect numbered / bulleted lists via XML numPr element
    num_pr = para._element.find(qn("w:numPr"))
    if num_pr is not None:
        para_type = "list_item"
        ilvl_el = num_pr.find(qn("w:ilvl"))
        if ilvl_el is not None:
            level = int(ilvl_el.get(qn("w:val"), 0))

    # Extract formatted runs
    runs = []
    for run in para.runs:
        if run.text:
            runs.append(
                {
                    "text": run.text,
                    "bold": bool(run.bold),
                    "italic": bool(run.italic),
                    "underline": bool(run.underline),
                    "superscript": bool(run.font.superscript) if run.font.superscript else False,
                    "subscript": bool(run.font.subscript) if run.font.subscript else False,
                }
            )

    if not runs and not para.text.strip():
        return None

    return {
        "type": para_type,
        "style": style_name,
        "level": level,
        "text": para.text,
        "runs": runs,
        "alignment": str(para.alignment) if para.alignment else None,
    }


def extract_table_info(table) -> dict:
    """Return structured info for a table."""
    rows = []
    for row in table.rows:
        row_data = []
        for cell in row.cells:
            cell_text = " ".join(p.text for p in cell.paragraphs).strip()
            row_data.append(cell_text)
        rows.append(row_data)

    return {
        "type": "table",
        "rows": rows,
        "num_cols": len(rows[0]) if rows else 0,
        "num_rows": len(rows),
    }


def extract_docx_content(file_bytes: bytes) -> dict:
    """
    Extract structured content from a .docx file while preserving
    the original order of paragraphs and tables.
    """
    doc = Document(io.BytesIO(file_bytes))

    content: dict = {"metadata": {}, "elements": []}

    # Document metadata
    try:
        props = doc.core_properties
        content["metadata"] = {
            "title": props.title or "",
            "author": props.author or "",
            "subject": props.subject or "",
        }
    except Exception:
        pass

    # Build element → object maps so we can look up by XML element
    para_map = {p._element: p for p in doc.paragraphs}
    table_map = {t._element: t for t in doc.tables}

    for child in doc.element.body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p" and child in para_map:
            info = extract_paragraph_info(para_map[child])
            if info:
                content["elements"].append(info)

        elif tag == "tbl" and child in table_map:
            content["elements"].append(extract_table_info(table_map[child]))

    return content


# ---------------------------------------------------------------------------
# LaTeX conversion via Claude
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """\
You are an expert LaTeX document converter.
Your job is to convert structured Word document content (provided as JSON) into
a complete, compilable LaTeX document that reproduces the original as faithfully
as possible.

Rules:
1. Choose an appropriate document class (article, report, book) based on content.
2. Include all necessary packages: inputenc (utf8), fontenc (T1), geometry,
   hyperref, booktabs (for tables), graphicx, amsmath, etc.
3. Preserve ALL text formatting: bold → \\textbf{}, italic → \\textit{},
   underline → \\underline{}, superscript → \\textsuperscript{},
   subscript → \\textsubscript{}.
4. Map heading styles: Heading 1 → \\section, Heading 2 → \\subsection,
   Heading 3 → \\subsubsection, Title → \\title, Subtitle → \\subtitle.
5. Convert tables to proper LaTeX tabular/booktabs environments with column
   separators inferred from the data.
6. Convert bulleted lists to itemize and numbered lists to enumerate.
7. For quotes use the quotation environment.
8. If the document contains a References or Bibliography section:
   a. Extract every reference and create a proper @article/@book/@misc BibTeX entry.
   b. Detect the citation style (APA, IEEE, MLA, Chicago, Vancouver, …) from
      the formatting pattern and choose the matching BibTeX style:
        APA       → apalike
        IEEE      → ieeetr
        Vancouver → unsrt
        MLA/Chicago → plain (or chicago with the chicago package)
        Numbered   → unsrt
   c. Insert \\cite{key} in the body text wherever an in-text citation appears.
9. Handle special characters and accents correctly.
10. Close the document with \\end{document}.
"""

USER_TEMPLATE = """\
Convert the following Word document content to LaTeX.

DOCUMENT CONTENT (JSON):
{content_json}

Respond with EXACTLY this structure — nothing else outside the markers:

===LATEX_START===
<complete LaTeX document>
===LATEX_END===

===BIB_START===
<BibTeX entries, or the single word EMPTY if there are no references>
===BIB_END===

===BIBSTYLE===
<BibTeX style name, e.g. apalike / ieeetr / plain / unsrt — or NONE if no references>
===BIBSTYLE_END===
"""


def convert_to_latex(content: dict) -> tuple[str, str]:
    """
    Call Claude (streaming) to convert extracted document content to LaTeX.
    Returns (latex_source, bib_source).  bib_source is empty string if no refs.
    """
    content_json = json.dumps(content, ensure_ascii=False, indent=2)
    user_message = USER_TEMPLATE.format(content_json=content_json)

    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=16000,
        thinking={"type": "adaptive"},
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    ) as stream:
        response = stream.get_final_message()

    full_text = "".join(
        block.text for block in response.content if block.type == "text"
    )

    # --- Parse markers ---
    def extract_between(start_tag: str, end_tag: str) -> str:
        match = re.search(
            re.escape(start_tag) + r"(.*?)" + re.escape(end_tag),
            full_text,
            re.DOTALL,
        )
        return match.group(1).strip() if match else ""

    latex_content = extract_between("===LATEX_START===", "===LATEX_END===")
    bib_raw = extract_between("===BIB_START===", "===BIB_END===")
    bib_style = extract_between("===BIBSTYLE===", "===BIBSTYLE_END===")

    bib_content = "" if bib_raw.upper() == "EMPTY" else bib_raw
    if bib_style.upper() == "NONE":
        bib_style = ""

    # Inject bibliography commands if Claude forgot them
    if bib_content and "\\end{document}" in latex_content:
        has_bib_cmd = (
            "\\bibliography{" in latex_content
            or "\\printbibliography" in latex_content
        )
        if not has_bib_cmd:
            style_cmd = f"\\bibliographystyle{{{bib_style}}}\n" if bib_style else ""
            latex_content = latex_content.replace(
                "\\end{document}",
                f"{style_cmd}\\bibliography{{references}}\n\\end{{document}}",
            )

    return latex_content, bib_content


# ---------------------------------------------------------------------------
# FastAPI endpoints
# ---------------------------------------------------------------------------

@app.post(
    "/convert",
    summary="Convert a Word document to LaTeX",
    response_description="ZIP archive containing document.tex and (optionally) references.bib",
)
async def convert_word_to_latex(
    file: UploadFile = File(..., description="Word document (.docx)"),
):
    # --- Validate ---
    if not file.filename:
        raise HTTPException(status_code=400, detail="No filename provided.")

    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail="Only .docx files are supported.",
        )

    file_bytes = await file.read()
    if not file_bytes:
        raise HTTPException(status_code=400, detail="The uploaded file is empty.")

    # --- Extract Word content ---
    try:
        content = extract_docx_content(file_bytes)
    except Exception as exc:
        raise HTTPException(
            status_code=422,
            detail=f"Failed to parse the Word document: {exc}",
        )

    if not content.get("elements"):
        raise HTTPException(
            status_code=422,
            detail="The document is empty or contains no readable content.",
        )

    # --- Convert with Claude ---
    try:
        latex_content, bib_content = convert_to_latex(content)
    except anthropic.APIError as exc:
        raise HTTPException(status_code=502, detail=f"Claude API error: {exc}")
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}")

    if not latex_content:
        raise HTTPException(
            status_code=500,
            detail="Claude did not return any LaTeX content.",
        )

    # --- Build ZIP ---
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("document.tex", latex_content.encode("utf-8"))
        if bib_content:
            zf.writestr("references.bib", bib_content.encode("utf-8"))
    zip_buffer.seek(0)

    base_name = re.sub(r"\.docx$", "", file.filename, flags=re.IGNORECASE)
    safe_name = re.sub(r"[^\w\-.]", "_", base_name)

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="{safe_name}_latex.zip"',
            "X-Has-Bibliography": "true" if bib_content else "false",
        },
    )


@app.get("/health", summary="Health check")
async def health_check():
    return {"status": "ok"}
