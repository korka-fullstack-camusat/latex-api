import io
import re
import json
import zipfile
import xml.etree.ElementTree as ET
from typing import Optional

import anthropic
from docx import Document
from docx.oxml.ns import qn
import os

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

app = FastAPI(
    title="Word to LaTeX Converter",
    description="Upload a Word document (.docx) and get a compilable LaTeX file + references.bib + images",
)

# ---------------------------------------------------------------------------
# CORS — allow the Next.js frontend to call the API from the browser.
# Set ALLOWED_ORIGINS to a comma-separated list of origins in production,
# e.g.  ALLOWED_ORIGINS=https://myapp.com,https://www.myapp.com
# ---------------------------------------------------------------------------
_raw_origins = os.getenv("ALLOWED_ORIGINS", "*")
_allowed_origins = [o.strip() for o in _raw_origins.split(",") if o.strip()]

app.add_middleware(
    CORSMiddleware,
    allow_origins=_allowed_origins,
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
    # Expose custom response headers so the browser JS can read them.
    expose_headers=["Content-Disposition", "X-Has-Bibliography", "X-Image-Count"],
)

client = anthropic.Anthropic()


# ---------------------------------------------------------------------------
# Image extraction
# ---------------------------------------------------------------------------

# Map MIME / extension → LaTeX-friendly extension
EXT_MAP = {
    "jpeg": "jpg", "jpg": "jpg", "png": "png",
    "gif": "png",  "bmp": "png", "tiff": "png",
    "emf": "pdf",  "wmf": "pdf",
}

def extract_images_from_docx(file_bytes: bytes) -> dict[str, bytes]:
    """
    Open the .docx ZIP and extract every media file.
    Returns {relationship_id: (clean_filename, raw_bytes)}.
    """
    rel_to_image: dict[str, tuple[str, bytes]] = {}

    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        # Parse word/_rels/document.xml.rels to map rId → media path
        try:
            rels_xml = z.read("word/_rels/document.xml.rels")
        except KeyError:
            return {}

        ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
        rels_tree = ET.fromstring(rels_xml)
        counter = 1

        for rel in rels_tree.findall("r:Relationship", ns):
            rel_type = rel.get("Type", "")
            if "image" not in rel_type.lower():
                continue

            rel_id = rel.get("Id", "")
            target = rel.get("Target", "")

            # Resolve path inside the ZIP
            if target.startswith("/"):
                zip_path = target.lstrip("/")
            else:
                zip_path = f"word/{target}"

            try:
                img_bytes = z.read(zip_path)
            except KeyError:
                continue

            # Determine a clean extension
            orig_ext = zip_path.rsplit(".", 1)[-1].lower() if "." in zip_path else "png"
            ext = EXT_MAP.get(orig_ext, orig_ext)
            clean_name = f"figure_{counter}.{ext}"
            counter += 1

            rel_to_image[rel_id] = (clean_name, img_bytes)

    return rel_to_image


# ---------------------------------------------------------------------------
# Paragraph / table extraction
# ---------------------------------------------------------------------------

# VML namespace not registered in all python-docx versions — use URI directly
_VML_IMAGEDATA = "{urn:schemas-microsoft-com:vml}imagedata"
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def get_image_refs_in_para(para, rel_to_image: dict) -> list[str]:
    """Return list of image filenames embedded in this paragraph."""
    refs = []
    # w:drawing → wp:inline/wp:anchor → a:blip r:embed
    for blip in para._element.iter(qn("a:blip")):
        r_embed = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
        if r_embed and r_embed in rel_to_image:
            refs.append(rel_to_image[r_embed][0])
    # Also catch v:imagedata (older .docx) — use hardcoded URI, not qn("v:...")
    for imgdata in para._element.iter(_VML_IMAGEDATA):
        r_id = imgdata.get(_R_ID)
        if r_id and r_id in rel_to_image:
            refs.append(rel_to_image[r_id][0])
    return refs


def extract_paragraph_info(para, rel_to_image: dict) -> Optional[dict]:
    """Return structured info for a paragraph, or None if empty and no images."""
    style_name = para.style.name if para.style else "Normal"

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

    # Detect list via XML
    num_pr = para._element.find(qn("w:numPr"))
    if num_pr is not None:
        para_type = "list_item"
        ilvl_el = num_pr.find(qn("w:ilvl"))
        if ilvl_el is not None:
            level = int(ilvl_el.get(qn("w:val"), 0))

    # Detect embedded images
    image_refs = get_image_refs_in_para(para, rel_to_image)
    if image_refs:
        para_type = "image"

    # Formatted runs
    runs = []
    for run in para.runs:
        if run.text:
            runs.append({
                "text": run.text,
                "bold": bool(run.bold),
                "italic": bool(run.italic),
                "underline": bool(run.underline),
                "superscript": bool(run.font.superscript) if run.font.superscript else False,
                "subscript": bool(run.font.subscript) if run.font.subscript else False,
            })

    if not runs and not para.text.strip() and not image_refs:
        return None

    info: dict = {
        "type": para_type,
        "style": style_name,
        "level": level,
        "text": para.text,
        "runs": runs,
        "alignment": str(para.alignment) if para.alignment else None,
    }
    if image_refs:
        info["images"] = image_refs   # e.g. ["figure_1.png", "figure_2.png"]

    return info


def extract_table_info(table) -> dict:
    rows = []
    for row in table.rows:
        row_data = [
            " ".join(p.text for p in cell.paragraphs).strip()
            for cell in row.cells
        ]
        rows.append(row_data)
    return {
        "type": "table",
        "rows": rows,
        "num_cols": len(rows[0]) if rows else 0,
        "num_rows": len(rows),
    }


def extract_docx_content(file_bytes: bytes) -> tuple[dict, dict[str, bytes]]:
    """
    Returns:
      content  – structured document (JSON-serialisable)
      images   – {clean_filename: raw_bytes}  to include in ZIP
    """
    doc = Document(io.BytesIO(file_bytes))
    rel_to_image = extract_images_from_docx(file_bytes)
    # flat map: clean_filename → bytes (what the ZIP needs)
    images_bytes: dict[str, bytes] = {v[0]: v[1] for v in rel_to_image.values()}

    content: dict = {"metadata": {}, "elements": [], "image_list": sorted(images_bytes)}

    try:
        props = doc.core_properties
        content["metadata"] = {
            "title": props.title or "",
            "author": props.author or "",
            "subject": props.subject or "",
        }
    except Exception:
        pass

    para_map = {p._element: p for p in doc.paragraphs}
    table_map = {t._element: t for t in doc.tables}

    for child in doc.element.body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p" and child in para_map:
            info = extract_paragraph_info(para_map[child], rel_to_image)
            if info:
                content["elements"].append(info)

        elif tag == "tbl" and child in table_map:
            content["elements"].append(extract_table_info(table_map[child]))

    return content, images_bytes


# ---------------------------------------------------------------------------
# LaTeX conversion via Claude
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """\
You are an expert LaTeX document converter.
Convert structured Word document content (JSON) into a complete, compilable
LaTeX document that reproduces the original as faithfully as possible.

Rules:
1. Choose document class (article / report / book) based on content.
2. Always include these packages in the preamble:
   \\usepackage[utf8]{inputenc}
   \\usepackage[T1]{fontenc}
   \\usepackage[margin=2.5cm]{geometry}
   \\usepackage{graphicx}
   \\usepackage{booktabs}
   \\usepackage{amsmath}
   \\usepackage{caption}
   \\usepackage{float}
   \\usepackage{setspace}
   \\usepackage{parskip}
   \\usepackage[colorlinks=true, linkcolor=blue, citecolor=blue, urlcolor=blue, filecolor=blue]{hyperref}

3. PARAGRAPH FORMATTING (mandatory):
   - ALWAYS add \\setlength{\\parindent}{0pt} after \\begin{document} (no indent at start of paragraphs).
   - ALWAYS add \\setlength{\\parskip}{6pt} for spacing between paragraphs.
   - Do NOT use \\noindent individually; the global setting handles it.

4. HYPERREF LINKS: Use colorlinks=true with blue color for ALL link types.
   NEVER use colored boxes (pdfborder or default boxed links). Links must appear
   as blue-colored text, not wrapped in green/red/colored rectangles.

5. Formatting: bold→\\textbf{}, italic→\\textit{}, underline→\\underline{},
   superscript→\\textsuperscript{}, subscript→\\textsubscript{}.
6. Headings: Heading 1→\\section, 2→\\subsection, 3→\\subsubsection,
   Title→\\title{} + \\maketitle, Subtitle→use \\date{} or subtitle package.
7. Tables → tabular with booktabs (\\toprule, \\midrule, \\bottomrule).
8. Bulleted lists → itemize; numbered → enumerate.
9. Quotes → quotation environment.

10. IMAGES (important):
    Each element with "type":"image" contains an "images" list of filenames
    (e.g. ["figure_1.png"]).  These files ARE included in the ZIP alongside
    the .tex file.
    For every image element generate:

    \\begin{figure}[H]
      \\centering
      \\includegraphics[width=0.8\\linewidth]{figure_1}
      \\caption{<caption text if next element is a Caption, else leave descriptive placeholder>}
      \\label{fig:figure_1}
    \\end{figure}

    - Use the filename WITHOUT extension in \\includegraphics{}.
    - If the following element has type "caption", use its text and skip it as
      a standalone paragraph.
    - If there is no caption, write a short descriptive placeholder like
      \\caption{Figure extracted from document}.
    - Always use the float package option [H] so figures stay in place.

11. References / Bibliography:
    a. Detect citation style from the references section of the document.
    b. Create proper @article/@book/@misc BibTeX entries.

    c. APA style (most common — Author, Year format):
       - ALWAYS add \\usepackage[round,authoryear]{natbib} to the preamble.
       - Use bibliographystyle{apalike}.
       - In-text parenthetical citation  → \\citep{key}   → produces (Author, Year)
       - In-text narrative citation       → \\citet{key}   → produces Author (Year)
       - NEVER use plain \\cite{} for APA; it produces [key] which is wrong.

    d. IEEE / numeric styles (ieeetr, unsrt, plain):
       - Do NOT add natbib.
       - Use plain \\cite{key} → produces [1], [2], …

    e. Default to APA/natbib if the citation style cannot be determined.

12. Handle special characters and accents.
13. Close with \\end{document}.
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
<BibTeX style name: apalike / ieeetr / plain / unsrt — or NONE if no references>
===BIBSTYLE_END===

===CITATION_TYPE===
<apa_natbib if style is APA (uses natbib + \\citep/\\citet) — or numeric if style is IEEE/Vancouver/plain>
===CITATION_TYPE_END===
"""


def convert_to_latex(content: dict) -> tuple[str, str]:
    content_json = json.dumps(content, ensure_ascii=False, indent=2)
    user_message = USER_TEMPLATE.format(content_json=content_json)

    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=16000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    ) as stream:
        response = stream.get_final_message()

    full_text = "".join(
        block.text for block in response.content if block.type == "text"
    )

    def extract_between(start: str, end: str) -> str:
        m = re.search(re.escape(start) + r"(.*?)" + re.escape(end), full_text, re.DOTALL)
        return m.group(1).strip() if m else ""

    latex_content = extract_between("===LATEX_START===", "===LATEX_END===")
    bib_raw       = extract_between("===BIB_START===",   "===BIB_END===")
    bib_style     = extract_between("===BIBSTYLE===",     "===BIBSTYLE_END===")
    citation_type = extract_between("===CITATION_TYPE===", "===CITATION_TYPE_END===").lower()

    # Fallback: if markers were missing, try to find a LaTeX document in the raw text
    if not latex_content:
        # Look for \documentclass ... \end{document}
        m = re.search(r"(\\documentclass.*?\\end\{document\})", full_text, re.DOTALL)
        if m:
            latex_content = m.group(1).strip()
        # Also try markdown code fences: ```latex ... ``` or ``` ... ```
        if not latex_content:
            m = re.search(r"```(?:latex|tex)?\s*(\\documentclass.*?\\end\{document\})\s*```", full_text, re.DOTALL)
            if m:
                latex_content = m.group(1).strip()

    bib_content = "" if bib_raw.upper() == "EMPTY" else bib_raw
    if bib_style.upper() == "NONE":
        bib_style = ""

    is_apa = "apa" in citation_type or bib_style == "apalike"

    # Ensure natbib is in the preamble for APA documents
    if bib_content and is_apa:
        if "natbib" not in latex_content:
            latex_content = latex_content.replace(
                "\\begin{document}",
                "\\usepackage[round,authoryear]{natbib}\n\\begin{document}",
            )
        # Replace any leftover plain \cite{ with \citep{ for APA
        latex_content = re.sub(r'\\cite\{', r'\\citep{', latex_content)

    # Inject bibliography commands if Claude forgot
    if bib_content and "\\end{document}" in latex_content:
        if "\\bibliography{" not in latex_content and "\\printbibliography" not in latex_content:
            style_cmd = f"\\bibliographystyle{{{bib_style}}}\n" if bib_style else ""
            latex_content = latex_content.replace(
                "\\end{document}",
                f"{style_cmd}\\bibliography{{references}}\n\\end{{document}}",
            )

    return latex_content, bib_content


# ---------------------------------------------------------------------------
# FastAPI endpoint
# ---------------------------------------------------------------------------

@app.post(
    "/convert",
    summary="Convert a Word document to LaTeX",
    response_description="ZIP: document.tex + references.bib (if any) + all extracted images",
)
async def convert_word_to_latex(
    file: UploadFile = File(..., description="Word document (.docx)"),
):
    if not file.filename:
        raise HTTPException(status_code=400, detail="No filename provided.")
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported.")

    file_bytes = await file.read()
    if not file_bytes:
        raise HTTPException(status_code=400, detail="The uploaded file is empty.")

    # Extract content + images
    try:
        content, images_bytes = extract_docx_content(file_bytes)
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"Failed to parse Word document: {exc}")

    if not content.get("elements"):
        raise HTTPException(status_code=422, detail="The document is empty or has no readable content.")

    # Convert to LaTeX via Claude
    try:
        latex_content, bib_content = convert_to_latex(content)
    except anthropic.APIError as exc:
        raise HTTPException(status_code=502, detail=f"Claude API error: {exc}")
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}")

    if not latex_content:
        # Log first 500 chars to help diagnose future issues
        preview = "(empty response)"
        raise HTTPException(
            status_code=500,
            detail=f"Claude did not return any LaTeX content. Response preview: {preview}",
        )

    # Build ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("document.tex", latex_content.encode("utf-8"))

        if bib_content:
            zf.writestr("references.bib", bib_content.encode("utf-8"))

        # Include every extracted image at the root of the ZIP
        for img_name, img_data in images_bytes.items():
            zf.writestr(img_name, img_data)

    zip_buffer.seek(0)

    base_name = re.sub(r"\.docx$", "", file.filename, flags=re.IGNORECASE)
    safe_name  = re.sub(r"[^\w\-.]", "_", base_name)

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="{safe_name}_latex.zip"',
            "X-Has-Bibliography": "true" if bib_content else "false",
            "X-Image-Count": str(len(images_bytes)),
        },
    )


@app.get("/health", summary="Health check")
async def health_check():
    return {"status": "ok"}
