import asyncio
import io
import re
import json
import time
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from typing import Optional

import anthropic
from docx import Document
from docx.oxml.ns import qn
import os

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
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

# Limit concurrent Claude API calls to protect against rate-limit bursts.
# At 1000 simultaneous users most requests wait in queue rather than
# hammering the API and getting 429s.
_API_SEMAPHORE = asyncio.Semaphore(20)


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

3. LANGUAGE / BABEL (mandatory):
   - Detect the document language from its text content.
   - For French documents: add \\usepackage[french]{babel} to the preamble
     (after inputenc / fontenc). This handles French spacing rules, guillemets,
     section name translations (Chapitre, Table des matières…).
   - For English documents: add \\usepackage[english]{babel}.
   - For bilingual documents: \\usepackage[french,english]{babel} (last = main).
   - The \\usepackage[T1]{fontenc} already listed in rule 2 is REQUIRED for
     correct rendering of accented characters in all European languages.

4. PARAGRAPH FORMATTING (mandatory):
   - ALWAYS add \\setlength{\\parindent}{0pt} after \\begin{document} (no indent at start of paragraphs).
   - ALWAYS add \\setlength{\\parskip}{6pt} for spacing between paragraphs.
   - Do NOT use \\noindent individually; the global setting handles it.

5. HYPERREF LINKS: Use colorlinks=true with blue color for ALL link types.
   NEVER use colored boxes (pdfborder or default boxed links). Links must appear
   as blue-colored text, not wrapped in green/red/colored rectangles.

6. Formatting: bold→\\textbf{}, italic→\\textit{}, underline→\\underline{},
   superscript→\\textsuperscript{}, subscript→\\textsubscript{}.
7. Headings: Heading 1→\\section, 2→\\subsection, 3→\\subsubsection,
   Title→\\title{} + \\maketitle, Subtitle→use \\date{} or subtitle package.
8. Tables → tabular with booktabs (\\toprule, \\midrule, \\bottomrule).
9. Bulleted lists → itemize; numbered → enumerate.
10. Quotes → quotation environment.

11. IMAGES (important):
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

12. References / Bibliography — CRITICAL RULES:
    a. NEVER use \\begin{thebibliography}...\\end{thebibliography} in the LaTeX
       document. ALWAYS use an external BibTeX file referenced via
       \\bibliography{references}. Every single reference entry MUST appear
       between ===BIB_START=== and ===BIB_END=== as proper BibTeX.

    b. Convert EVERY reference found in the document — whether it appears as a
       numbered list, an author-year list, a bulleted list, or any other format —
       into a @article, @book, @inproceedings, @misc, or @online BibTeX entry.
       If you can only determine partial information, use @misc with as many
       fields as possible. Do NOT leave any reference unconverted.

    c. Detect citation style from the references section of the document.

    d. APA style (most common — Author, Year format):
       - ALWAYS add \\usepackage[round,authoryear]{natbib} to the preamble.
       - Use bibliographystyle{apalike}.
       - In-text parenthetical citation  → \\citep{key}   → produces (Author, Year)
       - In-text narrative citation       → \\citet{key}   → produces Author (Year)
       - NEVER use plain \\cite{} for APA; it produces [key] which is wrong.

    e. IEEE / numeric styles (ieeetr, unsrt, plain):
       - Do NOT add natbib.
       - Use plain \\cite{key} → produces [1], [2], …

    f. Default to APA/natbib if the citation style cannot be determined.

    g. In the LaTeX body, replace the original reference list with:
       \\bibliographystyle{apalike}   % or ieeetr / plain etc.
       \\bibliography{references}
       (do NOT reproduce the references as plain text in the document body)

13. Handle special characters and accents.
14. Close with \\end{document}.

15. MATHEMATICAL EQUATIONS:
    - Inline math: wrap in $...$
    - Display / numbered equation: use \\begin{equation}...\\end{equation}
    - Un-numbered display: \\[ ... \\]
    - Common symbols: \\alpha, \\beta, \\sum_{i=1}^{n}, \\int_a^b, \\frac{a}{b}, \\sqrt{x}
    - Superscripts from text (x²) → x^{2}; subscripts (x₁) → x_{1}
    - Always add \\usepackage{amsmath} (already listed in rule 2).

16. CODE BLOCKS AND VERBATIM:
    - Add \\usepackage{listings} and \\usepackage{xcolor} to the preamble.
    - Add this setup after the package declarations:
      \\lstset{basicstyle=\\ttfamily\\small, breaklines=true, frame=single,
               keywordstyle=\\color{blue}, commentstyle=\\color{gray},
               stringstyle=\\color{orange}}
    - For named code blocks use:
      \\begin{lstlisting}[language=Python]  % or Bash / SQL / C / Java / JavaScript
      ...code...
      \\end{lstlisting}
    - For generic command-line / unformatted text use \\begin{verbatim}...\\end{verbatim}.
    - Detect the language from context (Python keywords, SQL SELECT/FROM, bash $ prompts…).
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
<ALL BibTeX entries converted from every reference in the document.
 Use @article/@book/@inproceedings/@misc/@online as appropriate.
 If there are truly NO references anywhere in the document, write only: EMPTY>
===BIB_END===

===BIBSTYLE===
<BibTeX style name: apalike / ieeetr / plain / unsrt — or NONE if no references>
===BIBSTYLE_END===

===CITATION_TYPE===
<apa_natbib if style is APA (uses natbib + \\citep/\\citet) — or numeric if style is IEEE/Vancouver/plain>
===CITATION_TYPE_END===
"""

# Templates for chunked conversion (long documents)
_FIRST_CHUNK_TEMPLATE = """\
Convert the following Word document content to LaTeX (part 1 of {n_total}).

DOCUMENT CONTENT (JSON):
{content_json}

Generate the complete LaTeX preamble and body for these elements.
Do NOT include \\end{{document}} — more content follows in subsequent parts.

Respond with EXACTLY:

===LATEX_START===
<complete preamble + body up to (but NOT including) \\end{{document}}>
===LATEX_END===
"""

_MIDDLE_CHUNK_TEMPLATE = """\
Continue the LaTeX document (part {i} of {n_total}).

DOCUMENT CONTENT (JSON):
{content_json}

Output ONLY the raw LaTeX body lines for these elements.
No \\documentclass, no preamble, no \\end{{document}}, no markers.
"""

_LAST_CHUNK_TEMPLATE = """\
Finish the LaTeX document (final part {i} of {n_total}).

DOCUMENT CONTENT (JSON):
{content_json}

Respond with EXACTLY this structure:

===LATEX_CONTINUATION===
<body for these elements, ending with \\end{{document}}>
===LATEX_CONTINUATION_END===

===BIB_START===
<ALL BibTeX entries converted from every reference visible in THIS chunk.
 This chunk contains the bibliography/references section of the document.
 Convert EVERY item in the reference list to a proper BibTeX entry.
 If there are truly NO references, write only: EMPTY>
===BIB_END===

===BIBSTYLE===
<apalike / ieeetr / plain / unsrt — or NONE>
===BIBSTYLE_END===

===CITATION_TYPE===
<apa_natbib or numeric>
===CITATION_TYPE_END===
"""

# JSON chars thresholds
_SINGLE_CALL_LIMIT = 80_000   # below → one API call
_CHUNK_TARGET      = 30_000   # target JSON chars per chunk (smaller = faster parallel)
_SMALL_DOC_LIMIT   = 25_000   # below → use Haiku (faster) for single call

# Model selection
_MODEL_SMART = "claude-sonnet-4-6"          # preamble, BibTeX, complex logic
_MODEL_FAST  = "claude-haiku-4-5-20251001"  # middle chunks (body-only, 5× faster)

# LaTeX template overrides
_TEMPLATE_HINTS: dict[str, str] = {
    "article": "Use \\documentclass{article}.",
    "report":  "Use \\documentclass{report} with chapter-level sectioning.",
    "ieee":    "Use \\documentclass[conference]{IEEEtran} in two-column IEEE layout. "
               "Add \\usepackage{cite} instead of natbib for IEEE numeric citations.",
    "beamer":  "Use \\documentclass{beamer}. Wrap each logical section in "
               "\\begin{frame}{Title}...\\end{frame}. Choose a clean theme like 'Madrid'.",
}


def _extract_between(text: str, start: str, end: str) -> str:
    m = re.search(re.escape(start) + r"(.*?)" + re.escape(end), text, re.DOTALL)
    return m.group(1).strip() if m else ""


def _call_claude_sync(user_message: str, max_tokens: int = 8000,
                      template: str = "auto", model: str = _MODEL_SMART) -> str:
    """Blocking Claude call — run inside a thread via _call_claude_async."""
    system = SYSTEM_PROMPT
    if template in _TEMPLATE_HINTS:
        system = system + f"\n\nTEMPLATE OVERRIDE: {_TEMPLATE_HINTS[template]}"

    last_exc: Exception = RuntimeError("Unknown error")
    for attempt in range(3):
        try:
            with client.messages.stream(
                model=model,
                max_tokens=max_tokens,
                system=system,
                messages=[{"role": "user", "content": user_message}],
            ) as stream:
                response = stream.get_final_message()
            return "".join(block.text for block in response.content if block.type == "text")
        except anthropic.APIStatusError as exc:
            if exc.status_code in (500, 529) and attempt < 2:
                last_exc = exc
                time.sleep(2 ** attempt)
            else:
                raise
        except anthropic.APIConnectionError as exc:
            if attempt < 2:
                last_exc = exc
                time.sleep(2 ** attempt)
            else:
                raise
    raise last_exc


async def _call_claude_async(user_message: str, max_tokens: int = 8000,
                             template: str = "auto", model: str = _MODEL_SMART) -> str:
    """Non-blocking wrapper: runs the sync Claude call in a thread pool.

    The global semaphore caps concurrent API calls so we never flood the
    Anthropic API regardless of how many HTTP requests FastAPI is handling.
    """
    async with _API_SEMAPHORE:
        return await asyncio.to_thread(_call_claude_sync, user_message, max_tokens, template, model)


_BIBTEX_ENTRY_RE = re.compile(
    r'(@(?:article|book|inproceedings|proceedings|incollection|misc|'
    r'online|url|phdthesis|mastersthesis|techreport|conference)\s*\{[^@]+\})',
    re.DOTALL | re.IGNORECASE,
)


def _extract_bibtex_fallback(text: str) -> str:
    """Extract any BibTeX entries present anywhere in the raw Claude response."""
    entries = _BIBTEX_ENTRY_RE.findall(text)
    return "\n\n".join(e.strip() for e in entries)


def _thebibliography_to_bibtex(latex: str) -> tuple[str, str]:
    """If Claude generated \\begin{thebibliography} in the .tex, convert each
    \\bibitem to a minimal @misc entry and remove the environment from the .tex.
    Returns (cleaned_latex, extra_bib_entries).
    """
    m = re.search(
        r'\\begin\{thebibliography\}.*?\\end\{thebibliography\}',
        latex, re.DOTALL,
    )
    if not m:
        return latex, ""

    block = m.group(0)
    items = re.findall(r'\\bibitem(?:\[.*?\])?\{([^}]+)\}(.*?)(?=\\bibitem|\Z)', block, re.DOTALL)
    bib_entries = []
    for key, body in items:
        body_clean = re.sub(r'\s+', ' ', body).strip()
        bib_entries.append(
            f"@misc{{{key},\n  note = {{{body_clean}}},\n}}"
        )

    cleaned = latex[:m.start()] + "\\bibliography{references}\n" + latex[m.end():]
    return cleaned, "\n\n".join(bib_entries)


def _postprocess(latex_content: str, bib_raw: str, bib_style: str, citation_type: str,
                 raw_response: str = "") -> tuple[str, str]:
    """Inject bibliography and fix APA citations.

    Priority for BibTeX content:
    1. Between ===BIB_START=== / ===BIB_END=== markers
    2. @article/@book/... entries found anywhere in the raw response
    3. \\begin{thebibliography} converted to @misc entries
    """
    bib_raw = bib_raw.strip()
    bib_style = bib_style.strip()
    citation_type = citation_type.strip().lower()

    bib_content = "" if bib_raw.upper() == "EMPTY" else bib_raw
    if bib_style.upper() == "NONE":
        bib_style = ""

    # Fallback 1: scan entire raw response for BibTeX entries
    if not bib_content and raw_response:
        bib_content = _extract_bibtex_fallback(raw_response)

    # Fallback 2: convert \begin{thebibliography} if still nothing
    if not bib_content:
        latex_content, bib_content = _thebibliography_to_bibtex(latex_content)
    elif "\\begin{thebibliography}" in latex_content:
        # We have BibTeX — strip the thebibliography block from .tex
        latex_content, _ = _thebibliography_to_bibtex(latex_content)

    is_apa = "apa" in citation_type or bib_style == "apalike"

    if bib_content and is_apa:
        if "natbib" not in latex_content:
            latex_content = latex_content.replace(
                "\\begin{document}",
                "\\usepackage[round,authoryear]{natbib}\n\\begin{document}",
            )
        latex_content = re.sub(r'\\cite\{', r'\\citep{', latex_content)

    if bib_content and "\\end{document}" in latex_content:
        if "\\bibliography{" not in latex_content and "\\printbibliography" not in latex_content:
            style_cmd = f"\\bibliographystyle{{{bib_style}}}\n" if bib_style else ""
            latex_content = latex_content.replace(
                "\\end{document}",
                f"{style_cmd}\\bibliography{{references}}\n\\end{{document}}",
            )

    return latex_content, bib_content


_REF_HEADINGS = {
    # English
    "references", "bibliography", "works cited", "literature",
    "sources", "reference list", "cited works", "citations",
    # French
    "references bibliographiques", "bibliographie", "references",
    "liste de references", "liste des references", "sources bibliographiques",
    "ouvrages cites", "ouvrages consultes", "travaux cites",
    # Generic
    "literature cited", "further reading", "notes and references",
}


def _normalise_text(t: str) -> str:
    """Lowercase, strip accents, collapse whitespace."""
    return unicodedata.normalize("NFD", t).encode("ascii", "ignore").decode().lower().strip()


def _find_references_section(elements: list) -> tuple[list, list]:
    """Split elements into (body_elements, references_elements).

    Searches from the end for a heading/paragraph whose normalised text
    matches or starts with a known bibliography keyword.  Everything from
    that element onward goes to the last chunk so Claude always sees the full
    reference list when generating BibTeX.
    """
    for i in range(len(elements) - 1, -1, -1):
        elem = elements[i]
        if elem.get("type") in ("heading", "paragraph"):
            t = _normalise_text(elem.get("text", ""))
            if any(t == h or t.startswith(h) for h in _REF_HEADINGS):
                return elements[:i], elements[i:]   # heading stays with refs
    return elements, []


def _chunk_elements(elements: list) -> list[list]:
    """Split elements into groups whose JSON size ≈ _CHUNK_TARGET chars."""
    chunks, current, current_len = [], [], 0
    for elem in elements:
        s = len(json.dumps(elem, ensure_ascii=False))
        if current and current_len + s > _CHUNK_TARGET:
            chunks.append(current)
            current, current_len = [elem], s
        else:
            current.append(elem)
            current_len += s
    if current:
        chunks.append(current)
    return chunks


async def _convert_single(content: dict, template: str = "auto",
                          model: str = _MODEL_SMART) -> tuple[str, str, str]:
    content_json = json.dumps(content, ensure_ascii=False, indent=2)
    full_text = await _call_claude_async(
        USER_TEMPLATE.format(content_json=content_json),
        max_tokens=16000, template=template, model=model,
    )

    latex_content = _extract_between(full_text, "===LATEX_START===", "===LATEX_END===")
    bib_raw       = _extract_between(full_text, "===BIB_START===",   "===BIB_END===")
    bib_style     = _extract_between(full_text, "===BIBSTYLE===",     "===BIBSTYLE_END===")
    citation_type = _extract_between(full_text, "===CITATION_TYPE===", "===CITATION_TYPE_END===").lower()

    # Fallback regexes if markers missing
    if not latex_content:
        m = re.search(r"(\\documentclass.*?\\end\{document\})", full_text, re.DOTALL)
        if m:
            latex_content = m.group(1).strip()
    if not latex_content:
        m = re.search(r"```(?:latex|tex)?\s*(\\documentclass.*?\\end\{document\})\s*```", full_text, re.DOTALL)
        if m:
            latex_content = m.group(1).strip()

    latex_content, bib_content = _postprocess(latex_content, bib_raw, bib_style, citation_type, full_text)
    return latex_content, bib_content, full_text


async def _convert_chunked(content: dict, template: str = "auto") -> tuple[str, str, str]:
    all_elements = content.get("elements", [])
    metadata     = content.get("metadata", {})
    image_list   = content.get("image_list", [])

    # Always keep the references section in the last chunk so Claude can
    # generate complete BibTeX regardless of how the document is split.
    body_elements, ref_elements = _find_references_section(all_elements)
    chunks = _chunk_elements(body_elements)

    if ref_elements:
        if chunks:
            chunks[-1] = chunks[-1] + ref_elements
        else:
            chunks = [ref_elements]

    n = len(chunks)
    if n == 1:
        return await _convert_single(content, template)

    # -----------------------------------------------------------------------
    # Build one prompt per chunk, then fire ALL in parallel.
    #
    # Strategy:
    #   chunk[0]   → _FIRST_CHUNK_TEMPLATE  → _MODEL_SMART (preamble + packages)
    #   chunk[1..n-2] → _MIDDLE_CHUNK_TEMPLATE → _MODEL_FAST  (body only, Haiku)
    #   chunk[n-1] → _LAST_CHUNK_TEMPLATE   → _MODEL_SMART (BibTeX + closing)
    #
    # Since every chunk is independent (no chunk reads the output of another),
    # they can all run simultaneously.  Total wall-clock time ≈ slowest chunk
    # instead of sum of all chunks.
    # -----------------------------------------------------------------------

    def _chunk_json(elements, meta=None, imgs=None):
        return json.dumps(
            {"metadata": meta or {}, "elements": elements, "image_list": imgs or []},
            ensure_ascii=False, indent=2,
        )

    tasks = []
    # First chunk
    tasks.append(_call_claude_async(
        _FIRST_CHUNK_TEMPLATE.format(n_total=n, content_json=_chunk_json(chunks[0], metadata, image_list)),
        max_tokens=8000, template=template, model=_MODEL_SMART,
    ))
    # Middle chunks — Haiku (fast, body-text only)
    for i, chunk in enumerate(chunks[1:-1], start=2):
        tasks.append(_call_claude_async(
            _MIDDLE_CHUNK_TEMPLATE.format(i=i, n_total=n, content_json=_chunk_json(chunk)),
            max_tokens=6000, template=template, model=_MODEL_FAST,
        ))
    # Last chunk
    tasks.append(_call_claude_async(
        _LAST_CHUNK_TEMPLATE.format(i=n, n_total=n, content_json=_chunk_json(chunks[-1])),
        max_tokens=8000, template=template, model=_MODEL_SMART,
    ))

    results: list[str] = await asyncio.gather(*tasks)

    first_text   = results[0]
    middle_texts = results[1:-1]
    last_text    = results[-1]

    # --- Assemble in order ---
    latex_body = _extract_between(first_text, "===LATEX_START===", "===LATEX_END===")
    if not latex_body:
        m = re.search(r"(\\documentclass.*)", first_text, re.DOTALL)
        latex_body = m.group(1).strip() if m else first_text.strip()
    latex_body = re.sub(r'\\end\{document\}\s*$', '', latex_body).rstrip()

    for text in middle_texts:
        latex_body += "\n\n" + text.strip()

    last_body     = _extract_between(last_text, "===LATEX_CONTINUATION===", "===LATEX_CONTINUATION_END===")
    bib_raw       = _extract_between(last_text, "===BIB_START===",          "===BIB_END===")
    bib_style     = _extract_between(last_text, "===BIBSTYLE===",            "===BIBSTYLE_END===")
    citation_type = _extract_between(last_text, "===CITATION_TYPE===",       "===CITATION_TYPE_END===").lower()

    if not last_body:
        last_body = last_text.strip()

    latex_content = latex_body + "\n\n" + last_body
    if "\\end{document}" not in latex_content:
        latex_content += "\n\\end{document}"

    full_raw = "\n\n---CHUNK BREAK---\n\n".join(results)
    latex_content, bib_content = _postprocess(latex_content, bib_raw, bib_style, citation_type, full_raw)
    return latex_content, bib_content, full_raw


async def convert_to_latex(content: dict, template: str = "auto") -> tuple[str, str, str]:
    content_json_size = len(json.dumps(content, ensure_ascii=False))
    if content_json_size <= _SINGLE_CALL_LIMIT:
        # Small documents: use Haiku — much faster, fully capable for simple docs
        model = _MODEL_FAST if content_json_size <= _SMALL_DOC_LIMIT else _MODEL_SMART
        return await _convert_single(content, template, model=model)
    return await _convert_chunked(content, template)


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
    template: str = Form("auto", description="LaTeX template: auto | article | report | ieee | beamer"),
):
    if not file.filename:
        raise HTTPException(status_code=400, detail="No filename provided.")
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported.")

    if template not in ("auto", "article", "report", "ieee", "beamer"):
        raise HTTPException(status_code=400, detail=f"Unknown template '{template}'. Choose: auto, article, report, ieee, beamer.")

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

    # Convert to LaTeX via Claude (non-blocking async)
    try:
        latex_content, bib_content, raw_response = await convert_to_latex(content, template)
    except anthropic.APIError as exc:
        raise HTTPException(status_code=502, detail=f"Claude API error: {exc}")
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}")

    if not latex_content:
        preview = raw_response[:500].replace("\n", " ") if raw_response else "(empty response)"
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
