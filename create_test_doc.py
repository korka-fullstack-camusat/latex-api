"""
Generate a test .docx document with:
  - title, abstract, multiple sections
  - 3 embedded PNG images (generated with pure Python, no PIL required)
  - multiple in-text APA citations
  - a formatted References section in APA style
"""
import io
import struct
import zlib
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ---------------------------------------------------------------------------
# Minimal PNG generator (no PIL / Pillow needed)
# ---------------------------------------------------------------------------

def _png_chunk(chunk_type: bytes, data: bytes) -> bytes:
    crc = zlib.crc32(chunk_type + data) & 0xFFFFFFFF
    return struct.pack(">I", len(data)) + chunk_type + data + struct.pack(">I", crc)


def make_solid_png(width: int, height: int, r: int, g: int, b: int) -> bytes:
    """Return raw bytes of a solid-colour PNG."""
    # IHDR
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    # Raw image data: filter byte 0 + RGB pixels per row
    raw_rows = b""
    row = bytes([0] + [r, g, b] * width)
    for _ in range(height):
        raw_rows += row
    idat_data = zlib.compress(raw_rows)
    return (
        b"\x89PNG\r\n\x1a\n"
        + _png_chunk(b"IHDR", ihdr)
        + _png_chunk(b"IDAT", idat_data)
        + _png_chunk(b"IEND", b"")
    )


# Three distinct images
IMAGES = [
    ("bar_chart.png",    make_solid_png(400, 220, 70,  130, 200)),   # blue
    ("scatter_plot.png", make_solid_png(400, 220, 200, 100,  60)),   # orange
    ("heatmap.png",      make_solid_png(400, 220,  80, 170,  90)),   # green
]


# ---------------------------------------------------------------------------
# Build the document
# ---------------------------------------------------------------------------

doc = Document()

# ── Title ──────────────────────────────────────────────────────────────────
title = doc.add_heading("Impact of Digital Learning Tools on Academic Performance", 0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ── Authors / date ─────────────────────────────────────────────────────────
p = doc.add_paragraph("Marie Dupont, Jean-Paul Martin, Amara Diallo")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.runs[0].bold = True

p2 = doc.add_paragraph("Université de Lyon — March 2025")
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()

# ── Abstract ───────────────────────────────────────────────────────────────
doc.add_heading("Abstract", level=1)
doc.add_paragraph(
    "This study examines the relationship between the adoption of digital learning "
    "tools and academic outcomes in higher education. Drawing on a sample of 1,240 "
    "undergraduate students across five French universities, we find that structured "
    "use of interactive platforms significantly improves examination scores "
    "(Dupont & Martin, 2022). These results align with the self-regulated learning "
    "framework proposed by Zimmerman (2002) and extend the meta-analytic findings "
    "reported by Means et al. (2013)."
)

# ── 1. Introduction ────────────────────────────────────────────────────────
doc.add_heading("1. Introduction", level=1)
doc.add_paragraph(
    "The rapid expansion of e-learning technologies has reshaped pedagogical "
    "practices worldwide (OECD, 2021). Despite abundant anecdotal evidence, "
    "empirical research on the causal mechanisms linking tool usage to learning "
    "outcomes remains limited (Selwyn, 2016). Early work by Clark (1994) argued "
    "that media per se do not influence learning; rather, it is the instructional "
    "method embedded in the medium that matters."
)
doc.add_paragraph(
    "More recently, a meta-analysis by Means et al. (2013) covering 50 controlled "
    "experiments concluded that online instruction produced modestly better outcomes "
    "than face-to-face instruction on average (Cohen's d = 0.20, 95% CI [0.12, 0.28]). "
    "Building on this foundation, the present study focuses specifically on "
    "formative-assessment platforms."
)

# ── 2. Theoretical Framework ───────────────────────────────────────────────
doc.add_heading("2. Theoretical Framework", level=1)
doc.add_paragraph(
    "Our conceptual model integrates two complementary theories. First, "
    "self-regulated learning theory (Zimmerman, 2002) posits that learners who "
    "monitor and adjust their study strategies achieve higher performance. Second, "
    "cognitive load theory (Sweller, 1988) suggests that well-designed interfaces "
    "reduce extraneous load and free working memory for deep processing."
)

doc.add_heading("2.1 Self-Regulated Learning", level=2)
doc.add_paragraph(
    "Zimmerman (2002) describes self-regulated learning as a cyclical process "
    "comprising forethought, performance, and self-reflection phases. Digital "
    "dashboards that visualise progress have been shown to support all three phases "
    "(Tanes et al., 2011; Dupont & Martin, 2022)."
)

doc.add_heading("2.2 Cognitive Load", level=2)
doc.add_paragraph(
    "Sweller (1988) originally identified three types of cognitive load: intrinsic, "
    "extraneous, and germane. Interface design that minimises extraneous load — e.g. "
    "clear navigation, concise feedback — is particularly important for novice learners "
    "(Clark, 1994; Selwyn, 2016)."
)

# ── 3. Methodology ─────────────────────────────────────────────────────────
doc.add_heading("3. Methodology", level=1)
doc.add_paragraph(
    "We employed a quasi-experimental design with pre- and post-test measurements "
    "across two semesters (September 2023 – June 2024). Treatment groups used "
    "the Moodle-based adaptive quiz module for at least three hours per week; "
    "control groups followed traditional paper-based exercises."
)

# Image 1 — bar chart
doc.add_paragraph()
doc.add_picture(io.BytesIO(IMAGES[0][1]), width=Inches(4.5))
last = doc.paragraphs[-1]
last.alignment = WD_ALIGN_PARAGRAPH.CENTER
cap1 = doc.add_paragraph("Figure 1. Mean pre- and post-test scores by group (treatment vs. control).")
cap1.style = doc.styles["Caption"] if "Caption" in doc.styles else doc.styles["Normal"]
cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph(
    "Figure 1 displays the mean scores before and after intervention. The treatment "
    "group improved by 14.3 percentage points on average, compared with 5.7 points "
    "for the control group (Dupont & Martin, 2022)."
)

# ── 4. Results ─────────────────────────────────────────────────────────────
doc.add_heading("4. Results", level=1)
doc.add_paragraph(
    "A two-way repeated-measures ANOVA revealed a significant Group × Time "
    "interaction, F(1, 1238) = 84.3, p < .001, η² = .06. Post-hoc comparisons "
    "confirmed that the treatment group outperformed controls at post-test "
    "(M = 72.4 vs. M = 61.8; Cohen's d = 0.58). These effect sizes are larger "
    "than those reported by Means et al. (2013) but consistent with interventions "
    "that include personalised feedback (Tanes et al., 2011)."
)

# Image 2 — scatter plot
doc.add_paragraph()
doc.add_picture(io.BytesIO(IMAGES[1][1]), width=Inches(4.5))
last = doc.paragraphs[-1]
last.alignment = WD_ALIGN_PARAGRAPH.CENTER
cap2 = doc.add_paragraph(
    "Figure 2. Scatter plot of weekly platform usage (hours) vs. final examination score (%)."
)
cap2.style = doc.styles["Caption"] if "Caption" in doc.styles else doc.styles["Normal"]
cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph(
    "As shown in Figure 2, platform engagement (measured in hours per week) "
    "correlates positively with final examination performance (r = .47, p < .001), "
    "in line with the self-regulation mechanisms described by Zimmerman (2002)."
)

doc.add_heading("4.1 Moderating Variables", level=2)
doc.add_paragraph(
    "Prior academic achievement moderated the treatment effect (β = −0.31, p = .012), "
    "suggesting that lower-achieving students benefited most from the adaptive "
    "feedback — a pattern consistent with findings in OECD (2021)."
)

# Image 3 — heatmap
doc.add_paragraph()
doc.add_picture(io.BytesIO(IMAGES[2][1]), width=Inches(4.5))
last = doc.paragraphs[-1]
last.alignment = WD_ALIGN_PARAGRAPH.CENTER
cap3 = doc.add_paragraph(
    "Figure 3. Heatmap of feature correlations: usage metrics, prior GPA, and exam score."
)
cap3.style = doc.styles["Caption"] if "Caption" in doc.styles else doc.styles["Normal"]
cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ── 5. Discussion ──────────────────────────────────────────────────────────
doc.add_heading("5. Discussion", level=1)
doc.add_paragraph(
    "Our findings confirm that structured use of digital formative-assessment tools "
    "can meaningfully enhance academic outcomes, echoing the broader literature "
    "(Means et al., 2013; Selwyn, 2016). The moderating role of prior achievement "
    "aligns with equity-oriented arguments advanced by OECD (2021): adaptive systems "
    "may help close performance gaps."
)
doc.add_paragraph(
    "However, Clark (1994)'s caveat remains relevant: the gains observed here may "
    "reflect the instructional design of the quizzes rather than the digital medium "
    "itself. Future research should disentangle these components using platform-blind "
    "designs (Sweller, 1988)."
)

# ── 6. Conclusion ──────────────────────────────────────────────────────────
doc.add_heading("6. Conclusion", level=1)
doc.add_paragraph(
    "This study provides quasi-experimental evidence that adaptive digital quizzes "
    "improve undergraduate examination scores. The effect is strongest for "
    "lower-achieving students, pointing to equity benefits. Institutions should "
    "integrate formative-assessment platforms within broader self-regulation support "
    "programmes (Zimmerman, 2002; Dupont & Martin, 2022)."
)

# ── References (APA 7th edition) ───────────────────────────────────────────
doc.add_heading("References", level=1)

refs = [
    ("Clark, R. E. (1994). Media will never influence learning. "
     "Educational Technology Research and Development, 42(2), 21–29. "
     "https://doi.org/10.1007/BF02299088"),

    ("Dupont, M., & Martin, J.-P. (2022). Formative assessment platforms and "
     "self-regulation in French higher education. Journal of Educational Technology, "
     "19(3), 145–162. https://doi.org/10.1016/j.jedtech.2022.03.005"),

    ("Means, B., Toyama, Y., Murphy, R., & Bakia, M. (2013). The effectiveness of "
     "online and blended learning: A meta-analysis of the empirical literature. "
     "Teachers College Record, 115(3), 1–47."),

    ("OECD. (2021). 21st-century readers: Developing literacy skills in a digital "
     "world. OECD Publishing. https://doi.org/10.1787/a83d84cb-en"),

    ("Selwyn, N. (2016). Is technology good for education? Polity Press."),

    ("Sweller, J. (1988). Cognitive load during problem solving: Effects on learning. "
     "Cognitive Science, 12(2), 257–285. https://doi.org/10.1207/s15516709cog1202_4"),

    ("Tanes, Z., Arnold, K. E., King, A. S., & Remnet, M. A. (2011). Using Signals for "
     "appropriate feedback: Perceptions and practices. Computers & Education, 57(4), "
     "2414–2422. https://doi.org/10.1016/j.compedu.2011.05.016"),

    ("Zimmerman, B. J. (2002). Becoming a self-regulated learner: An overview. "
     "Theory into Practice, 41(2), 64–70. https://doi.org/10.1207/s15430421tip4102_2"),
]

for ref in refs:
    p = doc.add_paragraph(ref, style="Normal")
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.5)

# Save
out_path = "/home/user/test_apa_with_images.docx"
doc.save(out_path)
print(f"Document saved → {out_path}")
