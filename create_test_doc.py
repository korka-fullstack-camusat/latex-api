"""
Generate a test .docx document with:
  - title + authors at the top
  - multiple sections with APA citations
  - 3 real matplotlib charts (bar chart, scatter plot, heatmap)
  - References / Sources section at the bottom
"""
import io
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.cm as cm

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ---------------------------------------------------------------------------
# Real chart generators
# ---------------------------------------------------------------------------

def make_bar_chart() -> bytes:
    """Pre/post test scores by group — bar chart."""
    fig, ax = plt.subplots(figsize=(6, 3.5))
    groups = ["Control\n(n=620)", "Treatment\n(n=620)"]
    pre  = [60.1, 59.8]
    post = [65.8, 74.1]
    x = np.arange(len(groups))
    w = 0.35
    bars1 = ax.bar(x - w/2, pre,  w, label="Pre-test",  color="#4C78A8")
    bars2 = ax.bar(x + w/2, post, w, label="Post-test", color="#F58518")
    ax.set_ylabel("Mean Score (%)")
    ax.set_title("Figure 1 — Mean Pre- and Post-Test Scores by Group")
    ax.set_xticks(x)
    ax.set_xticklabels(groups)
    ax.set_ylim(50, 80)
    ax.legend()
    ax.bar_label(bars1, fmt="%.1f", padding=3, fontsize=8)
    ax.bar_label(bars2, fmt="%.1f", padding=3, fontsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def make_scatter_plot() -> bytes:
    """Weekly usage (hours) vs exam score — scatter."""
    rng = np.random.default_rng(42)
    hours = rng.uniform(0.5, 10, 200)
    scores = 45 + 3.2 * hours + rng.normal(0, 7, 200)
    scores = np.clip(scores, 20, 100)

    fig, ax = plt.subplots(figsize=(6, 3.5))
    ax.scatter(hours, scores, alpha=0.55, color="#4C78A8", s=25, edgecolors="none")
    # Regression line
    m, b = np.polyfit(hours, scores, 1)
    xline = np.linspace(0.5, 10, 100)
    ax.plot(xline, m * xline + b, color="#E45756", linewidth=1.8, label=f"r = .47, p < .001")
    ax.set_xlabel("Weekly Platform Usage (hours)")
    ax.set_ylabel("Final Examination Score (%)")
    ax.set_title("Figure 2 — Usage Time vs. Exam Score")
    ax.legend(fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def make_heatmap() -> bytes:
    """Correlation heatmap of key variables."""
    labels = ["Usage\n(hrs/wk)", "Prior\nGPA", "Feedback\nClicks", "Quiz\nAttempts", "Exam\nScore"]
    data = np.array([
        [1.00,  0.12,  0.68,  0.55,  0.47],
        [0.12,  1.00,  0.08,  0.05,  0.39],
        [0.68,  0.08,  1.00,  0.61,  0.43],
        [0.55,  0.05,  0.61,  1.00,  0.38],
        [0.47,  0.39,  0.43,  0.38,  1.00],
    ])
    fig, ax = plt.subplots(figsize=(5.5, 4))
    im = ax.imshow(data, cmap="RdYlGn", vmin=-1, vmax=1)
    ax.set_xticks(range(len(labels)))
    ax.set_yticks(range(len(labels)))
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_yticklabels(labels, fontsize=8)
    for i in range(len(labels)):
        for j in range(len(labels)):
            ax.text(j, i, f"{data[i,j]:.2f}", ha="center", va="center",
                    fontsize=8, color="black")
    plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    ax.set_title("Figure 3 — Correlation Matrix of Key Variables")
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# Build the document
# ---------------------------------------------------------------------------

doc = Document()

# ── Title ───────────────────────────────────────────────────────────────────
title = doc.add_heading("Impact of Digital Learning Tools on Academic Performance", 0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ── Authors / date ──────────────────────────────────────────────────────────
p = doc.add_paragraph("Marie Dupont, Jean-Paul Martin, Amara Diallo")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.runs[0].bold = True

p2 = doc.add_paragraph("Université de Lyon — March 2025")
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()

# ── Abstract ────────────────────────────────────────────────────────────────
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

# ── 1. Introduction ─────────────────────────────────────────────────────────
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

# ── 2. Theoretical Framework ─────────────────────────────────────────────────
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

# ── 3. Methodology ───────────────────────────────────────────────────────────
doc.add_heading("3. Methodology", level=1)
doc.add_paragraph(
    "We employed a quasi-experimental design with pre- and post-test measurements "
    "across two semesters (September 2023 – June 2024). Treatment groups used "
    "the Moodle-based adaptive quiz module for at least three hours per week; "
    "control groups followed traditional paper-based exercises."
)

# Image 1 — bar chart
doc.add_paragraph()
doc.add_picture(io.BytesIO(make_bar_chart()), width=Inches(5.0))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
cap1 = doc.add_paragraph(
    "Figure 1. Mean pre- and post-test scores by group (treatment vs. control)."
)
cap1.style = doc.styles["Caption"] if "Caption" in doc.styles else doc.styles["Normal"]
cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph(
    "Figure 1 displays the mean scores before and after intervention. The treatment "
    "group improved by 14.3 percentage points on average, compared with 5.7 points "
    "for the control group (Dupont & Martin, 2022)."
)

# ── 4. Results ───────────────────────────────────────────────────────────────
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
doc.add_picture(io.BytesIO(make_scatter_plot()), width=Inches(5.0))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
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
doc.add_picture(io.BytesIO(make_heatmap()), width=Inches(5.0))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
cap3 = doc.add_paragraph(
    "Figure 3. Correlation matrix: usage metrics, prior GPA, feedback clicks, quiz attempts, and exam score."
)
cap3.style = doc.styles["Caption"] if "Caption" in doc.styles else doc.styles["Normal"]
cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ── 5. Discussion ─────────────────────────────────────────────────────────────
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

# ── 6. Conclusion ─────────────────────────────────────────────────────────────
doc.add_heading("6. Conclusion", level=1)
doc.add_paragraph(
    "This study provides quasi-experimental evidence that adaptive digital quizzes "
    "improve undergraduate examination scores. The effect is strongest for "
    "lower-achieving students, pointing to equity benefits. Institutions should "
    "integrate formative-assessment platforms within broader self-regulation support "
    "programmes (Zimmerman, 2002; Dupont & Martin, 2022)."
)

# ── Sources / References (APA 7th edition) ────────────────────────────────────
doc.add_heading("Sources", level=1)

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
