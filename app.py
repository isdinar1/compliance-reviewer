import streamlit as st
from google import genai as google_genai
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import json
import io
import re
import requests
from bs4 import BeautifulSoup

try:
    import pypdf
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

try:
    from docx import Document as DocxDocument
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Compliance Reviewer",
    page_icon="📋",
    layout="wide",
)

st.title("📋 Course Compliance Reviewer")
st.caption("Upload a course and a regulation/guidelines doc — get back a Word doc with every conflict and gap highlighted.")

# ── Gemini setup ─────────────────────────────────────────────────────────────
try:
    api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    api_key = None

if not api_key:
    st.error("Gemini API key not configured. Add GEMINI_API_KEY to your Streamlit secrets.")
    st.stop()

client = google_genai.Client(api_key=api_key)


# ── Text extraction helpers ───────────────────────────────────────────────────
def extract_text_from_file(f) -> str:
    name = f.name.lower()
    if name.endswith(".txt"):
        return f.read().decode("utf-8", errors="ignore")
    if name.endswith(".docx"):
        doc = Document(io.BytesIO(f.read()))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    if name.endswith(".pdf") and HAS_PYPDF:
        reader = pypdf.PdfReader(io.BytesIO(f.read()))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    return f.read().decode("utf-8", errors="ignore")


def extract_text_from_url(url: str) -> str:
    try:
        resp = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "header"]):
            tag.decompose()
        return soup.get_text(separator="\n", strip=True)[:40_000]
    except Exception as e:
        return f"[Could not fetch {url}: {e}]"


# ── Gemini review ─────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """You are a compliance analyst. You will be given:
1. A course script (training content)
2. One or more regulation/guideline documents

Your job: find every place where the course CONFLICTS with or is MISSING required content from the regulations.

Return a JSON object with this exact structure:
{
  "course_title": "inferred title or 'Untitled Course'",
  "summary": "2-3 sentence executive summary of overall compliance status",
  "findings": [
    {
      "id": 1,
      "type": "ISSUE",
      "section": "section or topic name from the course",
      "quote": "exact verbatim text from the course that is the problem (for ISSUE only, null for GAP)",
      "explanation": "clear explanation of the conflict or gap and what the regulation requires instead",
      "regulation_ref": "which regulation document and what it says"
    }
  ]
}

Rules:
- ISSUE = course text directly conflicts with or contradicts the regulation
- GAP = the regulation requires something that is completely absent from the course
- Be specific and cite exact quotes where possible
- Order findings by severity (most critical first)
- Return ONLY valid JSON, no markdown fences"""


def run_review(course_text: str, regulation_text: str) -> dict:
    prompt = f"""{SYSTEM_PROMPT}

--- COURSE SCRIPT ---
{course_text[:30000]}

--- REGULATIONS / GUIDELINES ---
{regulation_text[:30000]}"""

    response = client.models.generate_content(
        model="gemini-2.0-flash-lite",
        contents=prompt
    )
    raw = response.text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


# ── Word doc builder ──────────────────────────────────────────────────────────
def set_highlight(run, color: str):
    rpr = run._r.get_or_add_rPr()
    hl = OxmlElement("w:highlight")
    hl.set(qn("w:val"), color)
    rpr.append(hl)


def build_docx(course_text: str, result: dict) -> bytes:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # Header
    title_para = doc.add_paragraph()
    run = title_para.add_run(result.get("course_title", "Course Review"))
    run.bold = True
    run.font.size = Pt(14)

    sub = doc.add_paragraph()
    sub.add_run("Compliance Review").bold = True
    sub.runs[0].font.size = Pt(12)

    doc.add_paragraph(f"Total findings: {len(result.get('findings', []))}")
    doc.add_paragraph(result.get("summary", ""))

    # Legend
    legend = doc.add_paragraph()
    legend.add_run("Legend:  ")
    yr = legend.add_run("  ISSUE  ")
    set_highlight(yr, "yellow")
    legend.add_run("  = conflicts with regulation     ")
    cr = legend.add_run("  GAP  ")
    set_highlight(cr, "cyan")
    legend.add_run("  = content missing from course")

    doc.add_paragraph()

    # Build a quick lookup: quote → finding id
    findings = result.get("findings", [])
    issue_map = {
        f["quote"]: f
        for f in findings
        if f.get("type") == "ISSUE" and f.get("quote")
    }

    # Full course script with highlights
    heading = doc.add_paragraph()
    heading.add_run("FULL SCRIPT WITH ANNOTATIONS").bold = True

    gap_sections = {f.get("section", ""): f for f in findings if f.get("type") == "GAP"}

    for line in course_text.split("\n"):
        para = doc.add_paragraph()
        remaining = line

        # Check for ISSUE quotes in this line
        matched = False
        for quote, finding in issue_map.items():
            if quote in remaining:
                before, _, after = remaining.partition(quote)
                if before:
                    para.add_run(before)
                yr = para.add_run(quote)
                set_highlight(yr, "yellow")
                marker = para.add_run(f" [{finding['id']}]")
                marker.font.color.rgb = RGBColor(192, 0, 0)
                marker.bold = True
                if after:
                    para.add_run(after)
                matched = True
                break

        if not matched:
            para.add_run(remaining)

        # Insert GAP annotations after relevant sections
        for section, finding in list(gap_sections.items()):
            if section and section.lower() in line.lower():
                gap_para = doc.add_paragraph()
                gr = gap_para.add_run(
                    f"[{finding['id']}] GAP: {finding['explanation']}"
                )
                set_highlight(gr, "cyan")
                del gap_sections[section]

    # Insert any remaining GAPs at end
    for section, finding in gap_sections.items():
        gap_para = doc.add_paragraph()
        gr = gap_para.add_run(
            f"[{finding['id']}] GAP: {finding['explanation']}"
        )
        set_highlight(gr, "cyan")

    # Comment Legend page
    doc.add_page_break()
    cl = doc.add_paragraph()
    cl.add_run("Comment Legend").bold = True
    cl.runs[0].font.size = Pt(14)

    for f in findings:
        p = doc.add_paragraph()
        color = RGBColor(192, 0, 0) if f["type"] == "ISSUE" else RGBColor(0, 70, 127)
        label = f"[{f['id']}] {f['type']}: {f.get('section', '')} — {f['explanation']}"
        if f.get("regulation_ref"):
            label += f"\n    Regulation: {f['regulation_ref']}"
        r = p.add_run(label)
        r.font.color.rgb = color

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ── UI ────────────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("Course Content")
    course_sku = st.text_input("Course SKU / Title (optional)", placeholder="e.g. RVCT-882")
    course_file = st.file_uploader(
        "Upload course document", type=["docx", "pdf", "txt"], key="course"
    )
    course_text_input = st.text_area(
        "Or paste course text here", height=200, placeholder="Paste transcript or script..."
    )

with col2:
    st.subheader("Regulations / Guidelines")
    reg_files = st.file_uploader(
        "Upload regulation documents", type=["docx", "pdf", "txt"],
        accept_multiple_files=True, key="regs"
    )
    reg_urls = st.text_area(
        "And/or paste URLs to scan (one per line)",
        height=100,
        placeholder="https://cpr.heart.org/...\nhttps://..."
    )

st.divider()
run_btn = st.button("🔍 Run Compliance Review", type="primary", use_container_width=True)

if run_btn:
    # Gather course text
    course_text = ""
    if course_file:
        course_text = extract_text_from_file(course_file)
    elif course_text_input.strip():
        course_text = course_text_input.strip()

    if not course_text:
        st.error("Please upload a course document or paste the course text.")
        st.stop()

    if course_sku:
        course_text = f"Course SKU: {course_sku}\n\n{course_text}"

    # Gather regulation text
    reg_parts = []
    for rf in (reg_files or []):
        reg_parts.append(f"[Document: {rf.name}]\n{extract_text_from_file(rf)}")

    for url in (reg_urls or "").splitlines():
        url = url.strip()
        if url:
            with st.spinner(f"Fetching {url}..."):
                reg_parts.append(f"[URL: {url}]\n{extract_text_from_url(url)}")

    if not reg_parts:
        st.error("Please upload at least one regulation document or enter a URL.")
        st.stop()

    regulation_text = "\n\n".join(reg_parts)

    with st.spinner("Reviewing with AI — this takes 15–30 seconds..."):
        try:
            result = run_review(course_text, regulation_text)
        except Exception as e:
            st.error(f"Review failed: {e}")
            st.stop()

    findings = result.get("findings", [])
    issues = [f for f in findings if f["type"] == "ISSUE"]
    gaps = [f for f in findings if f["type"] == "GAP"]

    st.success(f"Found **{len(issues)} issues** and **{len(gaps)} gaps**")
    st.write(result.get("summary", ""))

    with st.expander("View all findings", expanded=True):
        for f in findings:
            icon = "🟡" if f["type"] == "ISSUE" else "🔵"
            st.markdown(f"**{icon} [{f['id']}] {f['type']} — {f.get('section', '')}**")
            if f.get("quote"):
                st.markdown(f"> {f['quote']}")
            st.markdown(f.get("explanation", ""))
            if f.get("regulation_ref"):
                st.caption(f"Regulation: {f['regulation_ref']}")
            st.divider()

    with st.spinner("Building Word document..."):
        docx_bytes = build_docx(course_text, result)

    title = result.get("course_title", course_sku or "Course").replace(" ", "_")
    st.download_button(
        label="⬇️ Download Reviewed Word Doc",
        data=docx_bytes,
        file_name=f"{title}_Compliance_Review.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )
