"""
Garment Spec Sheet OCR Extractor — Streamlit App
Upload image → two-pass OCR extraction → download Excel
"""

import streamlit as st
import os, re, json, io, time, base64
from pathlib import Path
from datetime import datetime

import requests
from PIL import Image, ImageEnhance
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Config ───────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Garment Spec OCR",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

API_KEY = os.getenv("MISTRAL_API_KEY", "")
MODEL = "pixtral-large-2411"
MAX_TOKENS = 16000
TIMEOUT_SEC = 300
MAX_RETRIES = 3

OUTPUT_DIR = Path("./outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

# ── Helpers ──────────────────────────────────────────────────────────────────

def preprocess_image(image_bytes: bytes) -> tuple:
    """Boost contrast + sharpness, encode as PNG base64."""
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    w, h = img.size

    longest = max(w, h)
    if longest < 1800:
        scale = 1800 / longest
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
        w, h = img.size

    img = ImageEnhance.Contrast(img).enhance(1.4)
    img = ImageEnhance.Sharpness(img).enhance(1.5)

    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=False)
    buf.seek(0)
    b64 = base64.b64encode(buf.getvalue()).decode()

    return b64, "image/png", w, h

def call_api(b64: str, mime: str, prompt: str) -> dict:
    """Call Mistral API."""
    if not API_KEY.strip():
        return {"error": "API key not configured"}

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": MODEL,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                {"type": "text", "text": prompt}
            ]
        }],
        "max_tokens": MAX_TOKENS,
        "temperature": 0.1,
    }

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = requests.post(
                "https://api.mistral.ai/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=TIMEOUT_SEC
            )

            if r.status_code == 401:
                return {"error": "Invalid API key"}
            if r.status_code == 429:
                time.sleep(30)
                continue
            if r.status_code != 200:
                return {"error": f"HTTP {r.status_code}"}

            rj = r.json()
            choice = rj["choices"][0]
            raw = choice["message"]["content"].strip()
            raw = re.sub(r'^```(?:json)?\s*', '', raw)
            raw = re.sub(r'\s*```$', '', raw).strip()

            return _parse(raw)

        except requests.exceptions.ReadTimeout:
            if attempt < MAX_RETRIES:
                time.sleep(15 * attempt)
            else:
                return {"error": "Request timeout"}
        except requests.exceptions.ConnectionError:
            if attempt < MAX_RETRIES:
                time.sleep(10)
            else:
                return {"error": "Connection failed"}
    
    return {"error": "Max retries exceeded"}

def _parse(raw: str) -> dict:
    """Parse JSON response."""
    for s in (raw, _repair(raw)):
        try:
            d = json.loads(s)
            if isinstance(d, dict) and "rows" in d:
                return d
        except:
            pass
    return {"error": f"JSON parse failed"}

def _repair(raw: str) -> str:
    """Repair truncated JSON."""
    last = raw.rfind('],')
    if last != -1:
        raw = raw[:last + 1]
    raw += ']' * max(raw.count('[') - raw.count(']'), 0)
    raw += '}' * max(raw.count('{') - raw.count('}'), 0)
    return raw

def normalize_slashes(result: dict) -> dict:
    """Fix: Replace '-' or '' in ADJ/S1 columns with '/'."""
    for row in result.get("rows", []):
        if not isinstance(row, list):
            continue
        for idx in range(3, len(row)):
            if (idx - 2) % 3 == 0:
                continue
            v = str(row[idx]).strip()
            if v in ("-", "", "—", "–"):
                row[idx] = "/"
    return result

def merge_passes(result1: dict, result2: dict) -> dict:
    """Merge pass-2 (right cols) into pass-1."""
    if not result2 or "rows" not in result2:
        return result1

    p2_map = {}
    for entry in result2["rows"]:
        if isinstance(entry, dict):
            desc = entry.get("desc", "").strip().lower()
            vals = entry.get("vals", [])
            if desc and len(vals) == 9:
                p2_map[desc] = vals

    rows = result1.get("rows", [])
    for row in rows:
        if not isinstance(row, list) or len(row) < 23:
            continue
        desc_key = str(row[0]).strip().lower()
        p2 = p2_map.get(desc_key)
        
        if not p2:
            for k, v in p2_map.items():
                if k[:15] in desc_key or desc_key[:15] in k:
                    p2 = v
                    break
        if not p2:
            continue

        for i, v2 in enumerate(p2):
            idx = 14 + i
            if idx < len(row) and v2 not in ("/", "", None):
                row[idx] = v2

    return result1

def write_excel(data: dict) -> bytes:
    """Generate Excel file. Returns bytes."""
    rows = data.get("rows", [])
    title = data.get("title", "Specification")
    size_groups = data.get("size_groups", [])
    n_groups = len(size_groups)
    n_cols = 2 + n_groups * 3

    wb = Workbook()
    ws = wb.active
    ws.title = "Specifications"

    # Row 1: Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(name='Arial', bold=True, size=12)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = Border(
        top=Side(style='thin'),
        bottom=Side(style='thin'),
        left=Side(style='thin'),
        right=Side(style='thin')
    )
    ws.row_dimensions[1].height = 22

    # Row 2: Size group headers
    HDR = "D9D9D9"
    ws.merge_cells(start_row=2, start_column=1, end_row=3, end_column=1)
    ws.merge_cells(start_row=2, start_column=2, end_row=3, end_column=2)

    for col in [1, 2]:
        c = ws.cell(row=2, column=col)
        c.font = Font(name='Arial', bold=True, size=8)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = Border(top=Side(style='thin'), bottom=Side(style='thin'),
                         left=Side(style='thin'), right=Side(style='thin'))
        c.fill = PatternFill('solid', start_color=HDR, end_color=HDR)

    ws.cell(row=2, column=1).value = "Description"
    ws.cell(row=2, column=2).value = "Tolerance"

    for gi, grp in enumerate(size_groups):
        bc = 3 + gi * 3
        ws.merge_cells(start_row=2, start_column=bc, end_row=2, end_column=bc + 2)
        c = ws.cell(row=2, column=bc)
        c.value = grp
        c.font = Font(name='Arial', bold=True, size=8)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.fill = PatternFill('solid', start_color=HDR, end_color=HDR)
        c.border = Border(top=Side(style='thin'), bottom=Side(style='thin'),
                         left=Side(style='thin'), right=Side(style='thin'))

    ws.row_dimensions[2].height = 22

    # Row 3: Sub-headers
    for gi in range(n_groups):
        bc = 3 + gi * 3
        for offset, lbl in enumerate(["BASE", "ADJ", "S1"]):
            c = ws.cell(row=3, column=bc + offset, value=lbl)
            c.font = Font(name='Arial', bold=True, size=7)
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.fill = PatternFill('solid', start_color=HDR, end_color=HDR)
            c.border = Border(top=Side(style='thin'), bottom=Side(style='thin'),
                             left=Side(style='thin'), right=Side(style='thin'))

    ws.row_dimensions[3].height = 14

    # Data rows
    BASE_BG = "EBEBEB"
    for ri, row in enumerate(rows):
        er = ri + 4
        bg = "FFFFFF" if ri % 2 == 0 else "F2F2F2"
        rd = list(row) if isinstance(row, list) else []

        for ci in range(1, n_cols + 1):
            val = rd[ci - 1] if ci - 1 < len(rd) else ""
            cell = ws.cell(row=er, column=ci)
            cell.value = str(val) if val not in (None, "") else ""
            cell.font = Font(name='Arial', size=8 if ci == 1 else 9)
            cell.alignment = Alignment(horizontal='left' if ci == 1 else 'center', vertical='center')
            cell.border = Border(top=Side(style='thin'), bottom=Side(style='thin'),
                                left=Side(style='thin'), right=Side(style='thin'))

            if ci > 2:
                col_within = (ci - 3) % 3
                if col_within == 0:
                    cell.font = Font(name='Arial', bold=True, size=8)
                    cell.fill = PatternFill('solid', start_color=BASE_BG, end_color=BASE_BG)
                else:
                    cell.fill = PatternFill('solid', start_color=bg, end_color=bg)
            else:
                cell.fill = PatternFill('solid', start_color=bg, end_color=bg)

        ws.row_dimensions[er].height = 16

    # Column widths
    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 7
    for gi in range(n_groups):
        bc = 3 + gi * 3
        ws.column_dimensions[get_column_letter(bc)].width = 7
        ws.column_dimensions[get_column_letter(bc + 1)].width = 7
        ws.column_dimensions[get_column_letter(bc + 2)].width = 7

    ws.freeze_panes = ws.cell(row=4, column=3)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── Prompts ──────────────────────────────────────────────────────────────────

PROMPT_FULL = """You are extracting data from a garment spec sheet photo. Output ONE row per measurement.

=== TABLE STRUCTURE ===
Col 1: Description
Col 2: Tolerance "+-X" e.g. "+-1", "+-0.5"
Then for each size group: 3 sub-columns → BASE | ADJ | S1

=== CELL MERGE RULE ===
For each ADJ and S1 cell, look at BOTH physical rows (Pos top + Neg bottom):
  1. Pos row has a number AND Neg row has "/" → use Pos number
  2. Neg row has a number AND Pos row has "/" → use Neg number
  3. BOTH rows have numbers → use Pos row value
  4. BOTH rows are "/" or blank → output "/"

=== BLANK CELL RULE ===
Any empty cell or checkmark MUST be "/" — never "-" or "".

=== BASE VALUE VERIFICATION ===
BASE values increase left to right. If a BASE looks same as previous, re-read column.

=== HANDWRITTEN NUMBER DECODING ===
-05=-0.5  +05=+0.5  -03=-0.3  +03=+0.3  -02=-0.2  +02=+0.2
-07=-0.7  +07=+0.7  -15=-1.5  +15=+1.5  -1=-1  +1=+1

=== CIRCLED NUMBERS ===
A digit inside a circle IS a real value. Sign (+/-) is OUTSIDE circle.
Circle "1" + no sign or "+" → +1
Circle "1" + "-" sign → -1
NEVER output "/" for a circled number.

=== UNCERTAINTY RULE ===
If unsure about ADJ or S1, output "/" not a guess.

=== OUTPUT FORMAT ===
Return strict JSON only. No markdown.

{
  "title": "<title from image>",
  "size_groups": ["1.5/2 YRS", "2/3 YRS", "3/4 YRS", "4/5 YRS", "5/6 YRS", "6/7 YRS", "7/8 YRS"],
  "rows": [
    ["A6 Length - SNP to Hem at Front", "+-1", "33.5", "-1", "-0.5", "35", "-0.5", "/", ...],
    ...
  ]
}

Each row = 2 + (num_size_groups × 3) values. Extract ALL measurements."""


PROMPT_RIGHTCOLS = """You are re-checking the RIGHT SIDE of a garment spec sheet.
Focus ONLY on the last 3 size groups: 5/6 YRS, 6/7 YRS, 7/8 YRS.

=== RULES ===

CELL MERGE RULE:
  1. Pos number + Neg "/" → use Pos
  2. Neg number + Pos "/" → use Neg
  3. Both numbers → use Pos
  4. Both "/" or blank → output "/"

BLANK CELL RULE: empty / checkmark / slash → ALWAYS "/" — never "-".

HANDWRITTEN shorthand:
  -05=-0.5  +05=+0.5  -03=-0.3  +03=+0.3  -02=-0.2  +02=+0.2
  -07=-0.7  +07=+0.7  -15=-1.5  +15=+1.5

CIRCLED NUMBERS: sign is OUTSIDE circle. No sign or "+" → positive. "-" → negative.

BASE VERIFICATION: 5/6 > 4/5, 6/7 > 5/6, 7/8 > 6/7.

UNCERTAINTY: when unsure → "/" not a guess.

=== OUTPUT FORMAT ===

{
  "rows": [
    {"desc": "A6 Length - SNP to Hem at Front", "vals": ["40", "-1", "+0.5", "40.5", "-1", "/", "41.5", "-0.5", "-1"]},
    ...
  ]
}

Extract ALL measurement rows. Output ONLY 9 values (5/6, 6/7, 7/8 YRS)."""

# ── Streamlit UI ────────────────────────────────────────────────────────────

st.markdown("""
<style>
    .main-title {
        font-size: 2.5rem;
        font-weight: 900;
        color: #1e293b;
        margin-bottom: 0.5rem;
    }
    .subtitle {
        font-size: 0.95rem;
        color: #64748b;
        margin-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([10, 1])
with col1:
    st.markdown('<p class="main-title">📊 Garment Spec OCR</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Upload photo → Extract measurements → Download Excel</p>', unsafe_allow_html=True)

if not API_KEY.strip():
    st.error("❌ API key not configured. Set `MISTRAL_API_KEY` environment variable.")
    st.stop()

# Upload section
uploaded_file = st.file_uploader(
    "Upload garment spec photo",
    type=["jpg", "jpeg", "png", "bmp", "webp", "tiff"],
    help="Clear photo of measurement table"
)

if uploaded_file:
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.success(f"✓ {uploaded_file.name} ({uploaded_file.size // 1024} KB)")
    
    with col2:
        extract_btn = st.button("🚀 Extract", use_container_width=True, type="primary")
    
    if extract_btn:
        try:
            # Progress tracking
            with st.spinner("📸 Preprocessing image..."):
                image_bytes = uploaded_file.getvalue()
                b64, mime, w, h = preprocess_image(image_bytes)
            
            # Pass 1
            with st.spinner("🔍 Pass 1: Extracting full table..."):
                result1 = call_api(b64, mime, PROMPT_FULL)
                
                if result1.get("error"):
                    st.error(f"Pass-1 failed: {result1['error']}")
                    st.stop()
                
                if not result1.get("rows"):
                    st.error("No data extracted. Try clearer photo.")
                    st.stop()
                
                result1 = normalize_slashes(result1)
            
            # Pass 2
            with st.spinner("🔍 Pass 2: Re-checking right columns..."):
                result2 = call_api(b64, mime, PROMPT_RIGHTCOLS)
                if result2.get("rows"):
                    result1 = merge_passes(result1, result2)
                    result1 = normalize_slashes(result1)
            
            # Excel
            with st.spinner("📝 Generating Excel..."):
                excel_bytes = write_excel(result1)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"spec_{timestamp}.xlsx"
            
            # Success
            st.success("✅ Extraction complete!")
            
            # Results display
            st.divider()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Title", result1.get("title", "Specification")[:20] + "...")
            with col2:
                st.metric("Size Groups", len(result1.get("size_groups", [])))
            with col3:
                st.metric("Measurements", len(result1.get("rows", [])))
            
            st.divider()
            
            # Size groups
            st.subheader("Size Groups")
            st.write(" • ".join(result1.get("size_groups", [])))
            
            # Sample rows
            st.subheader("Sample Measurements")
            sample_data = []
            for row in result1.get("rows", [])[:5]:
                if isinstance(row, list) and len(row) > 1:
                    sample_data.append({"Description": row[0], "Tolerance": row[1]})
            
            if sample_data:
                st.dataframe(sample_data, use_container_width=True, hide_index=True)
                remaining = len(result1.get("rows", [])) - 5
                if remaining > 0:
                    st.caption(f"... +{remaining} more measurements in Excel")
            
            st.divider()
            
            # Download button
            st.download_button(
                label="⬇️  Download Excel",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
else:
    st.info("👆 Upload an image to get started")
