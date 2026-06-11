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

API_KEY = os.getenv("MISTRAL_API_KEY", "o5RwmxtgR18zW6ppUffmrSrEPRdEay39")
MODEL = "pixtral-large-2411"
MAX_TOKENS = 16000
TIMEOUT_SEC = 300
MAX_RETRIES = 3

OUTPUT_DIR = Path("./outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

# ── Helpers ──────────────────────────────────────────────────────────────────

def preprocess_image(image_bytes: bytes) -> tuple:
    """
    Open image from bytes, boost contrast + sharpness, encode as PNG base64.
    Returns (base64_str, mime_type, width, height).
    """
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    w, h = img.size

    # Upscale if smaller than 1800px on longest side
    longest = max(w, h)
    if longest < 1800:
        scale = 1800 / longest
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
        w, h = img.size

    # Contrast + sharpness boost
    img = ImageEnhance.Contrast(img).enhance(1.4)
    img = ImageEnhance.Sharpness(img).enhance(1.5)

    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=False)
    buf.seek(0)
    b64 = base64.b64encode(buf.getvalue()).decode()

    return b64, "image/png", w, h

def call_api(b64: str, mime: str, prompt: str) -> dict:
    """Call Mistral API with image + prompt. Returns parsed JSON."""
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
                return {"error": f"HTTP {r.status_code}: {r.text[:200]}"}

            rj = r.json()
            choice = rj["choices"][0]
            finish = choice.get("finish_reason", "?")

            raw = choice["message"]["content"].strip()
            raw = re.sub(r'^```(?:json)?\s*', '', raw)
            raw = re.sub(r'\s*```$', '', raw).strip()

            if finish == "length":
                raw = _repair(raw)

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
    return {"error": f"JSON parse failed", "raw": raw[:300]}

def _repair(raw: str) -> str:
    """Repair truncated JSON."""
    last = raw.rfind('],')
    if last != -1:
        raw = raw[:last + 1]
    raw += ']' * max(raw.count('[') - raw.count(']'), 0)
    raw += '}' * max(raw.count('{') - raw.count('}'), 0)
    return raw

def normalize_slashes(result: dict) -> dict:
    """Fix #7: Replace any lone '-' or '' in ADJ/S1 columns with '/'."""
    fixed = 0
    for row in result.get("rows", []):
        if not isinstance(row, list):
            continue
        for idx in range(3, len(row)):
            if (idx - 2) % 3 == 0:
                continue  # skip BASE
            v = str(row[idx]).strip()
            if v in ("-", "", "—", "–"):
                row[idx] = "/"
                fixed += 1
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

    merged = 0
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
            if idx < len(row):
                if v2 not in ("/", "", None):
                    row[idx] = v2
                    merged += 1

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
Col 2: Tolerance — two stacked cells show "Pos(+) X" and "Neg(-) X". Output as "+-X" e.g. "+-1", "+-0.5"
Then for each size group (read ALL from header): 3 sub-columns → BASE | ADJ | S1

Each measurement has 2 physical rows in the image (Pos on top, Neg below).
Output ONLY ONE row per measurement.

=== CELL MERGE RULE (strict priority) ===

For each ADJ and S1 cell, look at BOTH physical rows (Pos top + Neg bottom):
  1. Pos row has a number  AND  Neg row has "/"  → use the Pos number
  2. Neg row has a number  AND  Pos row has "/"  → use the Neg number
  3. BOTH rows have numbers                       → use the Pos row value
  4. BOTH rows are "/" or blank or checkmark      → output "/"

BASE column: printed number (same in both rows). Output once.
A checkmark (✓) always counts as "/".

=== BLANK CELL RULE (critical) ===

Any ADJ or S1 cell that is empty, contains only a checkmark (✓), or contains a forward slash (/)
MUST be output as exactly "/" — never output "-", "", or null.
There is NO dash character "-" used alone as a value in this table.
If you are tempted to write "-" for a cell, write "/" instead.

=== BASE VALUE VERIFICATION (critical — prevents column drift) ===

BASE values are pre-printed numbers that ALWAYS increase left to right across size groups.
Before writing each BASE value:
  - Is it LARGER than the previous size group's BASE for this same measurement?
  - Does it match the printed number visible in THIS column (not the column to the left)?

COLUMN DRIFT WARNING: The most common error is copying the previous size group's BASE
instead of reading the new one, which shifts all values left by one column.

=== HANDWRITTEN NUMBER DECODING ===

Numbers use shorthand — NO decimal point is written:
  -1=-1   +1=+1   -05=-0.5  +05=+0.5  -03=-0.3  +03=+0.3
  -02=-0.2  +02=+0.2  -04=-0.4  +04=+0.4  -06=-0.6  +06=+0.6
  -07=-0.7  +07=+0.7  -15=-1.5  +15=+1.5  -2=-2  +2=+2

SIGN RULE: The sign written to the LEFT of digits ALWAYS belongs to that number.

=== CIRCLED NUMBERS (critical — sign is OUTSIDE the circle) ===

A digit inside a drawn circle/oval IS a real value, NOT decoration.
The sign (+/-) is written OUTSIDE the circle. Look for it BEFORE assigning a sign.

  Circle "1" + sign "+" written outside or no sign → +1
  Circle "1" + sign "-" written outside → -1
  NEVER output "/" for a circled number.
  NEVER assign negative without seeing a "-" explicitly written outside the circle.

=== UNCERTAINTY RULE ===

If you are not sure whether an ADJ or S1 cell contains a value or is blank,
output "/" — do NOT guess a value.

=== RIGHT COLUMNS — EXTRA ATTENTION ===

The last 2–3 size groups (5/6 YRS, 6/7 YRS, 7/8 YRS) are near the right edge.
Apply the BASE VERIFICATION rule especially carefully here.

=== OUTPUT FORMAT ===

Return strict JSON only. No markdown. No explanation.

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

BLANK CELL RULE: empty / checkmark / slash → ALWAYS output "/" — never "-".

HANDWRITTEN shorthand:
  -05=-0.5  +05=+0.5  -03=-0.3  +03=+0.3  -02=-0.2  +02=+0.2
  -07=-0.7  +07=+0.7  -15=-1.5  +15=+1.5

CIRCLED NUMBERS: sign is OUTSIDE circle. No sign or "+" → positive. "-" → negative.
NEVER output "/" for a circled number.

BASE VERIFICATION: BASE values increase left to right (5/6 > 4/5, 6/7 > 5/6, 7/8 > 6/7).

UNCERTAINTY RULE: when unsure about ADJ or S1 → output "/" not a guess.

=== OUTPUT FORMAT ===

Return strict JSON only. No markdown.

{
  "rows": [
    {"desc": "A6 Length - SNP to Hem at Front", "vals": ["40", "-1", "+0.5", "40.5", "-1", "/", "41.5", "-0.5", "-1"]},
    ...
  ]
}

Extract ALL measurement rows. Output ONLY 9 values per row (5/6, 6/7, 7/8 YRS)."""

# ── Routes ───────────────────────────────────────────────────────────────────

@app.get("/health")
async def health():
    return {"status": "ok"}

@app.post("/extract")
async def extract(file: UploadFile = File(...)):
    """
    Upload image → extract garment specs → return Excel file URL.
    """
    if not file:
        raise HTTPException(status_code=400, detail="No file uploaded")

    if not API_KEY.strip():
        raise HTTPException(status_code=500, detail="API key not configured")

    # Read file
    try:
        image_bytes = await file.read()
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"File read error: {str(e)}")

    # Preprocess
    try:
        b64, mime, w, h = preprocess_image(image_bytes)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Image preprocess error: {str(e)}")

    # Pass 1: Full table
    result1 = call_api(b64, mime, PROMPT_FULL)
    if result1.get("error"):
        raise HTTPException(status_code=500, detail=f"Pass-1 failed: {result1['error']}")

    if not result1.get("rows"):
        raise HTTPException(status_code=500, detail="Pass-1: no data extracted")

    result1 = normalize_slashes(result1)

    # Pass 2: Right columns re-check
    result2 = call_api(b64, mime, PROMPT_RIGHTCOLS)
    if result2.get("rows"):
        result1 = merge_passes(result1, result2)
        result1 = normalize_slashes(result1)

    # Generate Excel
    try:
        excel_bytes = write_excel(result1)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Excel generation error: {str(e)}")

    # Save Excel file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"spec_{timestamp}.xlsx"
    filepath = OUTPUT_DIR / filename

    try:
        with open(filepath, "wb") as f:
            f.write(excel_bytes)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"File save error: {str(e)}")

    return {
        "status": "success",
        "message": "Extraction complete",
        "title": result1.get("title", "Specification"),
        "size_groups": result1.get("size_groups", []),
        "rows_extracted": len(result1.get("rows", [])),
        "excel_filename": filename,
        "excel_url": f"/download/{filename}",
        "data": result1
    }

@app.get("/download/{filename}")
async def download(filename: str):
    """Download extracted Excel file."""
    filepath = OUTPUT_DIR / filename
    if not filepath.exists():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )

@app.get("/")
async def root():
    return {
        "name": "Garment Spec OCR Extractor",
        "version": "3.0",
        "endpoints": {
            "POST /extract": "Upload image and extract specs",
            "GET /download/{filename}": "Download Excel file"
        }
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
