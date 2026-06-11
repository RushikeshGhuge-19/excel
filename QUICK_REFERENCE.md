# Start Here

## What You Have
- `app.py` — FastAPI backend (upload → extract → Excel)
- `SpecExtractor.jsx` — React component (UI)
- `requirements.txt` — Python deps
- `package.json` — npm deps

## Get It Running

**Terminal 1 — Backend:**
```bash
pip install -r requirements.txt
export MISTRAL_API_KEY="sk_..."
python app.py
```

**Terminal 2 — Frontend:**
```bash
npx create-react-app spec-ocr
cd spec-ocr
npm install
# Copy SpecExtractor.jsx → src/SpecExtractor.jsx
# Update src/App.js:
#   import SpecExtractor from './SpecExtractor';
#   export default SpecExtractor;
npm start
```

Backend: `http://localhost:8000`
Frontend: `http://localhost:3000`

## Key Routes
- `POST /extract` — Upload image → Excel URL
- `GET /download/{filename}` — Download Excel
- `GET /health` — Check if running

## Two-Pass Logic
1. Extract all 7 size groups
2. Re-check right columns (5/6, 6/7, 7/8 YRS)
3. Merge confident values
4. Output Excel
