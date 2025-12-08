# MC Auto Grader (OMR Processor)

MC Auto Grader is a browser-based Optical Mark Recognition (OMR) tool for scanning, grading, and exporting results from multiple-choice answer sheets. It runs entirely in the client (no server upload), supports PDF multi-page processing and common image formats, and generates an Excel workbook with live formulas for analysis.

---

## Key Features

- Dark-mode UI optimized for scanning overlays.
- Supports PDF (multi-page) and image (.jpg, .png) uploads.
- Robust ID detection using a Candidate System to handle light marks and smudges.
- Strict answer-key parsing (accepts A–D only; supports formats like `1.A`, `2-B`, raw `ABCD`).
- Flexible weighting per question ranges (batch weights).
- Manual alignment with draggable template boxes + auto-alignment heuristics.
- Excel export (.xlsx) with three sheets and dynamic formulas (All Results, Scores, Statistics).
- All processing is performed locally in the browser.

---


## Usage

1. Upload sheets
   - Click the upload area or press Ctrl+O.
   - Accepts PDF and image files. PDF pages are rendered and loaded as separate pages.

2. Set answer key
   - Paste your key into the Answer Key textarea. Supported formats:
     - Single-line block: `ABCDABCD...`
     - Tokenized: `1.A 2.B 3.C`
   - The parser extracts only A–D characters and maps them to consecutive question numbers.

3. Configure weights
   - Open "Question Weighting" to add ranges (Start — End : Mark).
   - Default: Q1–60 = 1 mark.

4. Align template
   - The app auto-detects approximate X/Y offsets. Use draggable boxes to fine-tune.
   - Click "Realign Boxes" to reset to auto-detection.

5. Scan answers
   - Click "Scan Answers" or press Ctrl+Enter.
   - Progress shown per page. Results panel slides up with per-question previews.

6. Export
   - Click "Export Results" or press Ctrl+E to download an .xlsx file. Requires SheetJS (loaded via CDN).

---

## Excel Output Structure

- Sheet 1 — All Results
  - Rows: one student per row (Student ID derived from ID blocks or page index).
  - Columns: Q1, Q2, ... each cell contains A/B/C/D/BLANK/MULT.
  - Footer: Marks (per-question), Answer Key, Average / Percentage, Distribution (A/B/C/D).

- Sheet 2 — Scores
  - Student (refs Sheet1), Score (SUMPRODUCT dynamic formula comparing answers to the key and weights), Percentage (score / full mark).

- Sheet 3 — Statistics
  - Full mark (derived from weighting), Max, Min, Mean, Median, StdDev, Passing analysis (50% and 40%).

---

## Algorithm & Implementation Notes

Files: core UI and logic are implemented in [src/App.jsx](src/App.jsx). See functions: [`runBatchDetection`](src/App.jsx), [`exportExcel`](src/App.jsx), [`getQuestionMark`](src/App.jsx), [`processPdf`](src/App.jsx), [`processImage`](src/App.jsx), [`calculateRowLayout`](src/App.jsx), [`detectVerticalOffset`](src/App.jsx), [`detectHorizontalOffset`](src/App.jsx), [`getStandardRegions`](src/App.jsx).

ID detection (Candidate System)
- Each ID bubble scans the central 80% (padding 10%) to catch edge marks.
- A bubble becomes a candidate if its filled ratio meets a minimum threshold.
- 0 candidates → BLANK; 1 → selected; 2+ → compare top two: if winner lead > 5% → winner, else MULT.

Answer detection
- For each bubble compute fill ratio inside a reduced padding.
- If max fill is low or close to others it returns BLANK or MULT; otherwise picks the darkest bubble as answer.
- Row layout accounts for gap rows (e.g., every 6th row) to match printed forms (see [`calculateRowLayout`](src/App.jsx)).

Image & PDF handling
- Images are read into a Canvas; PDFs are rendered via pdf.js (loaded from CDN).
- Excel export uses SheetJS (XLSX) from CDN to build workbook and write dynamic formulas.

---

## Limitations & Recommendations

- All work is client-side. Very large PDFs (100+ pages) may consume significant memory.
- Heavily skewed or rotated scans may require pre-processing (rotate/crop) or manual alignment in-app.
- Always review exported results — automated detection can make mistakes.

---

## Development

Primary files:
- App & logic: [src/App.jsx](src/App.jsx)
- Entry: [src/main.jsx](src/main.jsx)
- HTML template: [index.html](index.html)
- Project config: [package.json](package.json), [vite.config.js](vite.config.js)

Run linting:
```sh
npm run lint
```

---

## License

MIT — see [LICENSE](LICENSE).

---

Thank you for using MC Auto Grader. Validate automated results manually before finalizing grades.