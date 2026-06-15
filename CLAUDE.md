# 3mensio PDF to Excel Tracker — Build Spec & Session Log

## Project Summary

Single-file browser tool:

```text
3mensio_pdf_to_excel.html
```

Accepts multiple 3mensio Structural Heart PDF reports, extracts data in the browser, previews extracted rows, and downloads an Excel workbook matching `Sample Tracker.xlsx` (columns A–CI, 87 columns).

The page has three views, switched via tab-style buttons in the button row:

1. **Main (Text Entries) view** — default. PDF upload zone, Process/Download/Clear buttons, counters, preview table, debug panel.
2. **Snap-Shot based 3mensio Entries view** — helper page for the snapshot/OCR workflow handled externally (see `3. 3mensio Tracker - Manual Entries\CLAUDE.md`).
3. **Excel Merger view** — merges the individually downloaded batch workbooks into one.

See "UI Views & Tabs" below for full behavior.

---

## Single File Rule

Only one output file: `3mensio_pdf_to_excel.html`

All CSS inside `<style>`. All JavaScript inside `<script>`. CDN libraries only.

```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
```

**Important:** Set `pdfjsLib.GlobalWorkerOptions.workerSrc = ''` to disable the web worker.
This is required because the file runs via `file://` protocol and CDN workers are blocked by CORS.

---

## Filename → Excel Row Rule

```javascript
function getTargetRowNumber(fileName) {
  const match = fileName.match(/^(\d+)\./);
  return match ? Number(match[1]) : null;
}
```

- `3.report.pdf` → row 3
- `27.report.pdf` → row 27
- Rows 1–2 are headers. Data starts at row 3.
- Blank rows are preserved between non-contiguous row numbers.
- Duplicate row numbers block download until resolved.

---

## Excel Output

- Sheet name: `Sample Tracker`
- Two-row grouped header matching `Sample Tracker.xlsx`
- 87 columns A through CI
- Empty/missing fields: output `NA` (not blank)

### Batch Download Filename

Each Download Excel produces a batch-numbered file named:

```text
<batch#>. <minRow> - TO - <maxRow>.xlsx
```

Examples: `1. 11 - TO - 15.xlsx`, `2. 18 - TO - 20.xlsx`

- `:` is not allowed in Windows filenames, so `.` follows the batch number.
- The batch counter persists across sessions in `localStorage` (`batchCounter`).
- Re-downloading the **same row range** reuses the same batch number
  (`batchLastRange` in `localStorage`); a new range increments the counter.
- A small borderless `↺` icon button next to Clear resets the counter —
  the next download becomes batch 1 again. Tooltip explains this; a green
  status message confirms the reset.

---

## UI Views & Tabs

The button row holds two tab-style buttons (`.view-tab`): **📷 Snap-Shot based 3mensio Entries**
and **📑 Excel Merger**. Rules:

- A tab only **opens** its view. Clicking the active tab again does nothing.
- Returning to the main view is done via a **↩ Go Back** button that appears
  (absolutely positioned right of the centered tab) only while a view is open.
- While a view is open: the main-view content (`#viewTop`, `#viewBottom`),
  Process/Download/Clear buttons, the `↺` reset icon, and the *other* tab are hidden,
  and the button row centers the active tab (`.snapshot-active` class on `.btn-row`).
- View switching is handled by `setSnapshotView(active)` / `setMergerView(active)`.

### Snap-Shot based 3mensio Entries view

Helper page for the snapshot/OCR workflow (performed externally by an AI tool with
folder access). Contains:

1. **⬇ Download Markdown File** button — appears in the button row, horizontally
   aligned with the tab, centered over the left column. Downloads the focused-spec
   `CLAUDE.md` (from `3. 3mensio Tracker - Manual Entries`) whose full text is
   **embedded at build time** in the `SNAPSHOT_MD` template literal (backticks,
   backslashes, `${` and `</script>` are escaped). Download works offline via a Blob;
   saved as `CLAUDE.md`. If the source spec changes, the embedded copy must be
   re-injected.
2. **"Instruction to use" box** (left, `.prompt-box`, centered header) — 7-step
   numbered guide (download the .md into an empty folder, copy numbered reports
   there, provide folder path, copy prompt, paste into a tool with folder access,
   entries are made sequentially, merge with text-based entries later), ending with
   a green centered line "Entire 3mensio Tracker Entries is Complete !!".
3. **"Initial Prompt to process PDF one by one." box** (right, `.prompt-box`) —
   a text input (`#workLocationInput`, placeholder `[Enter your working location]`)
   followed by the fixed instruction lines (make the excel file first with column
   headers, then process PDFs sequentially starting from "3.").
4. **📋 Copy to Clip Board** button centered below the right box — copies the typed
   working location (or the bracketed placeholder if empty) plus the instruction
   lines to the clipboard. Uses `navigator.clipboard.writeText` with a hidden
   `textarea` + `document.execCommand('copy')` fallback for `file://`; button text
   briefly flips to "✓ Copied!".

### Excel Merger view

Merges the individually downloaded batch workbooks into one.

- Upload zone styled like the PDF zone; accepts only `.xlsx` (click to browse or
  drag & drop). Selected files are listed below with a `✕` remove control;
  duplicate filenames are skipped.
- **🔀 Merge & Download** button enables once ≥ 2 files are selected.
- Merge logic (`doMerge`): each batch file already carries its data at absolute
  row positions (rows 1–2 headers, data from row 3). Every non-blank data row is
  placed at its original Excel row number; blank filler rows are ignored; row
  numbers missing from all files stay blank in the output.
- **Overlap check:** if the same row number holds data in two files, the merge is
  blocked and a red error (`.status-bar.error` in `#mergerStatus`) names the row
  and both conflicting files. Nothing downloads until resolved.
- Output uses the canonical `HEADER_ROW1`/`HEADER_ROW2`, `MERGE_RANGES`, and column
  widths; sheet name `Sample Tracker`; downloaded as
  `Merged <minRow> - TO - <maxRow>.xlsx` with a green success message showing the
  file/row counts.

---

## Confirmed Column Extractions (session-by-session)

### A — Case entry date
Auto-populated with **today's date** in `DD-MM-YYYY` format.
```javascript
const d = new Date();
row[COL.CASE_ENTRY_DATE] = `${String(d.getDate()).padStart(2,'0')}-${String(d.getMonth()+1).padStart(2,'0')}-${d.getFullYear()}`;
```

### B — Case entered by
Fixed value: `Hiren`

### C — Case done by
Extracted from PDF `Created By:` field.
Uses dedicated `extractCreatedBy(text)` function that:
1. Finds `Created By:` (case-insensitive)
2. Takes the rest of that line
3. Strips trailing Title Case label (e.g. `Hospital:`) using case-sensitive regex
   (ALL-CAPS words like `MERIL`, `TAVI`, `CORELAB` are NOT treated as labels)

```javascript
function extractCreatedBy(text) {
  const m = text.match(/Created\s*By\s*:\s*([^\n\r]+)/i);
  if (!m) return '';
  let val = m[1].trim();
  val = val.replace(/\s+[A-Z][a-z][A-Za-z]*(?:\s+[A-Z][a-z][A-Za-z]*)*\s*:.*$/, '');
  return val.trim();
}
```

Example: `MERIL TAVI CORELAB (AG)`

### D — Patient ID
Extracted from the page footer line format: `<ID> <PATIENT NAME> Page X of Y`
- The ID is always the **first token** on that line — can be numeric (`39037`), alphanumeric (`S0473VA`), or any format (may include letters, digits, dots, dashes, commas).
- Footer line is found via `Page 1 of X` pattern.

```javascript
function extractPatientID(text) {
  // Footer line: "<ID> <PATIENT NAME> Page X of Y" — ID is always first token
  const footerLine = text.match(/([^\n\r]*Page\s+1\s+of\s+\d[^\n\r]*)/i);
  if (footerLine) {
    const firstToken = footerLine[1].trim().match(/^(\S+)/);
    if (firstToken) return firstToken[1].trim();
  }
  // Fallback: "Patient ID:" label
  const labeled = text.match(/Patient\s*ID\s*:\s*(\S[^\n\r]*)/i);
  return labeled ? labeled[1].trim() : '';
}
```

### E — Patient Name
Extracted from `Name:` label (Patient Information section).
Stops before `Height:` on the same line using `extractAfterLabelStrict`.

### F — Age
Extracted from `Year Of Birth (Age): 1940 (86)` → returns `86` (number in brackets).
```javascript
const ageMatch = text.match(/Year\s*Of\s*Birth[^:]*:\s*\d+\s*\((\d+)\)/i);
if (ageMatch) row[COL.AGE] = ageMatch[1];
```

### G — Gender
Extracted from `Gender:` label. Mapped to single letter:
- `Female` → `F` (checked first to avoid matching "male" inside "female")
- `Male` → `M`

```javascript
if (/female/i.test(genderRaw)) row[COL.GENDER] = 'F';
else if (/male/i.test(genderRaw)) row[COL.GENDER] = 'M';
```

### H — Year
Extracted from `Year Of Birth (Age): 1939 (86)` → returns `1939` (the 4-digit year).
```javascript
const yearMatch = text.match(/Year\s*Of\s*Birth[^:]*:\s*(\d{4})/i);
if (yearMatch) row[COL.YEAR] = yearMatch[1];
```

### I — Patient Initial
Derived from Patient Name (column E): first letter of each word joined uppercase.
`CHERNOUH MOHAMED` → `CM`

```javascript
const pName = row[COL.PATIENT_NAME];
if (pName) {
  row[COL.PATIENT_INITIAL] = pName.trim().split(/\s+/).map(w => w[0]).join('').toUpperCase();
}
```

### J — Physician
Extracted from `Physician:` label using `extractAfterLabelStrict`.

### K — Hospital
Extracted from `Hospital:` label using `extractAfterLabelStrict`, then `toTitleCase()`.
Example: `CHU BLIDA` → `Chu Blida`

### L — City
Extracted from `City:` label using `extractAfterLabelStrict`, then `toTitleCase()`.
Example: `BLIDA` → `Blida`

### M — Country
Extracted from `Country:` label using `extractAfterLabelStrict`, then `toTitleCase()`.
Example: `ALGERIA` → `Algeria`

### N — Region
Fixed value: `Pending`

### O — Continent
Fixed value: `Pending`

### P — Received Date
Extracted via direct regex, strips trailing Title Case label (two-column layout).
```javascript
const rdMatch = text.match(/Received\s*Date\s*:\s*([^\n\r]+)/i);
if (rdMatch) {
  let rdVal = rdMatch[1].trim();
  rdVal = rdVal.replace(/\s+[A-Z][a-z][A-Za-z]*(?:\s+[A-Z][a-z][A-Za-z]*)*\s*:.*$/, '');
  row[COL.RECEIVED_DATE] = rdVal.trim();
}
```

### Q — Creation Date
Extracted via direct regex, strips trailing Title Case label (two-column layout).
```javascript
const cdRaw = text.match(/Creation\s*Date\s*:\s*([^\n\r]+)/i);
if (cdRaw) {
  let cdVal = cdRaw[1].trim();
  cdVal = cdVal.replace(/\s+[A-Z][a-z][A-Za-z]*(?:\s+[A-Z][a-z][A-Za-z]*)*\s*:.*$/, '');
  row[COL.CREATION_DATE] = cdVal.trim();
}
```

### R — CT Scan Date
Permissive regex to handle mixed-case label variations (`CT scan Date`, `CT Scan date`, etc.).
The separator before the date is optional — some reports write `CT Scan date 30-04-2026` with no
`:` or `-`. A trailing period (from the Comments sentence) is stripped.
```javascript
const ctMatch = text.match(/CT.{0,15}date\s*[-:]?\s*([^\n\r]+)/i);
if (ctMatch) row[COL.CT_SCAN_DATE] = ctMatch[1].trim().replace(/\.$/, '');
```

### S — Type of Valve
First line of Comments in the Aortic Valve section. Strips leading "Seems to be " qualifier.
```javascript
function extractAorticValveFirstCommentLine(text) {
  const avIdx = text.search(/aortic\s+valve/i);
  if (avIdx === -1) return '';
  const avText = text.substring(avIdx, avIdx + 3000);
  const commIdx = avText.search(/comments\s*:/i);
  if (commIdx === -1) return '';
  const afterComments = avText.substring(commIdx).replace(/comments\s*:\s*/i, '');
  const lines = afterComments.split(/[\n\r]+/);
  for (const ln of lines) {
    let t = ln.trim();
    if (t && t.length > 1) {
      t = t.replace(/^seems\s+to\s+be\s+/i, '');
      return t;
    }
  }
  return '';
}
```
Example: `Tricuspid Aortic Valve.` or `Bicuspid Type 1b Aortic Valve.`

### T — Phase
Primary: extracted from filename — first segment after `report_` (e.g. `_30%_` → `30%`, `_ECG_` → `ECG`).
Fallback: regex match `(\d+)% Phase` in PDF Comments text.

```javascript
const phaseFromFile = fileName.match(/report_([^_]+)_/i);
if (phaseFromFile) {
  row[COL.PHASE] = phaseFromFile[1].trim();
} else {
  const phaseMatch = text.match(/(\d+(?:\.\d+)?)\s*%\s*Phase/i);
  if (phaseMatch) row[COL.PHASE] = phaseMatch[1] + '%';
}
```
Examples: `30%`, `45%`, `ECG`

### U — AA Area
Extracted from `Area: 540.6 mm²` in Aortic Annulus section.
`Area Derived Ø:` won't match because it has extra words before the colon.
```javascript
const aaAreaMatch = text.match(/\bArea\s*:\s*([\d.]+)/);
if (aaAreaMatch) row[COL.AA_AREA] = aaAreaMatch[1];
```

### V — AA Area Derived Ø
Extracted from `Area Derived Ø: 26.2 mm`.
```javascript
const aaAreaDerivedMatch = text.match(/Area\s+Derived\s*[^\n:]*:\s*([\d.]+)/i);
if (aaAreaDerivedMatch) row[COL.AA_AREA_DERIVED] = aaAreaDerivedMatch[1];
```

### W — AA Perimeter
Extracted from `Perimeter: 85.0 mm` (not Perimeter Derived).
```javascript
const aaPerimMatch = text.match(/\bPerimeter\s*:\s*([\d.]+)/i);
if (aaPerimMatch) row[COL.AA_PERIMETER] = aaPerimMatch[1];
```

### X — AA Perimeter Derived Ø
Extracted from `Perimeter Derived Ø: 27.1 mm`.
```javascript
const aaPerimDerivedMatch = text.match(/Perimeter\s+Derived\s*[^\n:]*:\s*([\d.]+)/i);
if (aaPerimDerivedMatch) row[COL.AA_PERIMETER_DERIVED] = aaPerimDerivedMatch[1];
```

### Y — AA Minimum
Extracted from Measurements section: `Aortic Annulus Min Ø: 21.6 mm`.
```javascript
const aaMinMatch = text.match(/Aortic\s+Annulus\s+Min\s*[^\n:]*:\s*([\d.]+)/i);
if (aaMinMatch) row[COL.AA_MINIMUM] = aaMinMatch[1];
```

### Z — AA Maximum
Extracted from Measurements section: `Max Ø: 32.2 mm` following Aortic Annulus block.
```javascript
const aaMaxMatch = text.match(/Aortic\s+Annulus[\s\S]{0,200}?Max\s*[^\n:]*:\s*([\d.]+)/i);
if (aaMaxMatch) row[COL.AA_MAXIMUM] = aaMaxMatch[1];
```

### AA — AA Average
Extracted from Measurements section: `Average Ø: 26.9 mm` following Aortic Annulus block.
```javascript
const aaAvgMatch = text.match(/Aortic\s+Annulus[\s\S]{0,300}?Average\s*[^\n:]*:\s*([\d.]+)/i);
if (aaAvgMatch) row[COL.AA_AVERAGE] = aaAvgMatch[1];
```

### AB — AA Eccentricity
Extracted from `Eccentricity: 0.33`.
```javascript
const aaEccMatch = text.match(/Eccentricity\s*:\s*([\d.]+)/i);
if (aaEccMatch) row[COL.AA_ECCENTRICITY] = aaEccMatch[1];
```

### AC–AE — Ascending Aorta (Minimum, Maximum, Average)
Extracted from the Measurements table block "Ascending Aorta Ø ... Min: 29.1 mm Max: 31.0 mm Average: 30.0 mm".
The block is bounded between the "Ascending Aorta" label and the next "Aortic Annulus" label to avoid
matching the unrelated "Asc. Aorta Ø: 30.0 mm" summary value shown earlier on the page.

```javascript
const ascIdx = text.search(/Ascending\s+Aorta/i);
if (ascIdx !== -1) {
  const ascEnd = text.substring(ascIdx).search(/Aortic\s+Annulus/i);
  const ascBlock = ascEnd === -1 ? text.substring(ascIdx, ascIdx + 200) : text.substring(ascIdx, ascIdx + ascEnd);

  const ascMinMatch = ascBlock.match(/Min\s*:\s*([\d.]+)/i);
  if (ascMinMatch) row[COL.ASC_MINIMUM] = ascMinMatch[1];

  const ascMaxMatch = ascBlock.match(/Max\s*:\s*([\d.]+)/i);
  if (ascMaxMatch) row[COL.ASC_MAXIMUM] = ascMaxMatch[1];

  const ascAvgMatch = ascBlock.match(/Average\s*:\s*([\d.]+)/i);
  if (ascAvgMatch) row[COL.ASC_AVERAGE] = ascAvgMatch[1];
}
```

### AF–AH — LVOT (Minimum, Maximum, Average)
The Measurements table interleaves two columns per line, e.g.
`Sinotubular Junction Ø Min: 25.5 mm ... LVOT Ø Min: 22.3 mm` then `Max: 28.0 mm Max: 30.0 mm` etc.
LVOT is the **2nd "LVOT Ø"** column (right side, under Sinotubular Junction). The block is bounded
between "LVOT Ø ... Min:" and "Aorto-Mitral".

```javascript
const lvotIdx = text.search(/LVOT\s*Ø[\s\S]{0,50}?Min\s*:/i);
if (lvotIdx !== -1) {
  const lvotEnd = text.substring(lvotIdx).search(/Aorto-Mitral/i);
  const lvotBlock = lvotEnd === -1 ? text.substring(lvotIdx, lvotIdx + 200) : text.substring(lvotIdx, lvotIdx + lvotEnd);

  const lvotMinMatch = lvotBlock.match(/Min\s*:\s*([\d.]+)/i);
  if (lvotMinMatch) row[COL.LVOT_MINIMUM] = lvotMinMatch[1];

  const lvotMaxMatch = lvotBlock.match(/Max\s*:\s*([\d.]+)/i);
  if (lvotMaxMatch) row[COL.LVOT_MAXIMUM] = lvotMaxMatch[1];

  const lvotAvgMatch = lvotBlock.match(/Average\s*:\s*([\d.]+)/i);
  if (lvotAvgMatch) row[COL.LVOT_AVERAGE] = lvotAvgMatch[1];
}
```

### AI–AK — STJ (Minimum, Maximum, Average)
Same interleaved-line problem as above. From the start of "Measurements:", collect all `Min:`/`Max:`/`Average:`
matches across the whole table — the **2nd occurrence of each** is Sinotubular Junction (STJ); the 1st is
Ascending Aorta and the 3rd is LVOT.

```javascript
const measIdx = text.search(/Measurements\s*:/i);
const measText = measIdx !== -1 ? text.substring(measIdx) : text;
const stjMinAll = [...measText.matchAll(/Min\s*:\s*([\d.]+)/gi)];
const stjMaxAll = [...measText.matchAll(/Max\s*:\s*([\d.]+)/gi)];
const stjAvgAll = [...measText.matchAll(/Average\s*:\s*([\d.]+)/gi)];
if (stjMinAll[1]) row[COL.STJ_MINIMUM] = stjMinAll[1][1];
if (stjMaxAll[1]) row[COL.STJ_MAXIMUM] = stjMaxAll[1][1];
if (stjAvgAll[1]) row[COL.STJ_AVERAGE] = stjAvgAll[1][1];
```

### AL — Annulus to STJ Height
Fixed value: `Pending`.

### AM — RCA Height
Extracted from `RCA Height: 17.5 mm`.
```javascript
const rcaHeightMatch = text.match(/RCA\s+Height\s*:\s*([\d.]+)/i);
if (rcaHeightMatch) row[COL.RCA_HEIGHT] = rcaHeightMatch[1];
```

### AN — LCA Height
Extracted from `LCA Height: 12.4 mm`.
```javascript
const lcaHeightMatch = text.match(/LCA\s+Height\s*:\s*([\d.]+)/i);
if (lcaHeightMatch) row[COL.LCA_HEIGHT] = lcaHeightMatch[1];
```

### AO — SOV Height
Extracted from `Sinus of Valsalva Height  10.9 mm` (Measurements table — no colon before the value).
```javascript
const sovHeightMatch = text.match(/Sinus\s+of\s+Valsalva\s+Height\s+([\d.]+)/i);
if (sovHeightMatch) row[COL.SOV_HEIGHT] = sovHeightMatch[1];
```

### AP–AR — Sinus of Valsalva Diameters (Right/RCC, Left/LCC, Non/NCC)
Extracted from the "Sinus Of Valsalva Diameters:" box: `Left: 33.3 mm  Right: 31.1 mm  Non: 32.5 mm`.
Note column order is Right, Left, Non — i.e. AP=Right, AQ=Left, AR=Non.
```javascript
const sovLeftMatch = text.match(/Sinus\s+Of\s+Valsalva\s+Diameters[\s\S]{0,50}?Left\s*:\s*([\d.]+)/i);
if (sovLeftMatch) row[COL.SOV_LEFT] = sovLeftMatch[1];

const sovRightMatch = text.match(/Sinus\s+Of\s+Valsalva\s+Diameters[\s\S]{0,80}?Right\s*:\s*([\d.]+)/i);
if (sovRightMatch) row[COL.SOV_RIGHT] = sovRightMatch[1];

const sovNonMatch = text.match(/Sinus\s+Of\s+Valsalva\s+Diameters[\s\S]{0,110}?Non\s*:\s*([\d.]+)/i);
if (sovNonMatch) row[COL.SOV_NON] = sovNonMatch[1];
```

### AS — AO-LV Angle
Fixed value: `Pending`.

### AT–AV — ICD Measurements (4mm, 6mm, 8mm)
Fixed value: `Pending`.

### AW–BA — Aortic Valve Calcification (NC, RC, LC, Total, Grade)
Fixed value: `Pending`.

### BB–BC — Clock Angles (RCC, RCA)
Fixed value: `Pending`.

### BD–CA — Iliac/Femoral Vessel Measurements
The "Right" / "Left" Common Iliac / External Iliac / Femoral tables (page ~8) interleave both columns
per line, e.g. `Common Iliac Ø Min: 7.0 mm ... Min: 6.8 mm` then `Max: 7.9 mm Max: 7.4 mm` etc., and the
calcification row renders as both labels first then both values: `Common Iliac Calcification Common Iliac
Calcification Severe Severe`.

For each vessel section (Common Iliac / External Iliac / Femoral), the block is bounded between that
section's `Ø` label and the next section's label (or `Comments:` for Femoral). Within the block:
- `Min:`/`Max:`/`Average:` are collected with `matchAll` — **1st occurrence = Right, 2nd = Left**.
- Calcification grade is matched via `\b(None|Mild|Moderate|Severe)\b` — again 1st = Right, 2nd = Left.

```javascript
// Left Common Iliac (BD–BG) + Right Common Iliac (BP–BS)
const ciaIdx = text.search(/Common\s+Iliac\s*Ø/i);
if (ciaIdx !== -1) {
  const ciaEnd = text.substring(ciaIdx).search(/Common\s+Iliac\s+Calcification/i);
  const ciaBlock = ciaEnd === -1 ? text.substring(ciaIdx, ciaIdx + 200) : text.substring(ciaIdx, ciaIdx + ciaEnd);

  const lciaMinAll = [...ciaBlock.matchAll(/Min\s*:\s*([\d.]+)/gi)];
  const lciaMaxAll = [...ciaBlock.matchAll(/Max\s*:\s*([\d.]+)/gi)];
  const lciaAvgAll = [...ciaBlock.matchAll(/Average\s*:\s*([\d.]+)/gi)];
  if (lciaMinAll[1]) row[COL.LCIA_MIN] = lciaMinAll[1][1];
  if (lciaMaxAll[1]) row[COL.LCIA_MAX] = lciaMaxAll[1][1];
  if (lciaAvgAll[1]) row[COL.LCIA_AVG] = lciaAvgAll[1][1];

  if (lciaMinAll[0]) row[COL.RCIA_MIN] = lciaMinAll[0][1];
  if (lciaMaxAll[0]) row[COL.RCIA_MAX] = lciaMaxAll[0][1];
  if (lciaAvgAll[0]) row[COL.RCIA_AVG] = lciaAvgAll[0][1];
}

// Calcification grades (BG = Left, BS = Right)
if (ciaIdx !== -1) {
  const ciaCalcEnd = text.substring(ciaIdx).search(/External\s+Iliac\s*Ø/i);
  const ciaCalcBlock = ciaCalcEnd === -1 ? text.substring(ciaIdx, ciaIdx + 300) : text.substring(ciaIdx, ciaIdx + ciaCalcEnd);
  const ciaCalcAll = [...ciaCalcBlock.matchAll(/\b(None|Mild|Moderate|Severe)\b/gi)];
  if (ciaCalcAll[1]) row[COL.LCIA_CALC] = ciaCalcAll[1][1];
  if (ciaCalcAll[0]) row[COL.RCIA_CALC] = ciaCalcAll[0][1];
}

// Left External Iliac (BH–BK) + Right External Iliac (BT–BW)
const eiaIdx = text.search(/External\s+Iliac\s*Ø/i);
if (eiaIdx !== -1) {
  const eiaEnd = text.substring(eiaIdx).search(/Femoral\s*Ø/i);
  const eiaBlock = eiaEnd === -1 ? text.substring(eiaIdx, eiaIdx + 300) : text.substring(eiaIdx, eiaIdx + eiaEnd);

  const leiaMinAll = [...eiaBlock.matchAll(/Min\s*:\s*([\d.]+)/gi)];
  const leiaMaxAll = [...eiaBlock.matchAll(/Max\s*:\s*([\d.]+)/gi)];
  const leiaAvgAll = [...eiaBlock.matchAll(/Average\s*:\s*([\d.]+)/gi)];
  if (leiaMinAll[1]) row[COL.LEIA_MIN] = leiaMinAll[1][1];
  if (leiaMaxAll[1]) row[COL.LEIA_MAX] = leiaMaxAll[1][1];
  if (leiaAvgAll[1]) row[COL.LEIA_AVG] = leiaAvgAll[1][1];

  const leiaCalcAll = [...eiaBlock.matchAll(/\b(None|Mild|Moderate|Severe)\b/gi)];
  if (leiaCalcAll[1]) row[COL.LEIA_CALC] = leiaCalcAll[1][1];

  if (leiaMinAll[0]) row[COL.REIA_MIN] = leiaMinAll[0][1];
  if (leiaMaxAll[0]) row[COL.REIA_MAX] = leiaMaxAll[0][1];
  if (leiaAvgAll[0]) row[COL.REIA_AVG] = leiaAvgAll[0][1];
  if (leiaCalcAll[0]) row[COL.REIA_CALC] = leiaCalcAll[0][1];
}

// Left Femoral (BL–BO) + Right Femoral (BX–CA)
const faIdx = text.search(/Femoral\s*Ø/i);
if (faIdx !== -1) {
  const faEnd = text.substring(faIdx).search(/Comments\s*:/i);
  const faBlock = faEnd === -1 ? text.substring(faIdx, faIdx + 300) : text.substring(faIdx, faIdx + faEnd);

  const lfaMinAll = [...faBlock.matchAll(/Min\s*:\s*([\d.]+)/gi)];
  const lfaMaxAll = [...faBlock.matchAll(/Max\s*:\s*([\d.]+)/gi)];
  const lfaAvgAll = [...faBlock.matchAll(/Average\s*:\s*([\d.]+)/gi)];
  if (lfaMinAll[1]) row[COL.LFA_MIN] = lfaMinAll[1][1];
  if (lfaMaxAll[1]) row[COL.LFA_MAX] = lfaMaxAll[1][1];
  if (lfaAvgAll[1]) row[COL.LFA_AVG] = lfaAvgAll[1][1];

  const lfaCalcAll = [...faBlock.matchAll(/\b(None|Mild|Moderate|Severe)\b/gi)];
  if (lfaCalcAll[1]) row[COL.LFA_CALC] = lfaCalcAll[1][1];

  if (lfaMinAll[0]) row[COL.RFA_MIN] = lfaMinAll[0][1];
  if (lfaMaxAll[0]) row[COL.RFA_MAX] = lfaMaxAll[0][1];
  if (lfaAvgAll[0]) row[COL.RFA_AVG] = lfaAvgAll[0][1];
  if (lfaCalcAll[0]) row[COL.RFA_CALC] = lfaCalcAll[0][1];
}
```

### CB — Abdominal Aorta Calcification
Extracted from the Comments line, e.g. `Severe calcification observed in Abdominal Aorta` or
`Severe Calcification Observed at Abdominal Aorta`. The grade word must immediately precede
"calcification observed in/at (the) Abdominal Aorta" to avoid matching an unrelated grade word
from a later comment line.
```javascript
const abdAortaMatch = text.match(/\b(None|Mild|Moderate|Severe)\b\s*calcification\s+observed\s+(?:in|at)\s+(?:the\s+)?Abdominal\s+Aorta/i);
if (abdAortaMatch) row[COL.ABD_AORTA_CALC] = abdAortaMatch[1];
```

### CC — Aortic Comments
Fixed value: `Pending`.

### CD — Femoral Comments
Fixed value: `Pending`.

### CE — Subclavian/Carotid/Transapical Analysis
Fixed value: `Pending`.

### CF — Status
Fixed value: `NA`

### CG — Mail Notes
Fixed value: `Pending`. Manual entry.

### CH — Meril Representative
Fixed value: `Pending`. Manual entry.

### CI — Received from
Fixed value: `Pending`. Manual entry.

---

## PDF Text Extraction — Critical Implementation

3mensio PDFs render each character as a separate PDF.js text item.
Simple `items.map(i => i.str).join('\n')` produces spaced-out text like `C r e a t e d B y`.

**Correct approach:** group by Y coordinate and use X-position gaps to decide spacing.

```javascript
async function readPDFText(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let fullText = '';

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();

    const lineMap = new Map();
    tc.items.forEach(item => {
      if (!item.str) return;
      const y = Math.round(item.transform[5] / 5) * 5; // 5px rounding merges split words
      if (!lineMap.has(y)) lineMap.set(y, []);
      lineMap.get(y).push({ x: item.transform[4], w: item.width || 0, str: item.str });
    });

    const sortedYs = Array.from(lineMap.keys()).sort((a, b) => b - a);
    const pageLines = sortedYs.map(y => {
      const items = lineMap.get(y).sort((a, b) => a.x - b.x);
      let line = '';
      for (let i = 0; i < items.length; i++) {
        const cur = items[i];
        if (i === 0) { line += cur.str; continue; }
        const prev = items[i - 1];
        const gap = cur.x - (prev.x + prev.w);
        const charW = prev.w > 0 ? prev.w / (prev.str.length || 1) : 3;
        line += gap > charW * 0.5 ? ' ' + cur.str : cur.str;
      }
      return line.trim();
    }).filter(l => l.length > 0);

    fullText += pageLines.join('\n') + '\n';
  }
  return fullText;
}
```

---

## Field Extraction — Two-Column Layout

3mensio PDFs use a two-column layout. The same line contains two label:value pairs:
```
Creation Date: 01-06-2026 Physician: Pr Nedjar Dr Takhdem
Created By: MERIL TAVI CORELAB (AG) Hospital: Chu Blida
```

`extractAfterLabelStrict` handles this by:
1. Finding the label (case-insensitive)
2. Taking text up to end of line
3. Stopping before the next **Title Case** label (e.g. `Hospital:`, `City:`) using a **case-sensitive** regex
4. ALL-CAPS words (e.g. `MERIL`, `TAVI`) are NOT treated as label boundaries

```javascript
function extractAfterLabelStrict(text, label) {
  const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const labelRe = new RegExp(escaped + '\\s*:\\s*', 'i');
  const labelMatch = labelRe.exec(text);
  if (!labelMatch) return '';

  const afterLabel = text.substring(labelMatch.index + labelMatch[0].length);
  const nlIdx = afterLabel.search(/[\n\r]/);
  const sameLine = nlIdx >= 0 ? afterLabel.substring(0, nlIdx) : afterLabel;

  // Case-sensitive stop: Title Case word(s) + colon
  const titleCaseStop = /\s+[A-Z][a-z][A-Za-z]*(?:\s+[A-Z][a-z][A-Za-z]*)*\s*:/;
  const stopMatch = titleCaseStop.exec(sameLine);
  if (stopMatch && stopMatch.index > 0) return sameLine.substring(0, stopMatch.index).trim();

  if (!sameLine.trim()) {
    const nextLine = afterLabel.replace(/^[\n\r]+/, '').match(/^([^\n\r]+)/);
    if (nextLine) return nextLine[1].trim();
  }
  return sameLine.trim();
}
```

---

## Empty Cell Rule

If data is not found, leave the Excel cell **blank** (empty string).
Do NOT fill missing fields with `NA`.
Only `CF (Status)` is hardcoded to `NA`.

---

## toTitleCase Helper

Converts ALL-CAPS strings to Title Case (e.g. `CHU BLIDA` → `Chu Blida`, `ALGERIA` → `Algeria`).
Applied to Hospital (K), City (L), Country (M).

```javascript
function toTitleCase(str) {
  if (!str) return str;
  return str.replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
}
```

---

## Column Index Map

```javascript
const COL = {
  CASE_ENTRY_DATE: 0,   // A
  CASE_ENTERED_BY: 1,   // B
  CASE_DONE_BY: 2,      // C
  PATIENT_ID: 3,        // D
  PATIENT_NAME: 4,      // E
  AGE: 5,               // F
  GENDER: 6,            // G
  YEAR: 7,              // H
  PATIENT_INITIAL: 8,   // I
  PHYSICIAN: 9,         // J
  HOSPITAL: 10,         // K
  CITY: 11,             // L
  COUNTRY: 12,          // M
  REGION: 13,           // N
  CONTINENT: 14,        // O
  RECEIVED_DATE: 15,    // P
  CREATION_DATE: 16,    // Q
  CT_SCAN_DATE: 17,     // R
  TYPE_OF_VALVE: 18,    // S
  PHASE: 19,            // T
  AA_AREA: 20,          // U
  AA_AREA_DERIVED: 21,  // V
  AA_PERIMETER: 22,     // W
  AA_PERIMETER_DERIVED: 23, // X
  AA_MINIMUM: 24,       // Y
  AA_MAXIMUM: 25,       // Z
  AA_AVERAGE: 26,       // AA
  AA_ECCENTRICITY: 27,  // AB
  ASC_MINIMUM: 28,      // AC
  ASC_MAXIMUM: 29,      // AD
  ASC_AVERAGE: 30,      // AE
  LVOT_MINIMUM: 31,     // AF
  LVOT_MAXIMUM: 32,     // AG
  LVOT_AVERAGE: 33,     // AH
  STJ_MINIMUM: 34,      // AI
  STJ_MAXIMUM: 35,      // AJ
  STJ_AVERAGE: 36,      // AK
  ANNULUS_STJ_HEIGHT: 37, // AL
  RCA_HEIGHT: 38,       // AM
  LCA_HEIGHT: 39,       // AN
  SOV_HEIGHT: 40,       // AO
  SOV_RIGHT: 41,        // AP
  SOV_LEFT: 42,         // AQ
  SOV_NON: 43,          // AR
  AO_LV_ANGLE: 44,      // AS
  ICD_4MM: 45,          // AT
  ICD_6MM: 46,          // AU
  ICD_8MM: 47,          // AV
  CALC_NC: 48,          // AW
  CALC_RC: 49,          // AX
  CALC_LC: 50,          // AY
  CALC_TOTAL: 51,       // AZ
  CALC_GRADE: 52,       // BA
  CLOCK_RCC: 53,        // BB
  CLOCK_RCA: 54,        // BC
  LCIA_MIN: 55,         // BD
  LCIA_MAX: 56,         // BE
  LCIA_AVG: 57,         // BF
  LCIA_CALC: 58,        // BG
  LEIA_MIN: 59,         // BH
  LEIA_MAX: 60,         // BI
  LEIA_AVG: 61,         // BJ
  LEIA_CALC: 62,        // BK
  LFA_MIN: 63,          // BL
  LFA_MAX: 64,          // BM
  LFA_AVG: 65,          // BN
  LFA_CALC: 66,         // BO
  RCIA_MIN: 67,         // BP
  RCIA_MAX: 68,         // BQ
  RCIA_AVG: 69,         // BR
  RCIA_CALC: 70,        // BS
  REIA_MIN: 71,         // BT
  REIA_MAX: 72,         // BU
  REIA_AVG: 73,         // BV
  REIA_CALC: 74,        // BW
  RFA_MIN: 75,          // BX
  RFA_MAX: 76,          // BY
  RFA_AVG: 77,          // BZ
  RFA_CALC: 78,         // CA
  ABD_AORTA_CALC: 79,   // CB
  AORTIC_COMMENTS: 80,  // CC
  FEMORAL_COMMENTS: 81, // CD
  SUBCLAVIAN: 82,       // CE
  STATUS: 83,           // CF
  MAIL_NOTES: 84,       // CG
  MERIL_REP: 85,        // CH
  RECEIVED_FROM: 86     // CI
};
```

---

## Header Arrays & Merge Ranges

See existing `HEADER_ROW1`, `HEADER_ROW2`, and `MERGE_RANGES` constants in the HTML file.
Both header arrays must have exactly 87 values.
Sheet name must be `Sample Tracker` (matches the sample workbook).

---

## Debug Panel

A debug raw-text panel is embedded in the page (visible after processing).
Each processed PDF gets its own tab (labeled `Row N`, filename as tooltip) showing the
first 10,000 chars of that file's extracted text, for diagnosing extraction failures.
Tabs are rebuilt on every "Process PDFs" run. The first tab is shown by default.
Remove or hide this panel once all fields are confirmed working.

---

## Preview Table — Blank Cell Highlighting

In the on-page "Extracted Data Preview" table, any cell with no extracted value renders as
`—` with a yellow background (`.blank-val` class) so missing/uncertain fields are easy to
spot at a glance before download.

---

## Important Rules

- No backend, no Node.js, no Python, no server.
- No fake values. No guessing.
- Missing fields → blank (empty string). Only CF (Status) = `NA`.
- Patient ID must NOT be used as Type of Valve.
- Column order must never change.
- Always export all 87 columns even if most are `NA`.
- Code must remain readable — extraction rules will need ongoing refinement.
