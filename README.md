# VTU Results Utility (Combined App)

> If you find any bug, issue, or have suggestions for improvement, please open an issue on GitHub: https://github.com/Sam-Builds/VTUResultsUitlity/issues/new

A desktop utility for VTU result operations with two integrated workflows:

- VTU Parser (primary window): parse downloaded/scanned VTU PDFs and export structured Excel sheets.
- VTU Scraper ( child window): fetch result PDFs by USN range with manual/auto captcha handling.

## What This Build Includes

### 1) Combined App Architecture

- Entry app
- Main UI mode: parser-first interface
- Scraper opens as child modal from parser using OPEN VTU SCRAPER

### 2) Parser Workflow

- Folder input via Browse
- Parse & Export to Excel with parallel PDF parsing workers
- Cancel Parsing button with cancellation propagation and safe shutdown behavior
- Open Output Folder button
- Open Output Excel button (enabled after export)

### 3) Scraper Workflow

- Opens as modal from parser
- Can run Auto captcha mode (OCR) or manual captcha mode
- Output folder resolves to Desktop/prefix location
- Run completion callback is sent back to parser only when all requested USNs are downloaded
- Cancellation callback is propagated to parser flow when applicable
- Per-run diagnostics report is generated:
  - scrape_run_report_YYYYMMDD_HHMMSS.txt
  - Includes requested, downloaded, no-result, and errored USNs

### 4) Captcha Reliability Fix (Packaged Install Safe)

To avoid stale captcha reuse in packaged installs:

- Captcha image is captured to a unique temp file in OS temp directory
- A new file is generated per USN and attempt
- Temp captcha file is deleted after each attempt

This prevents accidental reuse of old packaged captcha images.

### 5) Persistent Configuration (AppData)

Parser now supports persistent sheet metadata via config file:

- Config file path:
  - %APPDATA%/VTUResultsUtility/config.json

Saved fields:

- College name
- Department name
- Year period
- Revaluation status
- Semester
- Faculty In-charge

Behavior:

- Values are loaded automatically on app startup
- Save button in Sheet Configuration writes current values to config

### 6) New Faculty In-charge Field

Parser Sheet Configuration includes a new Faculty In-charge textbox.

In exported Excel, this appears in the header block near the top as:

- Faculty In-charge: <value>

### 7) Subject Credits Popup Improvements

Credit popup supports:

- Row reorder with Up/Down controls
- Combine exactly 2 selected subjects
- Per-subject credit entry validation
- Per-subject highlight toggle
- Scrollable table area for long subject lists
- Fixed action buttons at bottom

### 8) Auto Highlight Rule Adjustment for External=0 Subjects

Before credit popup opens, parser checks parsed data by subject:

- If a subject has External mark = 0 for all rows, highlight is auto-disabled for that subject by default
- User can still manually re-enable highlight in the popup before export

### 9) Excel Output Structure

Parser exports one workbook with sheets:

- Result Sheet
- Credit Sheet
- Raw Data

Header block includes:

- College
- Department
- Result sheet title with year/revaluation/semester
- Faculty In-charge line

Other Excel behavior includes:

- Subject-wise INT/EXT/TOT columns
- Total marks, percentage, fail/backlog counts
- Credit Sheet GP/CP/SGPA logic
- Summary blocks and formatting
- Fallback save name if file is locked

## Guardrails and Stability

Implemented distribution-focused safeguards:

- Busy-state guards for parse/scraper actions
- Modal duplication prevention for scraper window
- Safe close protection when parse/scrape is running
- Parser cancel flag and thread-safe queue updates
- Widget existence checks to avoid TclError after window destruction
- Controlled state reset on completion/cancel/error

## Known Heavy Dependencies

Final bundle size is dominated by OCR/vision stack dependencies:

- torch
- easyocr
- opencv
- scipy/skimage-related transitive packages

This is expected for auto captcha OCR support.

## Troubleshooting
If OCR fails to initialize:

- Check ocr_crash.txt generated in runtime context
- Verify model folder is bundled correctly
- Verify required OCR dependencies are present in release environment

