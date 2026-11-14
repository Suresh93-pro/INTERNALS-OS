# INTERNALS OS v1.3

A local Flask + SQLite internal marks manager for Batch 1 and Batch 2.

## Included in this build
- Batch navigation with dedicated `/batch/1` and `/batch/2` routes.
- 61 Batch 1 students and 58 Batch 2 students (119 total).
- Missing Batch 1 roster entry added: `RA2511030010233 — SHAURYA SINGH`.
- Individual marks editor with automatic saving.
- Full spreadsheet-style marks entry, including all CT2 question fields.
- CT2 Part A: Q1–Q6, 5 marks each.
- CT2 Part B: Q7A/Q7B and Q8A/Q8B; only the higher mark for each question counts.
- CT2 raw score /50 and converted score /10.
- Grand total /100, colour-coded missing marks, and per-student completion progress.
- Neon hover effects and smooth page transitions.
- Main Batch 1 and Batch 2 Excel exports plus combined export; CT2 questions are excluded from main reports and shown only as CT2 /10.
- Separate Batch 1 CT2 and Batch 2 CT2 detailed Excel reports.
- Import edited main marks workbooks without overwriting question-wise CT2 marks; import CT2 detailed workbooks to update question-wise CT2 marks.
- Student directory and spreadsheet values refresh immediately after autosave.
- Roster seeding runs on every app startup using `INSERT OR IGNORE`, so new roster entries are added without overwriting existing marks.

## Run on Windows

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
py app.py
```

Open `http://127.0.0.1:5000`.

## Notes
- The SQLite database is `internals.db` in this folder. Back it up before replacing application files.
- No login is enabled, as requested. Keep the app on your own computer/local network unless access controls are added.
- Blank marks remain blank in the interface and Excel; they contribute zero to the current total.
- CT2 Q7/Q8 alternatives are never added together. The higher entered score is used; if both are blank, the question contributes zero.


## v1.2 CT2 detailed report
- Added a dedicated CT2 Detailed Report Excel export. The workbook contains Batch 1, Batch 2, and Combined worksheets.
- Each worksheet includes Q1-Q6, Q7A/Q7B and counted Q7, Q8A/Q8B and counted Q8, raw total out of 50, and converted score out of 10.
- CT2 rule: all six Part A questions count; only the higher of A/B counts for Q7 and Q8. Conversion is raw total divided by 5.
- The student marks workspace now labels the regular internal components separately from the question-wise CT2 section.


## v1.3 reporting and live-refresh updates
- Student marks saved in the editor now immediately refresh the matching row in the directory and spreadsheet, including CT2 /10, grand total /100, and progress.
- Batch 1 and Batch 2 detailed CT2 reports are separate downloadable Excel files.
- Main reports contain only the converted CT2 /10 column; individual CT2 question marks are omitted.
- Main workbook import updates regular internal components and leaves CT2 question marks unchanged. CT2 detailed workbook import updates only CT2 question fields.
