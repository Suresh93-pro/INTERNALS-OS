# INTERNALS OS

A local Flask + SQLite internal assessment marks manager for two student batches.

> **Privacy:** This public repository contains application source code only. Student names, registration numbers, marks, the local SQLite database, and exported spreadsheets are not included. Keep your real `roster.csv`, `internals.db`, and exported reports on a private device or private storage.

## Features

- Batch 1, Batch 2, and combined student directory views.
- Individual marks editor with automatic saving.
- Spreadsheet-style marks entry.
- CT1, Java/HackerRank, Tutorial Assessments, eCurricula, Quiz, Tkinter project/viva, and Surprise Test components.
- CT2 Part A: Q1–Q6, 5 marks each; all six questions count.
- CT2 Part B: Q7A/Q7B and Q8A/Q8B; only the higher alternative counts for each question.
- CT2 raw score out of 50 and converted score out of 10.
- Grand total out of 100 and student completion progress.
- Main Excel reports show CT2 only as a converted /10 score.
- Separate CT2 detailed reports include question-wise marks and totals.
- Import edited Excel workbooks.
- Neon-themed responsive interface and live directory refresh.

## Requirements

- Python 3.10 or newer
- Flask 3.x
- OpenPyXL 3.x

## Run on Windows

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
py app.py
```

Open `http://127.0.0.1:5000`.

## Add your private student roster

The real student roster is intentionally not tracked by Git. Create a local file named `roster.csv` beside `app.py` using this format:

```csv
registration_number,name,batch
YOUR_REGISTRATION_NUMBER,Student Name,Batch 1
ANOTHER_REGISTRATION_NUMBER,Another Student,Batch 2
```

You can start the app without `roster.csv`; the database will simply start without seeded students. When `roster.csv` is present, the app seeds missing students on startup without overwriting existing marks.

`roster.csv` is excluded by `.gitignore`. Do not remove that protection or commit student data.

## Local data and backups

- The SQLite database is `internals.db` beside `app.py`.
- Back up `internals.db` privately before replacing application files or changing versions.
- Excel exports may contain student information and are excluded from Git.
- No teacher login is enabled. The development server binds to `127.0.0.1`; do not expose it to the public internet without adding authentication and deployment security.

## Configuration

You may set a private `SECRET_KEY` environment variable. If omitted, a random local secret is generated at startup.
