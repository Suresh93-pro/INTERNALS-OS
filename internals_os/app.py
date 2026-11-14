from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, flash
from pathlib import Path
import sqlite3, csv, io, os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DB = BASE / 'internals.db'
ROSTER = BASE / 'roster.csv'
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'internals-os-local-key-change-if-hosting')

FIELDS = [
    ('ct1', 'CT1', 10), ('hackerrank', 'Java HackerRank', 10),
    ('tut1', 'Tutorial Assessment 1', 10), ('ecurricula', 'eCurricula', 10),
    ('ct2_q1', 'CT2 Q1', 5), ('ct2_q2', 'CT2 Q2', 5), ('ct2_q3', 'CT2 Q3', 5),
    ('ct2_q4', 'CT2 Q4', 5), ('ct2_q5', 'CT2 Q5', 5), ('ct2_q6', 'CT2 Q6', 5),
    ('ct2_7a', 'CT2 Q7A', 10), ('ct2_7b', 'CT2 Q7B', 10),
    ('ct2_8a', 'CT2 Q8A', 10), ('ct2_8b', 'CT2 Q8B', 10),
    ('tut2', 'Tutorial Assessment 2', 10), ('quiz', 'Unit 5 Quiz', 10),
    ('tut3_app', 'Tutorial Assessment 3 Application', 5), ('tut3_viva', 'Tutorial Assessment 3 Viva', 5),
    ('tkinter_project', 'Tkinter Project', 10), ('tkinter_viva', 'Tkinter Viva', 5), ('surprise', 'Surprise Test', 5)
]

# Main internal reports intentionally show CT2 only as a converted /10 score.
# Question-wise CT2 marks are available in the dedicated CT2 detailed reports.
MAIN_FIELDS = [
    ('ct1', 'CT1', 10), ('hackerrank', 'Java HackerRank', 10),
    ('tut1', 'Tutorial Assessment 1', 10), ('ecurricula', 'eCurricula', 10),
    ('tut2', 'Tutorial Assessment 2', 10), ('quiz', 'Unit 5 Quiz', 10),
    ('tut3_app', 'Tutorial Assessment 3 Application', 5), ('tut3_viva', 'Tutorial Assessment 3 Viva', 5),
    ('tkinter_project', 'Tkinter Project', 10), ('tkinter_viva', 'Tkinter Viva', 5), ('surprise', 'Surprise Test', 5)
]

def db_conn():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    with db_conn() as con:
        con.execute('''CREATE TABLE IF NOT EXISTS students (
            registration_number TEXT PRIMARY KEY, name TEXT NOT NULL, batch TEXT NOT NULL,
            ct1 REAL, hackerrank REAL, tut1 REAL, ecurricula REAL,
            ct2_q1 REAL, ct2_q2 REAL, ct2_q3 REAL, ct2_q4 REAL, ct2_q5 REAL, ct2_q6 REAL,
            ct2_7a REAL, ct2_7b REAL, ct2_8a REAL, ct2_8b REAL, tut2 REAL, quiz REAL,
            tut3_app REAL, tut3_viva REAL, tkinter_project REAL, tkinter_viva REAL, surprise REAL,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP)''')
        # Seed every startup, not only on an empty database. This safely adds
        # newly supplied roster entries (such as the missing Batch 1 student)
        # without overwriting existing marks or student records.
        with ROSTER.open(newline='', encoding='utf-8-sig') as f:
            rows = list(csv.DictReader(f))
        con.executemany('INSERT OR IGNORE INTO students(registration_number,name,batch) VALUES(?,?,?)',
            [(r['registration_number'], r['name'], r['batch']) for r in rows])

def num(v):
    if v is None or str(v).strip() == '': return None
    try: return float(v)
    except (ValueError, TypeError): raise ValueError('Marks must be numeric or blank.')

def clamp_validate(data):
    out = {}
    for key, label, maximum in FIELDS:
        value = num(data.get(key))
        if value is not None and not (0 <= value <= maximum):
            raise ValueError(f'{label} must be between 0 and {maximum}.')
        out[key] = value
    return out

def totals(s):
    def v(k): return float(s[k] or 0)
    raw = sum(v(f'ct2_q{i}') for i in range(1,7)) + max(v('ct2_7a'),v('ct2_7b')) + max(v('ct2_8a'),v('ct2_8b'))
    ct2 = raw / 5
    total = v('ct1') + v('hackerrank') + v('tut1') + v('ecurricula') + ct2 + v('tut2') + v('quiz') + v('tut3_app') + v('tut3_viva') + v('tkinter_project') + v('tkinter_viva') + v('surprise')
    # Completion is 11 standalone fields plus CT2 (six required questions and
    # one selected alternative for each of Q7 and Q8).
    standalone = ['ct1','hackerrank','tut1','ecurricula','tut2','quiz','tut3_app','tut3_viva','tkinter_project','tkinter_viva','surprise']
    done = sum(s[k] is not None for k in standalone)
    ct2_done = all(s[f'ct2_q{i}'] is not None for i in range(1,7)) and (s['ct2_7a'] is not None or s['ct2_7b'] is not None) and (s['ct2_8a'] is not None or s['ct2_8b'] is not None)
    done += int(ct2_done)
    return round(raw,2), round(ct2,2), round(total,2), done, 12

@app.route('/')
def index():
    batch = request.args.get('batch', 'Batch 1')
    if batch not in ('Batch 1','Batch 2','All'): batch='Batch 1'
    with db_conn() as con:
        rows = con.execute('SELECT * FROM students' + ('' if batch=='All' else ' WHERE batch=?') + ' ORDER BY registration_number', (() if batch=='All' else (batch,))).fetchall()
        counts = {b: con.execute('SELECT COUNT(*) FROM students WHERE batch=?',(b,)).fetchone()[0] for b in ('Batch 1','Batch 2')}
        all_rows = con.execute('SELECT * FROM students ORDER BY registration_number').fetchall()
    records=[]
    for r in rows:
        d=dict(r); d['ct2_raw'],d['ct2'],d['total'],d['progress_count'],d['progress_max']=totals(d); records.append(d)
    completed = sum(totals(dict(r))[3] == 12 for r in all_rows)
    return render_template('index.html', records=records, batch=batch, is_dashboard=('batch' not in request.args), counts=counts, total_students=sum(counts.values()), completed=completed, fields=MAIN_FIELDS)

@app.get('/batch/<int:batch_id>')
def batch_view(batch_id):
    if batch_id not in (1, 2): return redirect(url_for('index'))
    return redirect(url_for('index', batch=f'Batch {batch_id}'))

@app.get('/api/student/<reg>')
def get_student(reg):
    with db_conn() as con: row=con.execute('SELECT * FROM students WHERE registration_number=?',(reg,)).fetchone()
    if not row: return jsonify(error='Student not found'),404
    d=dict(row); d['ct2_raw'],d['ct2'],d['total'],d['progress_count'],d['progress_max']=totals(d)
    return jsonify(d)

@app.post('/api/student/<reg>')
def save_student(reg):
    data=request.get_json(silent=True) or request.form
    try: marks=clamp_validate(data)
    except ValueError as e: return jsonify(error=str(e)),400
    with db_conn() as con:
        cur=con.execute('UPDATE students SET '+','.join(f'{k}=?' for k,_,_ in FIELDS)+', updated_at=CURRENT_TIMESTAMP WHERE registration_number=?', (*marks.values(),reg))
        if cur.rowcount==0: return jsonify(error='Student not found'),404
        row=con.execute('SELECT * FROM students WHERE registration_number=?',(reg,)).fetchone()
    d=dict(row); d['ct2_raw'],d['ct2'],d['total'],d['progress_count'],d['progress_max']=totals(d)
    with db_conn() as con:
        completed=sum(1 for item in con.execute('SELECT * FROM students').fetchall() if totals(dict(item))[3] == 12)
        total_students=con.execute('SELECT COUNT(*) FROM students').fetchone()[0]
    d['completed_students']=completed
    d['total_students']=total_students
    return jsonify(ok=True, student=d)

@app.get('/export/<batch>')
def export_excel(batch):
    if batch not in ('Batch 1','Batch 2','All'): return 'Invalid batch',400
    with db_conn() as con:
        rows=con.execute('SELECT * FROM students'+('' if batch=='All' else ' WHERE batch=?')+' ORDER BY registration_number', (() if batch=='All' else (batch,))).fetchall()
    wb=Workbook(); ws=wb.active; ws.title='Internal Marks'
    headers=['Batch','Registration Number','Student Name']+[label for _,label,_ in MAIN_FIELDS]+['CT2 / 10','Grand Total / 100']
    ws.append(headers)
    for r in rows:
        d=dict(r); raw,ct2,total,_,_=totals(d)
        ws.append([d['batch'],d['registration_number'],d['name']]+[d[k] for k,_,_ in MAIN_FIELDS]+[ct2,total])
    for cell in ws[1]:
        cell.font=Font(bold=True,color='FFFFFF'); cell.fill=PatternFill('solid',fgColor='6D28D9'); cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.freeze_panes='D2'; ws.auto_filter.ref=ws.dimensions
    for col in ws.columns:
        letter=get_column_letter(col[0].column); width=min(max(max(len(str(c.value or '')) for c in col)+2,13),34); ws.column_dimensions[letter].width=width
    ws.row_dimensions[1].height=42
    out=io.BytesIO(); wb.save(out); out.seek(0)
    filename='Internals_OS_'+batch.replace(' ','_')+'.xlsx'
    return send_file(out,as_attachment=True,download_name=filename,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.get('/export/CT2/<batch>')
def export_ct2_batch_excel(batch):
    if batch not in ('Batch 1', 'Batch 2'):
        return 'Invalid batch', 400
    with db_conn() as con:
        rows = con.execute('SELECT * FROM students WHERE batch=? ORDER BY registration_number', (batch,)).fetchall()
    wb = Workbook(); ws = wb.active; ws.title = batch
    headers = ['Registration Number', 'Student Name', 'Q1 / 5', 'Q2 / 5', 'Q3 / 5', 'Q4 / 5', 'Q5 / 5', 'Q6 / 5', 'Q7A / 10', 'Q7B / 10', 'Q7 Counted / 10', 'Q8A / 10', 'Q8B / 10', 'Q8 Counted / 10', 'CT2 Raw Total / 50', 'CT2 Converted / 10']
    ws.append(headers)
    for row in rows:
        d = dict(row)
        q7 = max(float(d['ct2_7a'] or 0), float(d['ct2_7b'] or 0))
        q8 = max(float(d['ct2_8a'] or 0), float(d['ct2_8b'] or 0))
        raw = sum(float(d[f'ct2_q{i}'] or 0) for i in range(1, 7)) + q7 + q8
        ws.append([d['registration_number'], d['name'], *[d[f'ct2_q{i}'] for i in range(1, 7)], d['ct2_7a'], d['ct2_7b'], q7, d['ct2_8a'], d['ct2_8b'], q8, round(raw, 2), round(raw / 5, 2)])
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF'); cell.fill = PatternFill('solid', fgColor='6D28D9'); cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.freeze_panes = 'C2'; ws.auto_filter.ref = ws.dimensions; ws.row_dimensions[1].height = 40
    for col in ws.columns:
        letter = get_column_letter(col[0].column); ws.column_dimensions[letter].width = min(max(max(len(str(c.value or '')) for c in col) + 2, 12), 30)
    out = io.BytesIO(); wb.save(out); out.seek(0)
    filename = f"Internals_OS_{batch.replace(' ', '_')}_CT2_Detailed.xlsx"
    return send_file(out, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.get('/export/CT2')
def export_ct2_excel():
    """Export a CT2-only workbook with detailed question marks for each batch."""
    with db_conn() as con:
        all_rows = con.execute('SELECT * FROM students ORDER BY registration_number').fetchall()

    wb = Workbook()
    # Three ready-to-use views in one dedicated CT2 workbook.
    for sheet_name, batch_filter in [('Batch 1', 'Batch 1'), ('Batch 2', 'Batch 2'), ('Combined', None)]:
        ws = wb.active if sheet_name == 'Batch 1' else wb.create_sheet()
        ws.title = sheet_name
        headers = [
            'Batch', 'Registration Number', 'Student Name',
            'Q1 / 5', 'Q2 / 5', 'Q3 / 5', 'Q4 / 5', 'Q5 / 5', 'Q6 / 5',
            'Q7A / 10', 'Q7B / 10', 'Q7 Counted / 10',
            'Q8A / 10', 'Q8B / 10', 'Q8 Counted / 10',
            'CT2 Raw Total / 50', 'CT2 Converted / 10'
        ]
        ws.append(headers)
        for row in all_rows:
            d = dict(row)
            if batch_filter and d['batch'] != batch_filter:
                continue
            q7 = max(float(d['ct2_7a'] or 0), float(d['ct2_7b'] or 0))
            q8 = max(float(d['ct2_8a'] or 0), float(d['ct2_8b'] or 0))
            raw = sum(float(d[f'ct2_q{i}'] or 0) for i in range(1, 7)) + q7 + q8
            ws.append([
                d['batch'], d['registration_number'], d['name'],
                *[d[f'ct2_q{i}'] for i in range(1, 7)],
                d['ct2_7a'], d['ct2_7b'], q7,
                d['ct2_8a'], d['ct2_8b'], q8,
                round(raw, 2), round(raw / 5, 2)
            ])
        for cell in ws[1]:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='6D28D9')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.freeze_panes = 'D2'
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 38
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            width = min(max(max(len(str(c.value or '')) for c in col) + 2, 12), 30)
            ws.column_dimensions[letter].width = width

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(
        out, as_attachment=True, download_name='Internals_OS_CT2_Detailed_Report.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.post('/import')
def import_excel():
    file=request.files.get('file')
    if not file or not file.filename.lower().endswith('.xlsx'):
        flash('Choose an .xlsx file exported from Internals OS.','error'); return redirect(url_for('index'))
    try:
        wb=load_workbook(file, data_only=True); ws=wb.active
        headers=[c.value for c in ws[1]]; idx={h:i for i,h in enumerate(headers)}
        if 'Registration Number' not in idx: raise ValueError('Registration Number column is missing.')
        imported=0
        is_ct2 = 'Q1 / 5' in idx and 'Q6 / 5' in idx
        with db_conn() as con:
            for row in ws.iter_rows(min_row=2,values_only=True):
                reg=row[idx['Registration Number']]
                if not reg: continue
                reg=str(reg).strip()
                exists=con.execute('SELECT 1 FROM students WHERE registration_number=?',(reg,)).fetchone()
                if not exists: continue
                if is_ct2:
                    ct2_columns = [('ct2_q1','Q1 / 5',5),('ct2_q2','Q2 / 5',5),('ct2_q3','Q3 / 5',5),('ct2_q4','Q4 / 5',5),('ct2_q5','Q5 / 5',5),('ct2_q6','Q6 / 5',5),('ct2_7a','Q7A / 10',10),('ct2_7b','Q7B / 10',10),('ct2_8a','Q8A / 10',10),('ct2_8b','Q8B / 10',10)]
                    marks={key:num(row[idx[label]] if label in idx else None) for key,label,_ in ct2_columns}
                    for key,label,maximum in ct2_columns:
                        if marks[key] is not None and not (0 <= marks[key] <= maximum):
                            raise ValueError(f'{label} must be between 0 and {maximum}.')
                    con.execute('UPDATE students SET '+','.join(f'{k}=?' for k,_,_ in ct2_columns)+', updated_at=CURRENT_TIMESTAMP WHERE registration_number=?',(*marks.values(),reg))
                else:
                    data={k:(row[idx[label]] if label in idx else None) for k,label,_ in MAIN_FIELDS}
                    marks={k:num(data.get(k)) for k,_,_ in MAIN_FIELDS}
                    for k,_,maximum in MAIN_FIELDS:
                        if marks[k] is not None and not (0 <= marks[k] <= maximum):
                            raise ValueError(f'{k} must be between 0 and {maximum}.')
                    con.execute('UPDATE students SET '+','.join(f'{k}=?' for k,_,_ in MAIN_FIELDS)+', updated_at=CURRENT_TIMESTAMP WHERE registration_number=?',(*marks.values(),reg))
                imported+=1
        report_type='CT2 question-wise marks' if is_ct2 else 'internal marks'
        flash(f'Imported {report_type} for {imported} matching students.','success')
    except Exception as e: flash(f'Import failed: {e}','error')
    return redirect(url_for('index'))

@app.post('/api/bulk')
def bulk_save():
    payload=request.get_json(silent=True) or {}; updates=payload.get('updates',[]); saved=0
    try:
        with db_conn() as con:
            allowed={k: maximum for k,_,maximum in FIELDS}
            for item in updates:
                reg=item.get('registration_number')
                if not reg: raise ValueError('Registration number is required.')
                marks={}
                for key, value in item.items():
                    if key == 'registration_number': continue
                    if key not in allowed: raise ValueError(f'Unknown marks field: {key}')
                    parsed=num(value)
                    if parsed is not None and not (0 <= parsed <= allowed[key]):
                        raise ValueError(f'{key} must be between 0 and {allowed[key]}.')
                    marks[key]=parsed
                if marks:
                    con.execute('UPDATE students SET '+','.join(f'{k}=?' for k in marks)+', updated_at=CURRENT_TIMESTAMP WHERE registration_number=?',(*marks.values(),reg))
                saved+=1
        updated=[]
        with db_conn() as con:
            for item in updates:
                reg=item.get('registration_number')
                row=con.execute('SELECT * FROM students WHERE registration_number=?',(reg,)).fetchone()
                if row:
                    d=dict(row); d['ct2_raw'],d['ct2'],d['total'],d['progress_count'],d['progress_max']=totals(d); updated.append(d)
            completed=0
            for row in con.execute('SELECT * FROM students').fetchall():
                if totals(dict(row))[3] == 12: completed += 1
            total_students=con.execute('SELECT COUNT(*) FROM students').fetchone()[0]
        for student in updated:
            student['completed_students']=completed
            student['total_students']=total_students
        return jsonify(ok=True,saved=saved,students=updated,completed_students=completed,total_students=total_students)
    except Exception as e: return jsonify(error=str(e)),400

if __name__ == '__main__':
    init_db()
    app.run(host='127.0.0.1', port=5000, debug=False)
