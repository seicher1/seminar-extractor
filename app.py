from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from itertools import zip_longest
from openpyxl.utils import get_column_letter
from docx.table import Table
from docx.text.paragraph import Paragraph

app = Flask(__name__)

# Degrees (exact forms)
DEGREES = {
    "ierindnieks","ierindniece","kaprālis","kaprāliene",
    "seržants","seržante","virsseržants","virsseržante",
    "virsniekvietnieks","virsniekvietniece","leitnants","leitnante",
    "virsleitnants","virsleitnante","kapteinis","kapteine",
    "majors","majore","pulkvežleitnants","pulkvežleitnante",
    "pulkvedis","pulkvede","ģenerālis","ģenerāle"
}

# Simple lecturer name fallback
NAME_RE = re.compile(r"\b([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\s+([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\b")

# Date/time patterns
date_pattern = re.compile(r"202\d\. gada \d{1,2}\. [^\n,]+")
time_pattern = re.compile(r"no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')
    if not file or not file.filename.lower().endswith('.docx'):
        return 'Please upload a .docx file', 400
    doc = Document(file)

    # 1) Extract date & time
    full_text = "\n".join(p.text for p in doc.paragraphs)
    dm = date_pattern.search(full_text)
    tm = time_pattern.search(full_text)
    date_str = dm.group().strip() if dm else 'N/A'
    time_str = f"{tm.group(1)}–{tm.group(2)}" if tm else 'N/A'
    full_dt = f"{date_str} {time_str}"

    # 2) Participants parsing: tables first
    participants = []
    current_attendance = ''
    # Flow through document elements to capture attendance and tables
    for child in doc.element.body.iterchildren():
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            para = Paragraph(child, doc)
            text = para.text.strip()
            # detect attendance headings
            if 'Piedalīties' in text:
                low = text.lower()
                if 'klātienē' in low:
                    current_attendance = 'klātienē'
                elif 'attālināti' in low:
                    current_attendance = 'attālināti'
        elif tag == 'tbl':
            table = Table(child, doc)
            for row in table.rows:
                num = row.cells[0].text.strip()
                if not re.match(r'^\d+\.\d+', num):
                    continue
                seg = row.cells[1].text.strip()
                parts = seg.split()
                deg = parts[0] if parts and parts[0].lower() in DEGREES else ''
                rem = seg[len(deg):].strip() if deg else seg
                if ',' in rem:
                    name, job = map(str.strip, rem.split(',',1))
                else:
                    name, job = rem, ''
                participants.append({
                    'date': full_dt,
                    'degree': deg,
                    'participant': name,
                    'pjob': job,
                    'attendance': current_attendance
                })
    # 3) Fallback paragraph-based extraction if no participants from tables
    if not participants:
        paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        start = next((i for i,v in enumerate(paras) if 'deleģētas' in v), None)
        end = next((i for i,v in enumerate(paras) if 'vadīs' in v), None)
        block = paras[start+1:end] if start is not None and end is not None else paras
        for para_text in block:
            segs = re.split(r';', para_text)
            for seg in segs:
                seg = seg.strip().rstrip('.;')
                if not seg:
                    continue
                first = seg.split()[0].lower()
                if first in DEGREES:
                    parts = seg.split()
                    deg = parts[0]
                    rem = seg[len(deg):].strip()
                    if ',' in rem:
                        name, job = map(str.strip, rem.split(',',1))
                    else:
                        name, job = rem, ''
                    participants.append({
                        'date': full_dt,
                        'degree': deg,
                        'participant': name,
                        'pjob': job,
                        'attendance': ''
                    })

    # 4) Lecturers extraction
    lecturers = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if 'vadīs' not in text:
            continue
        tail = re.split(r'vadīs[:\-]?', text,1, flags=re.IGNORECASE)[-1]
        for seg in re.split(r';|\s+un\s+', tail):
            ent = seg.strip().rstrip('.;')
            if not ent:
                continue
            if ',' in ent:
                name, job = map(str.strip, ent.split(',',1))
            else:
                m = NAME_RE.search(ent)
                if m:
                    name = f"{m.group(1)} {m.group(2)}"
                    job = ent.replace(name,'').lstrip(', ').strip()
                else:
                    name, job = ent, ''
            lecturers.append({'lecturer':name,'ljob':job})
    # dedupe lecturers
    seen = set(); uniq=[]
    for lec in lecturers:
        key=(lec['lecturer'],lec['ljob'])
        if key not in seen:
            seen.add(key); uniq.append(lec)
    lecturers = uniq

    # 5) Build DataFrame rows
    rows = []
    for i, p in enumerate(participants):
        lec = lecturers[i] if i < len(lecturers) else {'lecturer':'','ljob':''}
        rows.append({
            'Date': p.get('date','') if i==0 else '',
            'Degree': p.get('degree',''),
            'Participant': p.get('participant',''),
            'Participant Job': p.get('pjob',''),
            'Attendance': p.get('attendance',''),
            'Lecturer': lec.get('lecturer',''),
            'Lecturer Job': lec.get('ljob','')
        })
    df = pd.DataFrame(rows)

    # 6) Export to Excel
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        ws = writer.sheets['Data']
        for idx, col in enumerate(df.columns,1):
            w = max(df[col].astype(str).map(len).max(), len(col)) + 2
            ws.column_dimensions[get_column_letter(idx)].width = w
    out.seek(0)
    return send_file(
        out,
        download_name='seminar_data.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT',10000)), debug=True)
