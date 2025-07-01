from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from itertools import zip_longest
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# Degree keywords (exact forms)
DEGREES = {
    "ierindnieks","ierindniece","kaprālis","kaprāliene",
    "seržants","seržante","virsseržants","virsseržante",
    "virsniekvietnieks","virsniekvietniece","leitnants","leitnante",
    "virsleitnants","virsleitnante","kapteinis","kapteine",
    "majors","majore","pulkvežleitnants","pulkvežleitnante",
    "pulkvedis","pulkvede","ģenerālis","ģenerāle"
}

# Lecturer name fallback matcher
NAME_RE = re.compile(r"\b([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\s+([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\b")

# Date/time patterns
date_pattern = re.compile(r"202\d\. gada \d{1,2}\. [^\n,]+")
time_pattern = re.compile(r"no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})")

# Attendance mapping by table index
ATT_MAP = {
    0: 'klātienē',
    1: 'attālināti',
    2: 'klātienē',
    3: 'attālināti'
}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f or not f.filename.lower().endswith('.docx'):
        return 'Lūdzu augšupielādējiet .docx failu.', 400
    doc = Document(f)

    # 1) Extract date & time
    text = "\n".join(p.text for p in doc.paragraphs)
    dm = date_pattern.search(text)
    tm = time_pattern.search(text)
    date_str = dm.group().strip() if dm else 'N/A'
    time_str = f"{tm.group(1)}–{tm.group(2)}" if tm else 'N/A'
    full_dt = f"{date_str} {time_str}"

    # 2) Parse participants from tables
    participants = []
    for ti, table in enumerate(doc.tables):
        att = ATT_MAP.get(ti, '')
        for row in table.rows:
            num = row.cells[0].text.strip()
            if not re.match(r"^\d+\.\d+", num):
                continue
            seg = row.cells[1].text.strip()
            # degree = first word if in DEGREES
            parts = seg.split()
            deg = parts[0] if parts and parts[0].lower() in DEGREES else ''
            # name: substring between degree and comma
            rem = seg[len(deg):].strip() if deg else seg
            name = rem.split(',',1)[0].strip()
            # job: after comma
            job = seg.split(',',1)[1].strip() if ',' in seg else ''
            participants.append({
                'degree': deg,
                'participant': name,
                'pjob': job,
                'attendance': att
            })

    # 3) Extract lecturers: any paragraph containing 'vadīs'
    lecturers = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if 'vadīs' not in t:
            continue
        tail = re.split(r'vadīs[:\-]?', t, 1, flags=re.IGNORECASE)[-1]
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
                    job = ent.replace(name, '').lstrip(', ').strip()
                else:
                    name, job = ent, ''
            lecturers.append({'lecturer': name, 'ljob': job})
    # dedupe
    seen = set(); uniq = []
    for lec in lecturers:
        key = (lec['lecturer'], lec['ljob'])
        if key not in seen:
            seen.add(key); uniq.append(lec)
    lecturers = uniq

    # 4) Build DataFrame rows
    rows = []
    for i, (p, l) in enumerate(zip_longest(participants, lecturers, fillvalue={})):  
        rows.append({
            'Date': full_dt if i==0 else '',
            'Degree': p.get('degree',''),
            'Participant': p.get('participant',''),
            'Participant Job': p.get('pjob',''),
            'Attendance': p.get('attendance',''),
            'Lecturer': l.get('lecturer',''),
            'Lecturer Job': l.get('ljob','')
        })
    df = pd.DataFrame(rows)

    # 5) Export to Excel
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        ws = writer.sheets['Data']
        for ci, col in enumerate(df.columns, 1):
            mx = df[col].astype(str).map(len).max()
            ws.column_dimensions[get_column_letter(ci)].width = max(mx, len(col)) + 2
    out.seek(0)
    return send_file(out, download_name='seminar_data.xlsx', as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT',10000)), debug=True)
