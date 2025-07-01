from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from itertools import zip_longest
from openpyxl.utils import get_column_letter

# Required imports for parsing order of elements
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

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

# Fallback name matcher
NAME_RE = re.compile(r"\b([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\s+([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\b")

# Date/time patterns
date_pattern = re.compile(r"202\d\. gada \d{1,2}\. [^\n,]+")
time_pattern = re.compile(r"no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f or not f.filename.lower().endswith('.docx'):
        return 'Please upload a .docx file', 400
    doc = Document(f)

    # 1) Extract date & time
    full_text = "\n".join(p.text for p in doc.paragraphs)
    dm = date_pattern.search(full_text)
    tm = time_pattern.search(full_text)
    date_str = dm.group().strip() if dm else 'N/A'
    time_str = f"{tm.group(1)}–{tm.group(2)}" if tm else 'N/A'
    full_dt = f"{date_str} {time_str}"

    # 2) Extract participants in document order using attendance headings + tables
    participants = []
    current_attendance = ''
    # Iterate body elements to preserve order
    for child in doc.element.body.iterchildren():
        tag = child.tag.split('}')[1]
        if tag == 'p':
            para = Paragraph(child, doc)
            text = para.text.strip()
            # Check for attendance headings
            m = re.match(r"^\d+\.\s*Piedalīties.*\((klātienē|attālināti)\)", text, re.IGNORECASE)
            if m:
                current_attendance = m.group(1).lower()
        elif tag == 'tbl':
            table = Table(child, doc)
            for row in table.rows:
                # Expect numbering in first cell
                num = row.cells[0].text.strip()
                if not re.match(r"^\d+\.\d+", num):
                    continue
                seg = row.cells[1].text.strip()
                # Degree = first word if in DEGREES
                parts = seg.split()
                deg = parts[0] if parts and parts[0].lower() in DEGREES else ''
                # Name: substring between degree and comma
                rem = seg[len(deg):].strip() if deg else seg
                if ',' in rem:
                    name, job = map(str.strip, rem.split(',',1))
                else:
                    name, job = rem, ''
                participants.append({
                    'Date': full_dt,
                    'Degree': deg,
                    'Participant': name,
                    'Participant Job': job,
                    'Attendance': current_attendance
                })

    # 3) Extract lecturers anywhere in doc
    lecturers = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if 'vadīs' not in text:
            continue
        tail = re.split(r'vadīs[:\-]?', text, 1, flags=re.IGNORECASE)[-1]
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
            lecturers.append({'Lecturer': name, 'Lecturer Job': job})
    # Deduplicate lecturers maintaining order
    seen = set(); unique_lect=[]
    for lec in lecturers:
        key=(lec['Lecturer'],lec['Lecturer Job'])
        if key not in seen:
            seen.add(key); unique_lect.append(lec)
    lecturers = unique_lect

    # 4) Combine participants and lecturers
    rows = []
    for i, p in enumerate(participants):
        lec = lecturers[i] if i < len(lecturers) else {'Lecturer':'','Lecturer Job':''}
        rows.append({
            'Date': p['Date'] if i==0 else '',
            'Degree': p['Degree'],
            'Participant': p['Participant'],
            'Participant Job': p['Participant Job'],
            'Attendance': p['Attendance'],
            'Lecturer': lec['Lecturer'],
            'Lecturer Job': lec['Lecturer Job']
        })
    df = pd.DataFrame(rows)

    # 5) Export to Excel
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        ws = writer.sheets['Data']
        for idx, col in enumerate(df.columns, 1):
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
