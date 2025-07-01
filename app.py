from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# dienesta pakapes (viriesu un sieviesu dzimte)
DEGREES = [
    "ierindnieks", "ierindniece",
    "kaprālis", "kaprāle",
    "seržants", "seržante",
    "virsseržants", "virsseržante",
    "virsniekvietnieks", "virsniekvietniece",
    "leitnants", "leitnante",
    "virsleitnants", "virsleitnante",
    "kapteinis", "kapteine",
    "majors", "majore",
    "pulkvežleitnants", "pulkvežleitnante",
    "pulkvedis", "pulkvede",
    "ģenerālis", "ģenerāle"
]
DEGREE_PATTERN = r'\b(?:' + '|'.join(map(re.escape, DEGREES)) + r')\b'
 #izvelk datumu no dokumenta (ludzu salabo sito rainer man paliek zaebal)
def extract_data(doc):
    first_text = ""
    for p in doc.paragraphs:
        if re.search(r'202\d\. gada', p.text):
            first_text = p.text
            break
    date = (re.search(r'202\d\. gada \d{1,2}\. [a-zāēūī]+', first_text) or re.search(r'202\d\. gada \d{1,2}\. [A-Za-z]+', first_text))
    time = re.search(
        r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})',
        first_text
    )
    dt = f"{date.group() if date else 'N/A'} {time.group(1)+'–'+time.group(2) if time else 'N/A'}"

    # 2) Participants: look for table first
    participants = []
    # find where "deleģētas" occurs
    start_idx = next((i for i,p in enumerate(doc.paragraphs) if "deleģētas" in p.text), None)

    if doc.tables and start_idx is not None:
        # parse each cell in first table
        table = doc.tables[0]
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                if not re.search(DEGREE_PATTERN, text, re.IGNORECASE):
                    continue
                # split on semicolons or newlines
                segments = re.split(r';|\n', text)
                for seg in segments:
                    seg = seg.strip()
                    if not seg or not re.search(DEGREE_PATTERN, seg, re.IGNORECASE):
                        continue
                    # degree: first matching keyword
                    deg = re.search(DEGREE_PATTERN, seg, re.IGNORECASE).group()
                    # name: two capitalized words
                    nm = (re.search(r'([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+(?:\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)+))', seg)
                          or re.search(r'\b([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\b', seg))
                    name = nm.group(1) if nm else ""
                    # job = text after "name,"
                    job = seg.split(f"{name},",1)[-1].strip(" .") if name and f"{name}," in seg else seg
                    participants.append({"degree": deg, "participant": name, "pjob": job})
    else:
        # fallback: paragraphs between “deleģētas” and lecturers
        if start_idx is not None:
            for para in doc.paragraphs[start_idx+1:]:
                text = para.text.strip()
                if not text or "Mācību semināru vadīs" in text:
                    break
                segments = re.split(r';', text)
                # get bold names
                bolds, buf = [], ""
                for run in para.runs:
                    if run.bold:
                        buf += run.text
                    elif buf:
                        bolds.append(buf.strip()); buf = ""
                if buf: bolds.append(buf.strip())
                idx = 0
                for seg in segments:
                    seg = seg.strip()
                    if not re.search(DEGREE_PATTERN, seg, re.IGNORECASE):
                        continue
                    deg = re.search(DEGREE_PATTERN, seg, re.IGNORECASE).group()
                    name = bolds[idx] if idx < len(bolds) else ""
                    idx += 1
                    job = seg.split(f"{name},",1)[-1].strip(" .") if name and f"{name}," in seg else seg
                    participants.append({"degree": deg, "participant": name, "pjob": job})

    # 3) Lecturers
    lecturers = []
    for para in doc.paragraphs:
        if "Mācību semināru vadīs" in para.text:
            tail = para.text.split("Mācību semināru vadīs",1)[-1].strip("–: ")
            parts = re.split(r'\s+un\s+|,\s*', tail)
            for part in parts:
                part = part.strip().rstrip(". ")
                m = re.search(r'([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+(?:\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)+))', part)
                if m:
                    nm = m.group(1)
                    jb = part.replace(nm,"").lstrip(", ").strip()
                else:
                    nm = "—"
                    jb = part
                lecturers.append({"lecturer": nm, "ljob": jb})
            break

    # 4) Build DataFrame
    rows = []
    length = max(len(participants), len(lecturers))
    for i in range(length):
        rows.append({
            "Datums un laiks": dt if i == 0 else "",
            "Dalībnieka dienesta pakāpe": participants[i]["degree"] if i < len(participants) else "",
            "Dalībnieks": participants[i]["participant"] if i < len(participants) else "",
            "Dalībnieka amats": participants[i]["pjob"] if i < len(participants) else "",
            "Lektors": lecturers[i]["lecturer"] if i < len(lecturers) else "",
            "Lektora amats": lecturers[i]["ljob"] if i < len(lecturers) else "",
        })

    return pd.DataFrame(rows)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f or not f.filename.lower().endswith('.docx'):
        return 'Lūdzu augšupielādējiet .docx failu.', 400

    doc = Document(f)
    df = extract_data(doc)

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
