from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# Your exact degree forms
DEGREES = [
    "ierindnieks","ierindniece",
    "kaprālis","kaprāliene",
    "seržants","seržante",
    "virsseržants","virsseržante",
    "virsniekvietnieks","virsniekvietniece",
    "leitnants","leitnante",
    "virsleitnants","virsleitnante",
    "kapteinis","kapteine",
    "majors","majore",
    "pulkvežleitnants","pulkvežleitnante",
    "pulkvedis","pulkvede",
    "ģenerālis","ģenerāle"
]
# Degree lookup lowercase
DEG_SET = {d.lower() for d in DEGREES}

# Simplest name matcher: two capitalized words
NAME_RE = re.compile(r'([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+(?:\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)+))')

def extract_data(doc):
    # –– 1) Date & Time anywhere in doc ––
    full_text = "\n".join(p.text for p in doc.paragraphs)
    d_m = re.search(r'202\d\. gada \d{1,2}\. [^\n,]+', full_text)
    t_m = re.search(r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})', full_text)
    date = d_m.group().strip() if d_m else "N/A"
    time = f"{t_m.group(1)}–{t_m.group(2)}" if t_m else "N/A"
    full_dt = f"{date} {time}"

    # –– 2) Gather all lines (paras + table cells) between deleģētas → vadīs ––
    texts = [p.text.strip() for p in doc.paragraphs]
    start = next((i for i,t in enumerate(texts) if "deleģētas" in t), None)
    end1  = next((i for i,t in enumerate(texts) if "Mācību semināru vadīs" in t), None)
    end2  = next((i for i,t in enumerate(texts) if "Mācības vadīs" in t), None)
    end   = end1 if end1 is not None else end2
    block_lines = []
    if start is not None and end is not None and end > start:
        # paragraphs
        block_lines = texts[start+1:end]
    else:
        block_lines = texts  # fallback full doc

    # include every table cell too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if txt:
                    block_lines.append(txt)

    # –– 3) Normalize & pair numbering with next line ––
    entries = []
    skip = False
    for idx, line in enumerate(block_lines):
        if skip:
            skip = False
            continue
        # if line is just "1.1." or "2.3." etc
        if re.match(r'^\d+\.\d+\.$', line):
            # pair with next non-empty
            for j in range(idx+1, len(block_lines)):
                nxt = block_lines[j].strip()
                if nxt:
                    entries.append(nxt)
                    skip = True
                    break
        else:
            # if line contains degree keyword, treat as standalone segment
            low = line.lower()
            if any(d in low for d in DEG_SET):
                entries.append(line)

    # –– 4) Parse each entry for degree, name, job ––
    participants = []
    for seg in entries:
        parts = seg.split()
        deg = ""
        if parts and parts[0].lower() in DEG_SET:
            deg = parts[0]
        # find name
        nm = ""
        m = NAME_RE.search(seg)
        if m:
            nm = m.group(1)
        # job = after comma
        jb = ""
        if ',' in seg and nm:
            jb = seg.split(f"{nm},",1)[-1].strip()
        participants.append({"degree": deg, "participant": nm, "pjob": jb})

    # –– 5) Lecturers (unchanged) ––
    lecturers = []
    for p in doc.paragraphs:
        if "Mācību semināru vadīs" in p.text or "Mācības vadīs" in p.text:
            tail = re.split(r'Mācību semināru vadīs|Mācības vadīs', p.text,1)[-1]
            parts = re.split(r'\s+un\s+|,\s*(?=[A-ZĀČĒĢĪĶĻŅŖŠŪŽ])', tail)
            for part in parts:
                t = part.strip().rstrip(".;")
                m = NAME_RE.match(t)
                if m:
                    nm = m.group(1)
                    jb = t.replace(nm,"").lstrip(", ").strip()
                else:
                    nm, jb = "—", t
                lecturers.append({"lecturer": nm, "ljob": jb})
            break

    # –– 6) Build DataFrame with auto‐fit columns ––
    rows = []
    L = max(len(participants), len(lecturers))
    for i in range(L):
        rows.append({
            "Date": full_dt if i==0 else "",
            "Degree": participants[i]["degree"]    if i<L else "",
            "Participant": participants[i]["participant"] if i<L else "",
            "Participant Job": participants[i]["pjob"] if i<L else "",
            "Lecturer": lecturers[i]["lecturer"] if i<L else "",
            "Lecturer Job": lecturers[i]["ljob"]  if i<L else ""
        })
    df = pd.DataFrame(rows)

    # autofit
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as w:
        df.to_excel(w, index=False, sheet_name='Data')
        ws = w.sheets['Data']
        for idx,col in enumerate(df.columns,1):
            mx = df[col].astype(str).map(len).max()
            ws.column_dimensions[get_column_letter(idx)].width = max(mx, len(col)) + 2
    output.seek(0)
    return output

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f or not f.filename.lower().endswith('.docx'):
        return 'Lūdzu augšupielādējiet .docx failu.', 400
    doc = Document(f)
    excel_io = extract_data(doc)
    return send_file(
        excel_io,
        download_name='seminar_data.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(
        host='0.0.0.0',
        port=int(os.environ.get('PORT', 10000)),
        debug=True
    )
