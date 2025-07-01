from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# Exact degree keywords (m/f)
DEGREES = {
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
}

# Simple name‐matcher: two capitalized words
NAME_RE = re.compile(r'\b([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\s+([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\b')

def extract_data(doc):
    # 1) Extract date & time from anywhere
    full_text = "\n".join(p.text for p in doc.paragraphs)
    date_m = re.search(r'202\d\. gada \d{1,2}\. [^\n,]+', full_text)
    time_m = re.search(
        r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})',
        full_text
    )
    date_str = date_m.group().strip() if date_m else "N/A"
    time_str = f"{time_m.group(1)}–{time_m.group(2)}" if time_m else "N/A"
    full_dt = f"{date_str} {time_str}"

    # 2) Build a list of all relevant lines (paras + table cells) between markers
    lines = []
    paras = [p.text.strip() for p in doc.paragraphs]
    start = next((i for i,t in enumerate(paras) if "deleģētas" in t), None)
    end1  = next((i for i,t in enumerate(paras) if "Mācību semināru vadīs" in t), None)
    end2  = next((i for i,t in enumerate(paras) if "Mācības vadīs" in t), None)
    end   = end1 if end1 is not None else end2

    if start is not None and end is not None and end > start:
        lines.extend(paras[start+1:end])
    else:
        lines.extend(paras)

    # add all table cell texts
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if txt:
                    lines.append(txt)

    # 3) Normalize those lines into complete “entries”
    entries = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        # if a lone numbering line "1.1."
        if re.match(r'^\d+\.\d+\.$', line):
            # grab next non-empty line
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines):
                entries.append(lines[j].strip())
            i = j + 1
            continue
        # if this line itself contains a degree keyword, treat as complete
        low = line.lower()
        first_word = low.split()[0] if low else ""
        if first_word in DEGREES:
            entries.append(line)
        i += 1

    # 4) Parse each entry into degree, name, job
    participants = []
    for seg in entries:
        parts = seg.split()
        degree = parts[0] if parts and parts[0].lower() in DEGREES else ""
        # name via NAME_RE
        name = ""
        nm = NAME_RE.search(seg)
        if nm:
            name = f"{nm.group(1)} {nm.group(2)}"
        # job = text after the first comma
        job = ""
        if ',' in seg and name:
            job = seg.split(f"{name},", 1)[1].strip()
        participants.append({
            "degree": degree,
            "participant": name,
            "pjob": job
        })

    # 5) Extract lecturers
    lecturers = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if "Mācību semināru vadīs" in text or "Mācības vadīs" in text:
            tail = re.split(r'Mācību semināru vadīs|Mācības vadīs', text, 1)[-1]
            # split on “ un ” or comma before capital
            parts = re.split(r'\s+un\s+|,\s*(?=[A-ZĀČĒĢĪĶĻŅŖŠŪŽ])', tail)
            for part in parts:
                t = part.strip().rstrip(".;")
                nm = NAME_RE.match(t)
                if nm:
                    lecturer = f"{nm.group(1)} {nm.group(2)}"
                    job = t.replace(lecturer, "").lstrip(", ").strip()
                else:
                    lecturer, job = "—", t
                lecturers.append({"lecturer": lecturer, "ljob": job})
            break

    # 6) Build DataFrame and write to Excel in-memory
    rows = []
    L = max(len(participants), len(lecturers))
    for idx in range(L):
        rows.append({
            "Date": full_dt if idx == 0 else "",
            "Degree": participants[idx]["degree"]    if idx < len(participants) else "",
            "Participant": participants[idx]["participant"] if idx < len(participants) else "",
            "Participant Job": participants[idx]["pjob"] if idx < len(participants) else "",
            "Lecturer": lecturers[idx]["lecturer"] if idx < len(lecturers) else "",
            "Lecturer Job": lecturers[idx]["ljob"]  if idx < len(lecturers) else ""
        })

    df = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        ws = writer.sheets['Data']
        for col_idx, col in enumerate(df.columns, 1):
            width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            ws.column_dimensions[get_column_letter(col_idx)].width = width
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
