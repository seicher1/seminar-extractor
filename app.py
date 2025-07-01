from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# Degrees (m+f)
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

# Build a non-capturing group for all degrees
deg_group = r'(?:' + '|'.join(re.escape(d) for d in DEGREES) + r')'

# One big raw regex:
#   \d+\.\d+\.      → 1.x.
#   \s*<degree>\s+  → your degree keyword
#   ([A-Z...]+)     → capture Name Surname
#   \s*,\s*([^;]+)  → comma + capture job until semicolon
PART_PATTERN = re.compile(
    r'\d+\.\d+\.\s*' +
    deg_group +
    r'\s+' +
    r'([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+(?:\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)+)' +
    r'\s*,\s*' +
    r'([^;]+)',
    re.IGNORECASE
)

def extract_data(doc):
    # 1) Date & Time
    all_text = "\n".join(p.text for p in doc.paragraphs)
    date_m = re.search(r'202\d\. gada \d{1,2}\. [^\n,]+', all_text)
    time_m = re.search(
        r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})',
        all_text
    )
    date = date_m.group().strip() if date_m else "N/A"
    time = f"{time_m.group(1)}–{time_m.group(2)}" if time_m else "N/A"
    full_dt = f"{date} {time}"

    # 2) Extract block between "deleģētas" and lecturers marker
    texts = [p.text for p in doc.paragraphs]
    start = next((i for i,t in enumerate(texts) if "deleģētas" in t), None)
    end1  = next((i for i,t in enumerate(texts) if "Mācību semināru vadīs" in t), None)
    end2  = next((i for i,t in enumerate(texts) if "Mācības vadīs" in t), None)
    end   = end1 if end1 is not None else end2
    block = ""
    if start is not None and end is not None and end > start:
        block = " ".join(texts[start+1:end])
    else:
        block = all_text

    # 3) Participants via regex
    participants = []
    for m in PART_PATTERN.finditer(block):
        participants.append({
            "degree": m.group(1),
            "participant": m.group(2),
            "pjob": m.group(3).strip()
        })

    # 4) Lecturers
    lecturers = []
    for p in doc.paragraphs:
        if "Mācību semināru vadīs" in p.text or "Mācības vadīs" in p.text:
            tail = re.split(r'Mācību semināru vadīs|Mācības vadīs', p.text, 1)[-1]
            parts = re.split(r'\s+un\s+|,\s*(?=[A-ZĀČĒĢĪĶĻŅŖŠŪŽ])', tail)
            for part in parts:
                text = part.strip().rstrip(".;")
                m = re.match(
                    r'([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+(?:\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)+))',
                    text
                )
                if m:
                    nm = m.group(1)
                    jb = text.replace(nm, "").lstrip(", ").strip()
                else:
                    nm, jb = "—", text
                lecturers.append({"lecturer": nm, "ljob": jb})
            break

    # 5) Assemble DataFrame with auto-fit columns later
    rows = []
    length = max(len(participants), len(lecturers))
    for i in range(length):
        rows.append({
            "Date": full_dt if i == 0 else "",
            "Degree": participants[i]["degree"]    if i < len(participants) else "",
            "Participant": participants[i]["participant"] if i < len(participants) else "",
            "Participant Job": participants[i]["pjob"] if i < len(participants) else "",
            "Lecturer": lecturers[i]["lecturer"] if i < len(lecturers) else "",
            "Lecturer Job": lecturers[i]["ljob"]  if i < len(lecturers) else ""
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
    df  = extract_data(doc)

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
    app.run(
        host='0.0.0.0',
        port=int(os.environ.get('PORT', 10000)),
        debug=True
    )
