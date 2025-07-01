from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# Correct degree keywords (masculine & feminine)
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

# Build non-capturing group for degrees
deg_group = r'(?:' + '|'.join(map(re.escape, DEGREES)) + r')'

# 1) Pattern to grab every “1.x.” or “2.x.” segment up to semicolon:
#    Group 1 = degree, Group 2 = the rest of the segment
ENTRY_PATTERN = re.compile(
    r'\d+\.\d+\.\s*(' + deg_group + r')\s*([^;]+)',
    re.IGNORECASE
)

# 2) Name matcher (two capitalized words)
NAME_PATTERN = re.compile(
    r'([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+(?:\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)+)',
    re.IGNORECASE
)

def extract_data(doc):
    # ── A) Extract full text and date/time ──────────────────────
    full_text = "\n".join(p.text for p in doc.paragraphs)
    date_m = re.search(r'202\d\. gada \d{1,2}\. [^\n,]+', full_text)
    time_m = re.search(
        r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})',
        full_text
    )
    date = date_m.group().strip() if date_m else "N/A"
    time = f"{time_m.group(1)}–{time_m.group(2)}" if time_m else "N/A"
    full_dt = f"{date} {time}"

    # ── B) Extract block between “deleģētas” and lecturers marker ──
    texts = [p.text for p in doc.paragraphs]
    start = next((i for i,t in enumerate(texts) if "deleģētas" in t), None)
    end1  = next((i for i,t in enumerate(texts) if "Mācību semināru vadīs" in t), None)
    end2  = next((i for i,t in enumerate(texts) if "Mācības vadīs" in t), None)
    end   = end1 if end1 is not None else end2
    if start is not None and end is not None and end > start:
        block = " ".join(texts[start+1:end])
    else:
        block = full_text  # fallback

    # ── C) Parse participants via ENTRY_PATTERN ────────────────
    participants = []
    for m in ENTRY_PATTERN.finditer(block):
        degree = m.group(1).strip()
        seg_rest = m.group(2).strip()
        # Extract name
        name_m = NAME_PATTERN.search(seg_rest)
        name = name_m.group(1).strip() if name_m else ""
        # Extract job = text after the name comma
        job = seg_rest.split(',',1)[1].strip() if ',' in seg_rest else ""
        participants.append({
            "degree": degree,
            "participant": name,
            "pjob": job
        })

    # ── D) Parse lecturers ───────────────────────────────────────
    lecturers = []
    for p in doc.paragraphs:
        if "Mācību semināru vadīs" in p.text or "Mācības vadīs" in p.text:
            tail = re.split(r'Mācību semināru vadīs|Mācības vadīs', p.text, 1)[-1]
            parts = re.split(r'\s+un\s+|,\s*(?=[A-ZĀČĒĢĪĶĻŅŖŠŪŽ])', tail)
            for part in parts:
                t = part.strip().rstrip(".;")
                # Reuse NAME_PATTERN to capture lecturer name
                nm_m = NAME_PATTERN.match(t)
                if nm_m:
                    nm = nm_m.group(1)
                    jb = t.replace(nm, "").lstrip(", ").strip()
                else:
                    nm, jb = "—", t
                lecturers.append({"lecturer": nm, "ljob": jb})
            break

    # ── E) Build rows and return DataFrame ───────────────────────
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
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            ws.column_dimensions[get_column_letter(idx)].width = max_len
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
