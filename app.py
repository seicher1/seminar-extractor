from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# List of degree keywords including masculine and feminine forms
DEGREES = [
    "ierindnieks", "ierindniece",
    "kaprālis", "kaprāliene",
    "seržants", "seržante",
    "virsseržants", "virsseržante",
    "virsnieka vietnieks", "virsnieces vietniece",
    "leitnants", "leitnante",
    "virsleitnants", "virsleitnante",
    "kapteinis", "kapteiniene",
    "majors", "majore",
    "pulkvežleitnants", "pulkvežleitnante",
    "pulkvedis", "pulkvede",
    "ģenerālis", "ģenerāliene"
]
DEGREE_PATTERN = r"\b(?:" + "|".join(map(re.escape, DEGREES)) + r")\b"


def extract_data(doc):
    # 1) Extract date & time from the first paragraph containing a date
    first_text = next((p.text for p in doc.paragraphs if re.search(r'202\d\. gada', p.text)), "")
    date_match = re.search(r'202\d\. gada \d{1,2}\. [a-zāēūī]+', first_text)
    time_match = re.search(
        r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})',
        first_text
    )
    date = date_match.group() if date_match else "N/A"
    time = f"{time_match.group(1)}–{time_match.group(2)}" if time_match else "N/A"
    full_dt = f"{date} {time}"

    # 2) Participants between 'deleģētas' and 'Mācību semināru vadīs'
    start_idx = next((i for i, p in enumerate(doc.paragraphs) if "deleģētas" in p.text), None)
    end_idx = next((i for i, p in enumerate(doc.paragraphs) if "Mācību semināru vadīs" in p.text), None)

    participants = []
    if start_idx is not None and end_idx is not None:
        for para in doc.paragraphs[start_idx+1:end_idx]:
            text = para.text.strip()
            if not text:
                continue
            # reconstruct bold names in this paragraph
            bolds, buf = [], ""
            for run in para.runs:
                if run.bold:
                    buf += run.text
                elif buf:
                    bolds.append(buf.strip())
                    buf = ""
            if buf:
                bolds.append(buf.strip())

            # split segments by semicolon
            segments = [seg.strip() for seg in text.split(';') if seg.strip()]
            bindex = 0
            for seg in segments:
                if not re.search(DEGREE_PATTERN, seg, re.IGNORECASE):
                    continue
                deg_m = re.search(DEGREE_PATTERN, seg, re.IGNORECASE)
                degree = deg_m.group(0) if deg_m else ""
                name = bolds[bindex] if bindex < len(bolds) else ""
                bindex += 1
                job = seg.split(f"{name},", 1)[-1].strip(" .") if name and f"{name}," in seg else seg
                participants.append({"degree": degree, "participant": name, "pjob": job})

    # 3) Lecturers extraction
    lecturers = []
    for para in doc.paragraphs:
        if "Mācību semināru vadīs" in para.text:
            tail = para.text.split("Mācību semināru vadīs", 1)[-1].strip("–: ")
            parts = re.split(r'\s+un\s+|,\s*', tail)
            for part in parts:
                part = part.strip().rstrip(". ")
                # fixed regex: no extra parenthesis
                m = re.search(r"([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][a-zāčēģīķļņŗšūž]+\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][a-zāčēģīķļņŗšūž]+)", part)
                if m:
                    nm = m.group(1)
                    jb = part.replace(nm, "").lstrip(", ").strip()
                else:
                    nm = "—"
                    jb = part
                lecturers.append({"lecturer": nm, "ljob": jb})
            break

    # 4) Assemble DataFrame
    rows = []
    max_len = max(len(participants), len(lecturers))
    for i in range(max_len):
        rows.append({
            "Date": full_dt if i == 0 else "",
            "Degree": participants[i]["degree"] if i < len(participants) else "",
            "Participant": participants[i]["participant"] if i < len(participants) else "",
            "Participant Job": participants[i]["pjob"] if i < len(participants) else "",
            "Lecturer": lecturers[i]["lecturer"] if i < len(lecturers) else "",
            "Lecturer Job": lecturers[i]["ljob"] if i < len(lecturers) else "",
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
