from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from openpyxl.utils import get_column_letter

app = Flask(__name__)

DEGREES = [
    "ierindnieks", "kaprālis", "seržants", "virsseržants",
    "virsnieka vietnieks", "leitnants", "virsleitnants",
    "kapteinis", "majors", "pulkvežleitnants",
    "pulkvedis", "ģenerālis"
]
DEGREE_PATTERN = r'\b(?:' + '|'.join(map(re.escape, DEGREES)) + r')\b'

def extract_data(doc):
    # 1) Extract date & time from paragraph 6
    date_para = doc.paragraphs[6].text.strip()
    date_match = re.search(r'202\d\. gada \d{1,2}\. [a-zāēūī]+', date_para)
    time_match = re.search(
        r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})',
        date_para
    )
    date = date_match.group() if date_match else "N/A"
    time = f"{time_match.group(1)}–{time_match.group(2)}" if time_match else "N/A"
    full_dt = f"{date} {time}"

    # 2) Locate start/end by keywords
    start_idx = next((i for i,p in enumerate(doc.paragraphs)
                      if "deleģētas" in p.text), None)
    end_idx   = next((i for i,p in enumerate(doc.paragraphs)
                      if "Mācību semināru vadīs" in p.text), None)

    participants = []
    if start_idx is not None and end_idx is not None:
        # Loop through the relevant paragraphs
        for para in doc.paragraphs[start_idx+1:end_idx]:
            text = para.text.strip()
            if not text:
                continue

            # 2a) Reconstruct all bold names in this paragraph
            bold_names = []
            buf = ""
            for run in para.runs:
                if run.bold:
                    buf += run.text
                elif buf:
                    bold_names.append(buf.strip())
                    buf = ""
            if buf:
                bold_names.append(buf.strip())

            # 2b) Split on semicolon to get each candidate segment
            segments = [seg.strip() for seg in text.split(';') if seg.strip()]
            name_idx = 0

            for seg in segments:
                # Only process if a known degree appears
                if not re.search(DEGREE_PATTERN, seg, re.IGNORECASE):
                    continue

                # Extract degree (first matching keyword)
                deg_m = re.search(DEGREE_PATTERN, seg, re.IGNORECASE)
                degree = deg_m.group(0).lower() if deg_m else ""

                # Extract the bold name for this segment if available
                name = bold_names[name_idx] if name_idx < len(bold_names) else ""
                name_idx += 1

                # Extract job: text after "name," 
                job = ""
                if name and f"{name}," in seg:
                    job = seg.split(f"{name},", 1)[-1].strip(" .")

                participants.append({
                    "degree": degree,
                    "participant": name,
                    "pjob": job
                })

    # 3) Extract lecturers as before
    lecturers = []
    for para in doc.paragraphs:
        if "Mācību semināru vadīs" in para.text:
            tail = para.text.split("Mācību semināru vadīs",1)[-1].strip("–: ")
            parts = re.split(r'\s+un\s+|,\s*', tail)
            for part in parts:
                p = part.strip().rstrip(".")
                m = re.search(
                    r'([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][a-zāčēģīķļņŗšūž]+\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][a-zāčēģīķļņŗšūž]+)',
                    p
                )
                if m:
                    nm = m.group(1)
                    jb = p.replace(nm,"").lstrip(", ").strip()
                else:
                    nm = "—"
                    jb = p
                lecturers.append({"lecturer": nm, "ljob": jb})
            break

    # 4) Build the final DataFrame aligning participants & lecturers
    rows = []
    max_len = max(len(participants), len(lecturers))
    for i in range(max_len):
        rows.append({
            "Datums un laiks": full_dt if i == 0 else "",
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
            width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            ws.column_dimensions[get_column_letter(idx)].width = width
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
