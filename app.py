from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from openpyxl.utils import get_column_letter

app = Flask(__name__)

def extract_data(doc):
    # 1) Date & Time from paragraph 6
    date_para = doc.paragraphs[6].text.strip()
    date_match = re.search(r'202\d\. gada \d{1,2}\. [a-zāēūī]+', date_para)
    time_match = re.search(
        r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})',
        date_para
    )
    date = date_match.group() if date_match else "N/A"
    time = f"{time_match.group(1)}–{time_match.group(2)}" if time_match else "N/A"
    full_dt = f"{date} {time}"

    # 2) Participants — first try the first table after the “deleģētas” line
    participants = []

    # find the index of the paragraph that contains “deleģētas”
    start_idx = next((i for i,p in enumerate(doc.paragraphs)
                      if "deleģētas" in p.text), None)

    # if there's at least one table, assume it's the participants table
    if doc.tables and start_idx is not None:
        table = doc.tables[0]
        for row in table.rows:
            # join all cells in row into one text blob
            row_text = " ".join(cell.text.strip() for cell in row.cells).strip()
            # look for lines starting with numbering 1.x.
            for match in re.finditer(r'(\d\.\d+)\.\s*(.+)', row_text):
                # content after the numbering
                content = match.group(2).strip()
                # degree = first word
                parts = content.split(',')
                # first segment has “Degree Name” separated by space(s)
                lead = parts[0].strip()
                deg = lead.split()[0]
                # name = the next one or two capitalized words
                name_match = re.match(r'\w+\s+([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+(?:\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)?)', content)
                name = name_match.group(1) if name_match else "N/A"
                # job = everything after the comma that follows the name
                job = content.split(f"{name},",1)[-1].strip() if f"{name}," in content else ""
                participants.append({
                    "degree": deg,
                    "participant": name,
                    "pjob": job
                })
    else:
        # fallback to old paragraph logic
        for para in doc.paragraphs[10:]:
            text = para.text.strip()
            if not text:
                continue
            if "Mācību semināru vadīs" in text:
                break
            bolds, buf = [], ""
            for run in para.runs:
                if run.bold:
                    buf += run.text
                elif buf:
                    bolds.append(buf.strip())
                    buf = ""
            if buf:
                bolds.append(buf.strip())
            segs = [s.strip() for s in text.split(";") if s.strip()]
            for i, seg in enumerate(segs):
                deg = seg.split()[0] if seg.split() else "N/A"
                nm = bolds[i] if i < len(bolds) else "N/A"
                jb = seg.split(f"{nm},",1)[-1].strip(" .") if f"{nm}," in seg else ""
                participants.append({"degree": deg, "participant": nm, "pjob": jb})
            if text.endswith("."):
                break

    # 3) Lecturers
    lecturers = []
    for para in doc.paragraphs:
        if "Mācību semināru vadīs" in para.text:
            line = para.text.split("Mācību semināru vadīs",1)[-1].strip("–: ")
            parts = re.split(r'\s+un\s+|,\s*', line)
            for part in parts:
                part = part.strip().rstrip(".")
                m = re.search(
                    r'([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][a-zāčēģīķļņŗšūž]+\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][a-zāčēģīķļņŗšūž]+)',
                    part
                )
                if m:
                    nm = m.group(1)
                    jb = part.replace(nm,"").lstrip(", ").strip()
                else:
                    nm = "—"
                    jb = part
                lecturers.append({"lecturer": nm, "ljob": jb})
            break

    # 4) Assemble into rows
    rows = []
    max_len = max(len(participants), len(lecturers))
    for i in range(max_len):
        rows.append({
            "Date": full_dt if i==0 else "",
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

    # export to Excel with auto column widths
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
