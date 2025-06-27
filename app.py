from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from openpyxl.utils import get_column_letter

app = Flask(__name__)

def extract_data(doc):
    # 1) datums un laiks no 6. rindkopas (it just works for some reason) (es velak uzlabosu)
    date_para = doc.paragraphs[6].text.strip()
    date_match = re.search(r'202\d\. gada \d{1,2}\. [a-zāēūī]+', date_para)
    time_match = re.search(
        r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})',
        date_para
    )
    date = date_match.group() if date_match else "N/A"
    time = f"{time_match.group(1)}–{time_match.group(2)}" if time_match else "N/A"
    full_dt = f"{date} {time}"

    # atrod vardu delegetas un sak meklet dalibniekus no turienes
    start_idx = None
    for i, para in enumerate(doc.paragraphs):
        if "deleģētas" in para.text:
            start_idx = i + 1
            break

    participants = []
    if start_idx is not None:
        # vakt datus kamer apstajas pie lektoriem
        for para in doc.paragraphs[start_idx:]:
            text = para.text.strip()
            if not text:
                continue
            # apstaties pie lektoriem
            if "Mācību semināru vadīs" in text:
                break

            # bold to normal
            bolds, buf = [], ""
            for run in para.runs:
                if run.bold:
                    buf += run.text
                elif buf:
                    bolds.append(buf.strip())
                    buf = ""
            if buf:
                bolds.append(buf.strip())

            # sadala semikolus ta lai vnk labak
            segs = [s.strip() for s in re.split(r';', text) if s.strip()]
            for i, seg in enumerate(segs):
                deg = seg.split()[0] if seg.split() else "N/A"
                name = bolds[i] if i < len(bolds) else "N/A"
                job = seg.split(f"{name},",1)[-1].strip(" .") if f"{name}," in seg else "N/A"
                participants.append({"degree": deg, "participant": name, "pjob": job})

    # lektori time
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

    # uztaisa kolonnas un rindinas prieks excel
    rows = []
    max_len = max(len(participants), len(lecturers))
    for i in range(max_len):
        rows.append({
            "Datums un laiks": full_dt if i==0 else "",
            "Pakāpe": participants[i]["degree"] if i < len(participants) else "",
            "Dalībnieks": participants[i]["participant"] if i < len(participants) else "",
            "Dalībnieka amats": participants[i]["pjob"] if i < len(participants) else "",
            "Lektors/i": lecturers[i]["lecturer"] if i < len(lecturers) else "",
            "Lekotra amats/i": lecturers[i]["ljob"] if i < len(lecturers) else "",
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
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            ws.column_dimensions[get_column_letter(idx)].width = max_len
    out.seek(0)

    return send_file(
        out,
        download_name='seminar_data.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
#prieks render
if __name__ == '__main__':
    app.run(
        host='0.0.0.0',
        port=int(os.environ.get('PORT', 10000)),
        debug=True
    )
