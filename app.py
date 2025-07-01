from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from itertools import zip_longest
from openpyxl.utils import get_column_letter

app = Flask(__name__)

DEGREES = {
    "ierindnieks","ierindniece","kaprālis","kaprāliene","seržants","seržante",
    "virsseržants","virsseržante","virsniekvietnieks","virsniekvietniece",
    "leitnants","leitnante","virsleitnants","virsleitnante","kapteinis","kapteine",
    "majors","majore","pulkvežleitnants","pulkvežleitnante","pulkvedis","pulkvede",
    "ģenerālis","ģenerāle"
}

# Simple two-word name matcher (for fallback)
NAME_RE = re.compile(r'\b([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\s+([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\b')

NUM_ONLY = re.compile(r'^\d+\.\d+\.$')
DATE_RE = re.compile(r'202\d\. gada \d{1,2}\. [^\n,]+')
TIME_RE = re.compile(r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})')

def extract_data(doc):
    # 1) Date & time
    full_text = "\n".join(p.text for p in doc.paragraphs)
    d = DATE_RE.search(full_text)
    t = TIME_RE.search(full_text)
    date_str = d.group().strip() if d else "N/A"
    time_str = f"{t.group(1)}–{t.group(2)}" if t else "N/A"
    full_dt = f"{date_str} {time_str}"

    # 2) Collect participant lines
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    start = next((i for i,v in enumerate(paras) if "deleģētas" in v), None)
    end = None
    if start is not None:
        for j in range(start+1, len(paras)):
            if "vadīs" in paras[j]:
                end = j
                break
    seg_lines = paras[start+1:end] if start is not None and end else paras[:]
    # include table cells
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if txt:
                    seg_lines.append(txt)

    # normalize to entries
    entries = []
    i = 0
    while i < len(seg_lines):
        ln = seg_lines[i]
        if NUM_ONLY.match(ln):
            k = i+1
            while k < len(seg_lines) and not seg_lines[k].strip():
                k += 1
            if k < len(seg_lines):
                entries.append(seg_lines[k].strip())
            i = k+1
            continue
        fw = ln.split()[0].lower() if ln.split() else ""
        if fw in DEGREES:
            entries.append(ln)
        i += 1

    # parse participants
    participants = []
    for seg in entries:
        parts = seg.split()
        deg = parts[0] if parts and parts[0].lower() in DEGREES else ""
        rem = seg[len(deg):].strip() if deg else seg
        if ',' in rem:
            name, job = map(str.strip, rem.split(',',1))
        else:
            name, job = rem, ""
        participants.append({"degree":deg, "participant":name, "pjob":job})

    # 3) Extract lecturers
    lecturers = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if "vadīs" not in txt:
            continue
        # take the part after 'vadīs'
        tail = re.split(r'vadīs[:\-]?', txt, 1, flags=re.IGNORECASE)[-1]
        # split multiple lecturers
        for seg in re.split(r';|\s+un\s+', tail):
            seg = seg.strip().rstrip('.;')
            if not seg:
                continue
            if ',' in seg:
                name, job = map(str.strip, seg.split(',',1))
            else:
                # fallback: try two-word match
                m = NAME_RE.search(seg)
                if m:
                    name = f"{m.group(1)} {m.group(2)}"
                    job = seg.replace(name, "").lstrip(", ").strip()
                else:
                    name, job = seg, ""
            lecturers.append({"lecturer":name, "ljob":job})
    # dedupe while preserving order
    seen = set()
    uniq = []
    for lec in lecturers:
        key = (lec["lecturer"], lec["ljob"])
        if key not in seen:
            seen.add(key)
            uniq.append(lec)
    lecturers = uniq

    # 4) Build rows
    rows = []
    for idx, (p, l) in enumerate(zip_longest(participants, lecturers, fillvalue={})):
        rows.append({
            "Date": full_dt if idx==0 else "",
            "Degree": p.get("degree",""),
            "Participant": p.get("participant",""),
            "Participant Job": p.get("pjob",""),
            "Lecturer": l.get("lecturer",""),
            "Lecturer Job": l.get("ljob",""),
        })
    df = pd.DataFrame(rows)

    # 5) Export to Excel
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
        ws = writer.sheets["Data"]
        for ci, col in enumerate(df.columns, 1):
            mx = df[col].astype(str).map(len).max()
            ws.column_dimensions[get_column_letter(ci)].width = max(mx, len(col)) + 2
    out.seek(0)
    return out

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".docx"):
        return "Please upload a .docx file", 400
    doc = Document(f)
    excel_io = extract_data(doc)
    return send_file(
        excel_io,
        download_name="seminar_data.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",10000)), debug=True)
