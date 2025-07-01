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
    # 1) Date & Time
    full_text = "\n".join(p.text for p in doc.paragraphs)
    date_m = DATE_RE.search(full_text)
    time_m = TIME_RE.search(full_text)
    date_str = date_m.group().strip() if date_m else "N/A"
    time_str = f"{time_m.group(1)}–{time_m.group(2)}" if time_m else "N/A"
    full_dt = f"{date_str} {time_str}"

    # 2) Participants & Attendance
    participants = []
    current_attendance = ""
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    i = 0
    while i < len(paras):
        line = paras[i]

        # Detect attendance heading:
        # e.g. "1. Piedalīties mācību seminārā 2025. ... (klātienē) tiek deleģētas..."
        m_att = re.match(
            r"^\d+\.\s*Piedalīties.*\((klātienē|attālināti)\)",
            line,
            re.IGNORECASE
        )
        if m_att:
            current_attendance = m_att.group(1).lower()
            i += 1
            continue

        # Detect a numbering-only line ("1.1.", "2.3.", etc.)
        if NUM_ONLY.match(line):
            entry = paras[i+1] if i+1 < len(paras) else line
            i += 2
        else:
            # If the line itself begins with a degree, treat it as the entry
            first = line.split()[0].lower()
            if first in DEGREES:
                entry = line
                i += 1
            else:
                i += 1
                continue

        # Parse the entry into degree, name, job
        parts = entry.split()
        deg = parts[0] if parts and parts[0].lower() in DEGREES else ""
        rem = entry[len(deg):].strip() if deg else entry

        if "," in rem:
            name, job = map(str.strip, rem.split(",", 1))
        else:
            name, job = rem, ""

        participants.append({
            "degree": deg,
            "participant": name,
            "pjob": job,
            "attendance": current_attendance
        })

    # 3) Lecturers: scan every paragraph containing “vadīs”
    lecturers = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if "vadīs" not in txt:
            continue
        # Everything after the first “vadīs”
        tail = re.split(r"vadīs[:\-]?", txt, 1, flags=re.IGNORECASE)[-1]
        # Split multiple lecturers on “;” or “ un ”
        for seg in re.split(r";|\s+un\s+", tail):
            seg = seg.strip().rstrip(".;")
            if not seg:
                continue
            if "," in seg:
                name, job = map(str.strip, seg.split(",", 1))
            else:
                # Fallback: first two‐word match
                m = NAME_RE.search(seg)
                if m:
                    name = f"{m.group(1)} {m.group(2)}"
                    job = seg.replace(name, "").lstrip(", ").strip()
                else:
                    name, job = seg, ""
            lecturers.append({"lecturer": name, "ljob": job})

    # Dedupe lecturers
    seen = set()
    uniq = []
    for lec in lecturers:
        key = (lec["lecturer"], lec["ljob"])
        if key not in seen:
            seen.add(key)
            uniq.append(lec)
    lecturers = uniq

    # 4) Combine into rows with the new Attendance column
    rows = []
    for idx, (p, l) in enumerate(zip_longest(participants, lecturers, fillvalue={})):
        rows.append({
            "Date": full_dt if idx == 0 else "",
            "Degree": p.get("degree", ""),
            "Participant": p.get("participant", ""),
            "Participant Job": p.get("pjob", ""),
            "Attendance": p.get("attendance", ""),
            "Lecturer": l.get("lecturer", ""),
            "Lecturer Job": l.get("ljob", ""),
        })
    df = pd.DataFrame(rows)

    # 5) Export to Excel
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
        ws = writer.sheets["Data"]
        for ci, col in enumerate(df.columns, 1):
            maxlen = df[col].astype(str).map(len).max()
            ws.column_dimensions[get_column_letter(ci)].width = max(maxlen, len(col)) + 2
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
