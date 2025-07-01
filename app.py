from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import io
import os
import re
from itertools import zip_longest
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# Degree keywords
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

# Name matcher: two capitalized words, hyphens allowed
NAME_RE = re.compile(r"\b([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\wāčēģīķļņšžūŗĀČĒĢĪĶĻŅŖŠŪŽŅŖŠŪŽ\-–]+)\s+([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\wāčēģīķļņšžūŗĀČĒĢĪĶĻŅŖŠŪŽŅŖŠŪŽ\-–]+))\b")

# Helper regex
NUM_ONLY = re.compile(r"^\d+\.\d+\.$")
DATE_RE = re.compile(r"202\d\. gada \d{1,2}\. [^\n,]+")
TIME_RE = re.compile(r"no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})")


def extract_data(doc):
    # 1) Date & Time
    full_text = "\n".join(p.text for p in doc.paragraphs)
    date_m = DATE_RE.search(full_text)
    time_m = TIME_RE.search(full_text)
    date_str = date_m.group().strip() if date_m else "N/A"
    time_str = f"{time_m.group(1)}–{time_m.group(2)}" if time_m else "N/A"
    full_dt = f"{date_str} {time_str}"

    # 2) Participants
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    # find participants block
    start = next((i for i,t in enumerate(paras) if "deleģētas" in t), None)
    # find next 'vadīs' after participants
    end = None
    if start is not None:
        for j in range(start+1, len(paras)):
            if "vadīs" in paras[j]:
                end = j
                break
    segment_lines = paras[start+1:end] if start is not None and end is not None else paras
    # add table cells
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if txt:
                    segment_lines.append(txt)
    # normalize entries
    entries = []
    i = 0
    while i < len(segment_lines):
        line = segment_lines[i]
        if NUM_ONLY.match(line):
            k = i+1
            while k < len(segment_lines) and not segment_lines[k].strip(): k += 1
            if k < len(segment_lines): entries.append(segment_lines[k].strip())
            i = k+1
            continue
        first = line.split()[0].lower()
        if first in DEGREES: entries.append(line)
        i += 1
    # parse participants
    participants = []
    for seg in entries:
        parts = seg.split(); deg = parts[0] if parts and parts[0].lower() in DEGREES else ""
        rem = seg[len(deg):].strip() if deg else seg
        nm = NAME_RE.search(rem)
        name = f"{nm.group(1)} {nm.group(2)}" if nm else ""
        job = ""
        if ',' in rem and name:
            _,_,after = rem.partition(f"{name},"); job = after.strip()
        participants.append({'degree':deg,'participant':name,'pjob':job})

    # 3) Lecturers: all 'vadīs' blocks
    lecturers = []
    # find all paragraphs with 'vadīs'
    for idx, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if 'vadīs' in t:
            # collect subsequent lines until blank or new date
            block = []
            j = idx+1
            while j < len(doc.paragraphs):
                nxt = doc.paragraphs[j].text.strip()
                if not nxt or re.match(r'^\d{4}\.', nxt): break
                block.append(nxt)
                j += 1
            # parse each line in block
            for line in block:
                entry = line.rstrip('.;').strip()
                if not entry: continue
                nm = NAME_RE.search(entry)
                if nm:
                    name = f"{nm.group(1)} {nm.group(2)}"
                    _,_,after = entry.partition(f"{name},"); job = after.strip()
                else:
                    if ',' in entry: name,job = map(str.strip, entry.split(',',1))
                    else: name,job = entry, ''
                lecturers.append({'lecturer':name,'ljob':job})
    # remove duplicates while preserving order
    seen = set(); uniq_lect=[]
    for lec in lecturers:
        key = (lec['lecturer'], lec['ljob'])
        if key not in seen:
            seen.add(key); uniq_lect.append(lec)
    lecturers = uniq_lect

    # 4) Assemble rows
    rows = []
    for i, (p, l) in enumerate(zip_longest(participants, lecturers, fillvalue={})):  
        rows.append({
            'Date': full_dt if i==0 else '',
            'Degree': p.get('degree',''),
            'Participant': p.get('participant',''),
            'Participant Job': p.get('pjob',''),
            'Lecturer': l.get('lecturer',''),
            'Lecturer Job': l.get('ljob','')
        })
    df = pd.DataFrame(rows)

    # 5) Excel export
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data'); ws=writer.sheets['Data']
        for ci,col in enumerate(df.columns,1):
            mx=df[col].astype(str).map(len).max(); ws.column_dimensions[get_column_letter(ci)].width=max(mx,len(col))+2
    out.seek(0); return out

@app.route('/')
def index(): return render_template('index.html')
@app.route('/upload', methods=['POST'])
def upload():
    f=request.files.get('file')
    if not f or not f.filename.lower().endswith('.docx'): return 'Upload .docx',400
    excel_io=extract_data(Document(f)); return send_file(excel_io, download_name='seminar_data.xlsx', as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__=='__main__': app.run(host='0.0.0.0', port=int(os.environ.get('PORT',10000)), debug=True)
