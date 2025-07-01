from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
from itertools import zip_longest
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# Degree keywords (m/f)
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
# Simple name matcher
NAME_RE = re.compile(r"\b([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\s+([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w–]+)\b")

# Patterns
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

    # 2) Collect lines
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    start = next((i for i,t in enumerate(paras) if "deleģētas" in t), None)
    end1 = next((i for i,t in enumerate(paras) if "Mācību semināru vadīs" in t), None)
    end2 = next((i for i,t in enumerate(paras) if "Mācības vadīs" in t), None)
    end = end1 if end1 is not None else end2
    lines = []
    if start is not None and end is not None and end>start:
        lines = paras[start+1:end]
    else:
        lines = paras[:]
    # add table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if txt:
                    lines.append(txt)

    # 3) Normalize entries
    entries = []
    i=0
    while i < len(lines):
        line = lines[i]
        if NUM_ONLY.match(line):
            # pair with next
            j=i+1
            while j<len(lines) and not lines[j]: j+=1
            if j<len(lines): entries.append(lines[j])
            i=j+1
            continue
        # check first word degree
        first = line.split()[0].lower()
        if first in DEGREES:
            entries.append(line)
        i+=1

    # 4) Parse participants
    participants=[]
    for seg in entries:
        parts=seg.split()
        deg = parts[0] if parts and parts[0].lower() in DEGREES else ""
        nm_match = NAME_RE.search(seg)
        name = f"{nm_match.group(1)} {nm_match.group(2)}" if nm_match else ""
        job = seg.split(',',1)[1].strip() if ',' in seg and name else seg
        participants.append({'degree':deg,'participant':name,'pjob':job})

    # 5) Extract lecturers
    lecturers=[]
    for p in doc.paragraphs:
        t=p.text.strip()
        if "Mācību semināru vadīs" in t or "Mācības vadīs" in t:
            tail = re.split(r'Mācību semināru vadīs|Mācības vadīs', t,1)[-1]
            parts = re.split(r'\s+un\s+|,\s*(?=[A-ZĀČĒĢĪĶĻŅŖŠŪŽ])', tail)
            for part in parts:
                txt=part.strip().rstrip('.;')
                m=NAME_RE.match(txt)
                if m:
                    nm=f"{m.group(1)} {m.group(2)}"
                    jb=txt.replace(nm,'').lstrip(', ').strip()
                else:
                    nm,jb="—",txt
                lecturers.append({'lecturer':nm,'ljob':jb})
            break

    # 6) Build rows with zip_longest
    rows=[]
    for idx, (p, l) in enumerate(zip_longest(participants,lecturers,fillvalue={})):
        rows.append({
            'Date': full_dt if idx==0 else '',
            'Degree': p.get('degree',''),
            'Participant': p.get('participant',''),
            'Participant Job': p.get('pjob',''),
            'Lecturer': l.get('lecturer',''),
            'Lecturer Job': l.get('ljob','')
        })

    df=pd.DataFrame(rows)
    # Write excel
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine='openpyxl') as w:
        df.to_excel(w,index=False,sheet_name='Data')
        ws=w.sheets['Data']
        for idx,col in enumerate(df.columns,1):
            width=max(df[col].astype(str).map(len).max(),len(col))+2
            ws.column_dimensions[get_column_letter(idx)].width=width
    out.seek(0)
    return out

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload',methods=['POST'])
def upload():
    f=request.files.get('file')
    if not f or not f.filename.lower().endswith('.docx'):
        return 'Please upload .docx file',400
    doc=Document(f)
    excel_io=extract_data(doc)
    return send_file(
        excel_io,
        download_name='seminar_data.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__=='__main__':
    app.run(host='0.0.0.0',port=int(os.environ.get('PORT',10000)),debug=True)
