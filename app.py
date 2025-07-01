from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io
import os
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

# Simple name matcher for lecturers (two words)
NAME_RE = re.compile(r"\b([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w\-]+)\s+([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][\w\-]+)\b")

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

    # 2) Participants block
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    start = next((i for i,t in enumerate(paras) if "deleģētas" in t), None)
    end = None
    if start is not None:
        for j in range(start+1, len(paras)):
            if "vadīs" in paras[j]: end = j; break
    segment_lines = paras[start+1:end] if start is not None and end else paras
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if txt: segment_lines.append(txt)

    # 3) Normalize entries
    entries = []
    i = 0
    while i < len(segment_lines):
        ln = segment_lines[i]
        if NUM_ONLY.match(ln):
            k = i+1
            while k < len(segment_lines) and not segment_lines[k].strip(): k+=1
            if k < len(segment_lines): entries.append(segment_lines[k].strip())
            i = k+1; continue
        first = ln.split()[0].lower()
        if first in DEGREES: entries.append(ln)
        i += 1

    # 4) Parse participants with new name logic
    participants = []
    for seg in entries:
        parts = seg.split()
        deg = parts[0] if parts and parts[0].lower() in DEGREES else ""
        # substring between degree and comma is full name
        name = ""
        job = ""
        if deg:
            rem = seg[len(deg):].strip()
            if ',' in rem:
                name_part, job_part = rem.split(',', 1)
                name = name_part.strip()
                job = job_part.strip()
            else:
                name = rem
        else:
            # fallback: before comma
            if ',' in seg:
                name, job = map(str.strip, seg.split(',',1))
            else:
                name = seg
        participants.append({'degree':deg,'participant':name,'pjob':job})

    # 5) Extract lecturers after participants
    last_par = 0
    for idx,p in enumerate(doc.paragraphs):
        if any(p.text.strip().startswith(f"{n}.") for n in range(1,10)): last_par=idx
    lecturers=[]
    for p in doc.paragraphs[last_par+1:]:
        t=p.text.strip()
        if not t or 'vadīs' not in t: continue
        tail = re.split(r'vadīs[:\-]?', t,1,flags=re.IGNORECASE)[-1]
        for part in re.split(r';|\s+un\s+', tail):
            ent = part.strip().rstrip('.;')
            if not ent: continue
            nm=''; jb=''
            if ',' in ent:
                nm, jb = map(str.strip, ent.split(',',1))
            else:
                nm = ent
            lecturers.append({'lecturer':nm,'ljob':jb})
    # dedupe
    seen=set(); uniq=[]
    for lec in lecturers:
        key=(lec['lecturer'],lec['ljob'])
        if key not in seen: seen.add(key); uniq.append(lec)
    lecturers=uniq

    # 6) Build rows
    rows=[]
    for i,(p,l) in enumerate(zip_longest(participants,lecturers,fillvalue={})):  
        rows.append({
            'Date': full_dt if i==0 else '',
            'Degree':p.get('degree',''),
            'Participant':p.get('participant',''),
            'Participant Job':p.get('pjob',''),
            'Lecturer':l.get('lecturer',''),
            'Lecturer Job':l.get('ljob','')
        })
    df=pd.DataFrame(rows)

    # 7) Excel export
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine='openpyxl') as w:
        df.to_excel(w,index=False,sheet_name='Data');ws=w.sheets['Data']
        for ci,col in enumerate(df.columns,1):
            mx=df[col].astype(str).map(len).max();ws.column_dimensions[get_column_letter(ci)].width=max(mx,len(col))+2
    out.seek(0); return out

@app.route('/')

def index(): return render_template('index.html')
@app.route('/upload',methods=['POST'])
def upload():
    f=request.files.get('file')
    if not f or not f.filename.lower().endswith('.docx'): return 'Upload .docx',400
    excel=extract_data(Document(f)); return send_file(excel,download_name='seminar_data.xlsx',as_attachment=True,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT',10000)),debug=True)
