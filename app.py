import os
import re
import pandas as pd
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from docx import Document

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

def extract_info(filepath):
    doc = Document(filepath)
    text = "\n".join([para.text for para in doc.paragraphs])

    # Date + Time
    date_match = re.search(r'(\d{4}\. gada \d{1,2}\. maijs)', text)
    time_match = re.search(r'no plkst\. *(\d{1,2}:\d{2}) līdz plkst\. *(\d{1,2}:\d{2})', text)
    date = date_match.group(1) if date_match else "N/A"
    time_range = f"{time_match.group(1)}–{time_match.group(2)}" if time_match else "N/A"
    full_date = f"{date} {time_range}"

    # Participants
    participants_block = re.search(r'deleģētas šādas Valsts policijas amatpersonas:(.*?)Mācību semināru vadīs',
                                   text, re.DOTALL)
    participants, participant_jobs = [], []
    if participants_block:
        lines = participants_block.group(1).split('\n')
        for line in lines:
            match = re.search(r'majore\s+([A-ZĒĪĀŪČĻŅŠŽ][a-zēīāūčļņšž]+\s+[A-ZĒĪĀŪČĻŅŠŽ][a-zēīāūčļņšž]+),\s+(.*)', line.strip())
            if match:
                participants.append(match.group(1).strip())
                participant_jobs.append(match.group(2).strip())

    # Lecturers
    lecturers_block = re.search(r'Mācību semināru vadīs\s*-\s*(.*?)(?=Nepieciešamā informācija|$)', text, re.DOTALL)
    lecturers, lecturer_jobs = [], []
    if lecturers_block:
        parts = lecturers_block.group(1).split("un")
        for part in parts:
            part = part.strip().strip(".")
            match = re.search(r'([A-ZĒĪĀŪČĻŅŠŽ][a-zēīāūčļņšž]+)\s+([A-ZĒĪĀŪČĻŅŠŽ][a-zēīāūčļņšž]+),\s*(.*)', part)
            if match:
                name = f"{match.group(2)},{match.group(1)}"
                job = match.group(3).strip()
                lecturers.append(name)
                lecturer_jobs.append(job)

    max_len = max(len(participants), len(lecturers))
    data = []
    for i in range(max_len):
        data.append({
            "Date": full_date if i == 0 else "",
            "Participant": participants[i] if i < len(participants) else "",
            "Participant Job": participant_jobs[i] if i < len(participant_jobs) else "",
            "Lecturer": lecturers[i] if i < len(lecturers) else "",
            "Lecturer Job": lecturer_jobs[i] if i < len(lecturer_jobs) else ""
        })

    return pd.DataFrame(data)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        f = request.files['file']
        if not f:
            return render_template('index.html', error="No file selected.")

        filename = secure_filename(f.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        f.save(filepath)

        df = extract_info(filepath)
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], filename.replace(".docx", ".csv"))
        df.to_csv(output_path, index=False, encoding='utf-8-sig')

        return send_file(output_path, as_attachment=True)

    return render_template('index.html')
