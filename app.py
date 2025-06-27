from flask import Flask, render_template, request, send_file
from docx import Document
import pandas as pd
import re
import io

app = Flask(__name__)

def extract_participant_data(doc):
    # 1) Extract date & time from paragraph 6
    date_para = doc.paragraphs[6].text.strip()
    date_match = re.search(r'202\d\. gada \d{1,2}\. [a-zāēūī]+', date_para)
    time_match = re.search(
        r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})',
        date_para
    )
    date = date_match.group() if date_match else "N/A"
    time = f"{time_match.group(1)}–{time_match.group(2)}" if time_match else "N/A"
    datetime_info = f"{date} {time}"

    # 2) Collect participants from paragraphs 10 onward until one ends with a dot
    participants_data = []
    for para in doc.paragraphs[10:]:
        text = para.text.strip()
        if not text:
            continue

        # 2a) Reconstruct bold names in this paragraph
        bold_names = []
        buffer = ""
        for run in para.runs:
            if run.bold:
                buffer += run.text
            elif buffer:
                bold_names.append(buffer.strip())
                buffer = ""
        if buffer:
            bold_names.append(buffer.strip())

        # 2b) Split into individual segments by semicolon
        segments = [seg.strip() for seg in text.split(";") if seg.strip()]
        for idx, seg in enumerate(segments):
            # Degree = first word
            degree = seg.split()[0] if seg.split() else "N/A"
            # Name = the idx-th bold name
            name = bold_names[idx] if idx < len(bold_names) else "N/A"
            # Job = text after “Name,”
            job = seg.split(f"{name},", 1)[-1].strip(" .") if f"{name}," in seg else "N/A"

            participants_data.append({
                "Date": datetime_info,
                "Degree": degree,
                "Name": name,
                "Job": job
            })

        # Stop at paragraph ending with a period
        if text.endswith("."):
            break

    return pd.DataFrame(participants_data)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')
    if not file or not file.filename.lower().endswith('.docx'):
        return 'Invalid file format. Please upload a .docx file.', 400

    # Load and parse
    doc = Document(file)
    df = extract_participant_data(doc)

    # Write to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Participants')
    output.seek(0)

    return send_file(
        output,
        download_name='participants.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(__import__('os').environ.get('PORT', 10000)), debug=True)
