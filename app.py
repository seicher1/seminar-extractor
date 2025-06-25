import os
import re
import pandas as pd
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from docx import Document
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

def extract_info(filepath):
    doc = Document(filepath)
    full_text = "\n".join([p.text for p in doc.paragraphs])

    # Extract date and time
    date_match = re.search(r'(\d{4}\. gada \d{1,2}\. maijs)', full_text)
    time_match = re.search(r'no plkst\. *(\d{1,2}:\d{2}) līdz plkst\. *(\d{1,2}:\d{2})', full_text)
    date = date_match.group(1) if date_match else "N/A"
    time_range = f"{time_match.group(1)}–{time_match.group(2)}" if time_match else "N/A"
    full_date = f"{date} {time_range}"

    # Extract participants and lecturers
    participants = []
    lecturers = []
    collecting_lecturers = False

    for para in doc.paragraphs:
        degree = ""
        name = ""
        job = ""
        bold_runs = [run for run in para.runs if run.bold]
        if bold_runs:
            # Combine all bold parts
            bold_text = " ".join([run.text.strip() for run in para.runs if run.bold]).strip()
            match = re.match(r'([A-ZĒĪĀŪČĻŅŠŽ][a-zēīāūčļņšž]+)\s+([A-ZĒĪĀŪČĻŅŠŽ][a-zēīāūčļņšž]+\s+[A-ZĒĪĀŪČĻŅŠŽ][a-zēīāūčļņšž]+)', bold_text)
            if match:
                # Lecturer
                name = f"{match.group(2).strip()},{match.group(1).strip()}"
                job = ""
                collecting_lecturers = True
            else:
                # Participant: degree + name
                parts = bold_text.split()
                if len(parts) >= 2:
                    degree = parts[0]
                    name = " ".join(parts[1:])
                    job = ""
                    collecting_lecturers = False
        else:
            # Add job description to last person
            job = para.text.strip()

        if name:
            if collecting_lecturers:
                lecturers.append({"name": name, "job": job})
            else:
                participants.append({"degree": degree, "name": name, "job": job})
        elif job:
            if collecting_lecturers and lecturers:
                lecturers[-1]["job"] += " " + job
            elif not collecting_lecturers and participants:
                participants[-1]["job"] += " " + job

    # Prepare dataframe
    max_len = max(len(participants), len(lecturers))
    rows = []

    for i in range(max_len):
        rows.append({
            "Date": full_date if i == 0 else "",
            "Degree": participants[i]["degree"] if i < len(participants) else "",
            "Participant": participants[i]["name"] if i < len(participants) else "",
            "Participant Job": participants[i]["job"] if i < len(participants) else "",
            "Lecturer": lecturers[i]["name"] if i < len(lecturers) else "",
            "Lecturer Job": lecturers[i]["job"] if i < len(lecturers) else ""
        })

    return pd.DataFrame(rows)

def save_to_excel(df, output_path):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Seminar Info")
        ws = writer.sheets["Seminar Info"]

        # Style headers
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="4F81BD")
        alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = alignment

        # Auto-adjust column widths
        for i, column in enumerate(df.columns, 1):
            max_length = max(df[column].astype(str).map(len).max(), len(column)) + 2
            ws.column_dimensions[get_column_letter(i)].width = max_length

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
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], filename.replace(".docx", ".xlsx"))
        save_to_excel(df, output_path)
        return send_file(output_path, as_attachment=True)

    return render_template('index.html')

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
