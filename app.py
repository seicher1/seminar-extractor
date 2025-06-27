import os
import re
import pandas as pd
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from docx import Document
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

def extract_data_from_docx(filepath):
    doc = Document(filepath)
    paragraphs = [p for p in doc.paragraphs if p.text.strip()]

    # === 📅 Date + Time ===
    date_pattern = r'202\d\. gada \d{1,2}\. [a-zāēūī]+'
    time_pattern = r'no plkst\. *(\d{1,2}:\d{2}) līdz plkst\. *(\d{1,2}:\d{2})'
    
    date = next((re.search(date_pattern, p.text, re.IGNORECASE).group()
                 for p in paragraphs if re.search(date_pattern, p.text, re.IGNORECASE)), "N/A")

    time_match = next((re.search(time_pattern, p.text)
                       for p in paragraphs if re.search(time_pattern, p.text)), None)
    
    time = f"{time_match.group(1)}–{time_match.group(2)}" if time_match else "N/A"
    full_datetime = f"{date} {time}"

    # === 🧍 Participants ===
    participants = []
    for p in paragraphs:
        if re.match(r'^1\.\d+\.', p.text.strip()):
            degree_match = re.match(r'^1\.\d+\.\s+(\w+)', p.text.strip())
            degree = degree_match.group(1) if degree_match else ""

            bold_name = next((run.text.strip() for run in p.runs if run.bold and run.text.strip()), "")
            job_match = re.search(r',\s*(.+)', p.text)

            job = job_match.group(1).strip() if job_match else ""

            participants.append({
                "degree": degree,
                "name": bold_name,
                "job": job
            })

    # === 👩‍🏫 Lecturers ===
    lecturers = []
    collecting = False
    for p in paragraphs:
        if "Mācību semināru vadīs" in p.text:
            collecting = True
            continue

        if collecting:
            bold_name = next((run.text.strip() for run in p.runs if run.bold and run.text.strip()), "")
            if bold_name:
                # New lecturer starts here
                normal_text = " ".join(run.text.strip() for run in p.runs if not run.bold).strip()
                lecturers.append({
                    "name": bold_name,
                    "job": normal_text
                })
            elif lecturers:
                # Continue adding to last job description
                lecturers[-1]["job"] += " " + p.text.strip()

    return full_datetime, participants, lecturers

def save_to_excel(date_time, participants, lecturers, output_path):
    rows = []
    max_len = max(len(participants), len(lecturers))
    for i in range(max_len):
        rows.append({
            "Date": date_time if i == 0 else "",
            "Degree": participants[i]["degree"] if i < len(participants) else "",
            "Participant": participants[i]["name"] if i < len(participants) else "",
            "Participant Job": participants[i]["job"] if i < len(participants) else "",
            "Lecturer": lecturers[i]["name"] if i < len(lecturers) else "",
            "Lecturer Job": lecturers[i]["job"] if i < len(lecturers) else ""
        })

    df = pd.DataFrame(rows)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Seminar Info')
        ws = writer.sheets['Seminar Info']

        header_font = Font(bold=True, color="FFFFFF")
        fill = PatternFill("solid", fgColor="4F81BD")
        align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = fill
            cell.alignment = align

        for col in ws.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col) + 2
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max_length

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        uploaded_file = request.files['file']
        if not uploaded_file or not uploaded_file.filename.endswith(".docx"):
            return render_template("index.html", error="Please upload a valid .docx file.")

        filename = secure_filename(uploaded_file.filename)
        docx_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        uploaded_file.save(docx_path)

        date_time, participants, lecturers = extract_data_from_docx(docx_path)
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], filename.replace(".docx", ".xlsx"))
        save_to_excel(date_time, participants, lecturers, output_path)

        return send_file(output_path, as_attachment=True)

    return render_template("index.html")
    
if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
