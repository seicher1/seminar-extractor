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

    # === 📅 Extract Date & Time ===
    date_regex = r'202\d\. gada \d{1,2}\. [a-zāēūī]+'
    time_regex = r'no plkst\. *(\d{1,2}:\d{2}) līdz plkst\. *(\d{1,2}:\d{2})'
    
    date = next((re.search(date_regex, p.text).group()
                 for p in paragraphs if re.search(date_regex, p.text)), "N/A")

    time_match = next((re.search(time_regex, p.text)
                       for p in paragraphs if re.search(time_regex, p.text)), None)
    time = f"{time_match.group(1)}–{time_match.group(2)}" if time_match else "N/A"

    full_datetime = f"{date} {time}"

    # === 🧍 Participants ===
    participants = []
    for p in paragraphs:
        text = p.text.strip()
        if text.lower().startswith(("majore", "inspektors", "kapteinis")):
            degree_match = re.match(r'^(\w+)', text)
            bold_name = next((run.text.strip() for run in p.runs if run.bold and run.text.strip()), "")
            job = text.split(",", 1)[-1].strip() if "," in text else ""

            participants.append({
                "degree": degree_match.group(1) if degree_match else "",
                "name": bold_name,
                "job": job
            })

    # === 👩‍🏫 Lecturers ===
    lecturers = []
    for p in paragraphs:
        if "Mācību semināru vadīs" in p.text:
            # Grab the entire paragraph
            lecturer_line = p.text.split("Mācību semināru vadīs -", 1)[-1]
            lecturer_segments = [seg.strip() for seg in lecturer_line.split("un")]

            for segment in lecturer_segments:
                name_match = re.search(r'([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][a-zāčēģīķļņŗšūž]+\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][a-zāčēģīķļņŗšūž]+)', segment)
                name = name_match.group(1) if name_match else "N/A"
                job = segment.replace(name, "").replace(",", "").strip() if name != "N/A" else segment.strip()

                lecturers.append({
                    "name": name,
                    "job": job
                })
            break  # Only one line of lecturers expected

    return full_datetime, participants, lecturers

def save_to_excel(date_time, participants, lecturers, output_path):
    rows = []
    max_len = max(len(participants), len(lecturers))
    for i in range(max_len):
        rows.append({
            "Datums un laiks": date_time if i == 0 else "",
            "Pakāpe": participants[i]["degree"] if i < len(participants) else "",
            "Dalībnieks": participants[i]["name"] if i < len(participants) else "",
            "Amats": participants[i]["job"] if i < len(participants) else "",
            "Semināra vadītājs": lecturers[i]["name"] if i < len(lecturers) else "",
            "Amats": lecturers[i]["job"] if i < len(lecturers) else ""
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
