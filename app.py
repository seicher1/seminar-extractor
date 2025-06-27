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
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    # === 📅 Date & Time ===
    date_regex = r'202\d\. gada \d{1,2}\. [a-zāēūī]+'
    time_regex = r'no plkst\.?\s*(\d{1,2}[:.]\d{2})\s*(?:līdz|–)\s*plkst\.?\s*(\d{1,2}[:.]\d{2})'

    date = next((re.search(date_regex, p).group() for p in paragraphs if re.search(date_regex, p)), "N/A")
    time_match = next((re.search(time_regex, p) for p in paragraphs if re.search(time_regex, p)), None)
    time = f"{time_match.group(1)}–{time_match.group(2)}" if time_match else "N/A"
    full_datetime = f"{date} {time}"

    # === 🧍‍♀️ Participants ===
    participants = []
    degree_keywords = r"(majore|virsleitnants|kapteinis|inspektors|seržants|leitnants|virsseržants)"

    # Find index range of participants block
    start_idx = end_idx = None
    for i, p in enumerate(paragraphs):
        if "Uz mācību semināriem" in p:
            start_idx = i + 1
        if "Mācību semināru vadīs" in p:
            end_idx = i
            break

    if start_idx is None or end_idx is None:
        return full_datetime, [], []

    # Combine all participant lines into one text blob
    participant_blob = " ".join(paragraphs[start_idx:end_idx])

    # Match each participant line
    pattern = rf"\b{degree_keywords}\b\s+([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][^\s,–]+(?:[- ][A-ZĀČĒĢĪĶĻŅŖŠŪŽ]?[^\s,–]+)*)[,–]\s*(.*?)(?:[;.](?=\s*\b{degree_keywords}\b)|$)"
    matches = re.finditer(pattern, participant_blob, re.IGNORECASE)

    for m in matches:
        degree = m.group(1).strip()
        name = m.group(2).strip()
        job = m.group(3).strip(" ,;.")
        participants.append({
            "degree": degree,
            "name": name,
            "job": job
        })

    # === 👩‍🏫 Lecturers ===
    lecturers = []
    for p in paragraphs:
        if "Mācību semināru vadīs" in p:
            line = p.split("Mācību semināru vadīs", 1)[-1].strip("–: ")
            entries = re.split(r'un|,', line)
            for entry in entries:
                text = entry.strip()
                name_match = re.search(r'([A-ZĀČĒĢĪĶĻŅŖŠŪŽ][a-zāčēģīķļņŗšūž]+\s+[A-ZĀČĒĢĪĶĻŅŖŠŪŽ][a-zāčēģīķļņŗšūž]+)', text)
                if name_match:
                    name = name_match.group(1)
                    job = text.replace(name, "").strip(", ")
                else:
                    name = "—"
                    job = text
                lecturers.append({
                    "name": name,
                    "job": job
                })
            break

    return full_datetime, participants, lecturers


def save_to_excel(date_time, participants, lecturers, output_path):
    rows = []
    max_len = max(len(participants), len(lecturers))
    for i in range(max_len):
        rows.append({
            "Datums un laiks": date_time if i == 0 else "",
            "Pakāpe": participants[i]["degree"] if i < len(participants) else "",
            "Dalībnieks": participants[i]["name"] if i < len(participants) else "",
            "Dalībnieka amats": participants[i]["job"] if i < len(participants) else "",
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
