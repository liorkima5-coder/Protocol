from flask import Flask, render_template_string, request, send_file
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = "uploads"
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# ---------- HTML של הטופס ----------
HTML_FORM = """
<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<title>דו"ח ישיבה</title>
<style>
body { font-family: Arial; background-color: #f8f8f8; direction: rtl; text-align: right; }
form { background: white; padding: 20px; border-radius: 10px; max-width: 600px; margin: 30px auto; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
input, textarea, button, select { width: 100%; margin-bottom: 10px; padding: 10px; font-size: 16px; border: 1px solid #ccc; border-radius: 5px; }
button { background-color: #007bff; color: white; border: none; cursor: pointer; }
button:hover { background-color: #0056b3; }
.add-row { background-color: #28a745; }
.add-row:hover { background-color: #1e7e34; }
hr { border: 0; border-top: 1px solid #ccc; margin: 10px 0; }
</style>
<script>
function addRow() {
  const container = document.getElementById('rows');
  const index = container.children.length + 1;
  const div = document.createElement('div');
  div.innerHTML = `
    <label>מס"ד:</label><input name="id_${index}" value="${index}">
    <label>נושא:</label><input name="topic_${index}">
    <label>מהות:</label><input name="essence_${index}">
    <label>הערות:</label><input name="remarks_${index}">
    <hr>`;
  container.appendChild(div);
}
</script>
</head>
<body>
  <form method="post" enctype="multipart/form-data">
    <h2>טופס יצירת דו"ח ישיבה</h2>

    <label>שם הפרויקט:</label>
    <input name="project_name" required>

    <label>נושא הפגישה:</label>
    <input name="meeting_subject" required>

    <label>תאריך:</label>
    <input type="date" name="date" required>

    <label>משתתפים:</label>
    <textarea name="participants"></textarea>
    
    <label>העתקים:</label>
    <textarea name="copies"></textarea>

    <label>אופן הפגישה:</label>
    <input name="meeting_type">

    <label>רשם:</label>
    <input name="recorder">

    <h3>סיכום פגישה</h3>
    <div id="rows"></div>
    <button type="button" class="add-row" onclick="addRow()">➕ הוסף שורה</button>

    <h3>העלאת תמונות</h3>
    <input type="file" name="images" accept="image/*" multiple>

    <button type="submit">📄 צור דו"ח</button>
  </form>
</body>
</html>
"""

# ---------- צד שרת ----------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        
        # --- (תיקון 1: אתחול doc מוקדם) ---
        template_path = "פרומט פרוטוקול ישיבה.docx"
        doc = DocxTemplate(template_path)
        
        # --- קריאת נתוני הטופס ---
        context = {
            "project_name": request.form["project_name"],
            "meeting_subject": request.form["meeting_subject"],
            "date": request.form["date"],
            "participants": request.form["participants"],
            "copies": request.form["copies"],
            "meeting_type": request.form["meeting_type"],
            "recorder": request.form["recorder"],
        }

        # --- בניית טבלת הסיכום ---
        summary_table = []
        i = 1
        while f"id_{i}" in request.form:
            summary_table.append({
                "id": request.form[f"id_{i}"],
                "topic": request.form[f"topic_{i}"],
                "essence": request.form[f"essence_{i}"],
                "remarks": request.form[f"remarks_{i}"]
            })
            i += 1
        context["summary_table"] = summary_table

        # --- העלאת תמונות ---
        images = []
        if "images" in request.files:
            for img in request.files.getlist("images"):
                if img.filename:
                    filename = secure_filename(img.filename)
                    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    img.save(path)
                    # --- (תיקון 2: שימוש ב-doc) ---
                    images.append(InlineImage(doc, path, width=Mm(80)))
        context["images"] = images

        # --- יצירת שם קובץ בטוח ---
        safe_name = "".join(c for c in context["project_name"] if c.isalnum() or c in (' ', '-', '_')).strip()
        output_path = f"דוח ישיבה - {safe_name}.docx"

        # --- יצירת הדוח ---
        # doc הוגדר למעלה, כאן רק הרינדור והשמירה
        doc.render(context)
        doc.save(output_path)

        return send_file(output_path, as_attachment=True)

    return render_template_string(HTML_FORM)

