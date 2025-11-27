import os
import json
import requests
import base64
from flask import Flask, render_template_string, request, send_file
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = "uploads"
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# --- הקישור לסקריפט שלך ---
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyoziY3lj7DKLV2P9Orf-1tqxRJ4iT-TobTXMhNGJFJNECi2Axv-cOVeaQ8VfMolQIt/exec"

def update_google_sheet_with_file(project, subject, date_val, file_path):
    try:
        # קריאת הקובץ והמרה ל-Base64
        with open(file_path, "rb") as f:
            file_content = base64.b64encode(f.read()).decode("utf-8")

        payload = {
            "project_name": project,
            "meeting_subject": subject,
            "date": date_val,
            "filename": os.path.basename(file_path),
            "file_content": file_content
        }
        
        response = requests.post(APPS_SCRIPT_URL, json=payload)
        print(f"Sent to Google Sheets. Response: {response.text}")
        
    except Exception as e:
        print(f"Error updating Google Sheet: {e}")

# ---------- HTML של הטופס ----------
HTML_FORM = """
<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>דו"ח ישיבה - מערכת ניהול</title>
<style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #e3f2fd; color: #333; direction: rtl; text-align: right; margin: 0; padding-bottom: 50px; }
    .container { background: white; padding: 25px; border-radius: 15px; max-width: 600px; margin: 40px auto; box-shadow: 0 4px 20px rgba(0,0,0,0.1); border-top: 5px solid #1976d2; }
    .logo-container { text-align: center; margin-bottom: 20px; }
    .logo-container img { max-height: 80px; }
    h2, h3 { color: #1565c0; text-align: center; margin-top: 0; }
    label { font-weight: bold; font-size: 14px; margin-top: 10px; display: block; color: #555; }
    input, textarea, select { width: 100%; box-sizing: border-box; margin-bottom: 15px; padding: 12px; font-size: 16px; border: 1px solid #ddd; border-radius: 8px; background-color: #fdfdfd; transition: border-color 0.3s; }
    input:focus, textarea:focus { border-color: #1976d2; outline: none; }
    button { width: 100%; padding: 14px; font-size: 18px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; transition: background 0.3s; }
    .add-row-btn { background-color: #4caf50; color: white; margin-bottom: 20px; }
    .add-row-btn:hover { background-color: #388e3c; }
    .submit-btn { background-color: #1976d2; color: white; margin-top: 10px; }
    .submit-btn:hover { background-color: #0d47a1; }
    .dynamic-row { background: #f1f8e9; padding: 15px; border-radius: 8px; margin-bottom: 10px; border: 1px solid #c5e1a5; }
    .footer { text-align: center; font-size: 12px; color: #777; margin-top: 30px; padding: 10px; }
</style>
<script>
    function saveFormData() {
        const formData = {};
        const inputs = document.querySelectorAll('input, textarea');
        inputs.forEach(el => { if (el.name && el.type !== 'file' && el.type !== 'submit') formData[el.name] = el.value; });
        const rowsContainer = document.getElementById('rows');
        formData['dynamic_rows_count'] = rowsContainer.children.length;
        localStorage.setItem('meetingReportData', JSON.stringify(formData));
    }
    function restoreFormData() {
        const saved = localStorage.getItem('meetingReportData');
        if (!saved) return;
        const formData = JSON.parse(saved);
        if (formData['dynamic_rows_count']) { for (let i = 0; i < formData['dynamic_rows_count']; i++) { addRow(); } }
        for (const [key, value] of Object.entries(formData)) { const el = document.getElementsByName(key)[0]; if (el) el.value = value; }
    }
    function addRow() {
        const container = document.getElementById('rows');
        const index = container.children.length + 1;
        const div = document.createElement('div');
        div.className = 'dynamic-row';
        div.innerHTML = `<label>#${index} נושא:</label><input name="topic_${index}" oninput="saveFormData()"><label>מהות:</label><input name="essence_${index}" oninput="saveFormData()"><label>הערות:</label><input name="remarks_${index}" oninput="saveFormData()"><input type="hidden" name="id_${index}" value="${index}">`;
        container.appendChild(div);
        saveFormData();
    }
    window.onload = function() {
        restoreFormData();
        document.querySelector('form').addEventListener('input', function(e) { if (e.target.type !== 'file') saveFormData(); });
    };
</script>
</head>
<body>
  <div class="container">
      <div class="logo-container"><img src="/hs.jpg" alt="Logo"></div>
      <form method="post" enctype="multipart/form-data" onsubmit="localStorage.removeItem('meetingReportData');">
        <h2>טופס יצירת דו"ח ישיבה</h2>
        <label>שם הפרויקט:</label><input name="project_name" required>
        <label>נושא הפגישה:</label><input name="meeting_subject" required>
        <label>תאריך:</label><input type="date" name="date" required>
        <label>משתתפים:</label><textarea name="participants" rows="3"></textarea>
        <label>העתקים:</label><textarea name="copies" rows="2"></textarea>
        <label>אופן הפגישה:</label><input name="meeting_type">
        <label>רשם:</label><input name="recorder">
        <h3>סיכום פגישה</h3>
        <div id="rows"></div>
        <button type="button" class="add-row-btn" onclick="addRow()">➕ הוסף שורה</button>
        <h3>העלאת תמונות</h3>
        <input type="file" name="images" accept="image/*" multiple style="background:none; border:none;">
        <button type="submit" class="submit-btn">📄 צור דו"ח והורד</button>
      </form>
      <div class="footer">עוצב ופותח על ידי ליאור קימה</div>
  </div>
</body>
</html>
"""

# ---------- צד שרת ----------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        template_path = "פרומט פרוטוקול ישיבה.docx"
        if not os.path.exists(template_path):
            return "שגיאה: קובץ התבנית לא נמצא בשרת.", 500

        doc = DocxTemplate(template_path)
        
        # איסוף הנתונים
        project_name = request.form.get("project_name", "")
        meeting_subject = request.form.get("meeting_subject", "")
        date_str = request.form.get("date", "") # מגיע בפורמט YYYY-MM-DD

        context = {
            "project_name": project_name,
            "meeting_subject": meeting_subject,
            "date": date_str,
            "participants": request.form.get("participants", ""),
            "copies": request.form.get("copies", ""),
            "meeting_type": request.form.get("meeting_type", ""),
            "recorder": request.form.get("recorder", ""),
        }

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

        images = []
        if "images" in request.files:
            for img in request.files.getlist("images"):
                if img.filename:
                    filename = secure_filename(img.filename)
                    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    img.save(path)
                    images.append(InlineImage(doc, path, width=Mm(80)))
        context["images"] = images

        # --- שינוי שם הקובץ (החלק המעודכן) ---
        def clean_text(text):
            # שומר רק אותיות, מספרים, רווחים ומקפים
            return "".join(c for c in text if c.isalnum() or c in (' ', '-', '_')).strip()

        safe_project = clean_text(project_name)
        safe_subject = clean_text(meeting_subject)
        
        # הפורמט: פרויקט + נושא + תאריך
        output_path = f"{safe_project} - {safe_subject} - {date_str}.docx"

        doc.render(context)
        doc.save(output_path)

        # עדכון גוגל שיטס ושמירה בדרייב
        update_google_sheet_with_file(project_name, meeting_subject, date_str, output_path)

        return send_file(output_path, as_attachment=True)

    return render_template_string(HTML_FORM)

@app.route('/hs.jpg')
def serve_logo():
    return send_file('hs.jpg')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
