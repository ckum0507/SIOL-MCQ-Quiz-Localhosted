from flask import Flask, render_template, request, redirect, session, send_file, abort
from flask_session import Session
from openpyxl import load_workbook, Workbook
from docx import Document
from datetime import datetime
import os, zipfile, shutil, time
import argparse
import qrcode
import socket

quiz_hosted = False
quiz_theme = ""

app = Flask(__name__)
app.config.from_pyfile("config.py")
app.config["SESSION_TYPE"] = "filesystem"
Session(app)

test_context = {
    "name": "",
    "admin": "",
    "timestamp": "",
    "port": ""
}

UPLOAD_FOLDER = "uploads"
EXPORT_FOLDER = "exports"
TEMP_FOLDER = "temp"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)
os.makedirs(TEMP_FOLDER, exist_ok=True)

quiz_time = 0
questions = []

def get_test_name():
    return test_context.get("name", "")

def get_active_rules_file():
    files = [
        f for f in os.listdir(UPLOAD_FOLDER)
        if f.endswith("_rules.docx") and not f.startswith("~$")
    ]
    return os.path.join(UPLOAD_FOLDER, sorted(files)[-1]) if files else None

def get_active_questions_file():
    files = [
        f for f in os.listdir(UPLOAD_FOLDER)
        if f.endswith("_questions.xlsx") and not f.startswith("~$")
    ]
    return os.path.join(UPLOAD_FOLDER, sorted(files)[-1]) if files else None

def load_questions_from_active_file():
    global questions
    q_file = get_active_questions_file()
    if not q_file:
        return False
    wb = load_workbook(q_file)
    ws = wb.active
    questions.clear()
    letters = ["A", "B", "C", "D", "E"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(cell == "~" for cell in row):
            break
        option_map = {
            letters[i]: str(opt)
            for i, opt in enumerate(row[1:6])
            if opt not in (None, "|")
        }
        questions.append({"text": row[0], "options": option_map, "correct": str(row[6]).strip().upper()})
    wb.close()
    return True

def rename_uploaded_files():
    if not all(test_context.values()):
        return
    prefix = f"{test_context['name']}_{test_context['admin']}_{test_context['timestamp']}_{test_context['port']}"
    for base, suffix in [("rules", "docx"), ("questions", "xlsx")]:
        old = os.path.join(UPLOAD_FOLDER, f"{base}.{suffix}")
        new = os.path.join(UPLOAD_FOLDER, f"{prefix}_{base}.{suffix}")
        if os.path.exists(old):
            os.replace(old, new)

def get_lan_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip

import base64
from io import BytesIO


@app.route("/admin", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        if (
            request.form["username"] == app.config["ADMIN_USERNAME"]
            and request.form["password"] == app.config["ADMIN_PASSWORD"]
        ):
            session["admin"] = True
            return redirect("/admin/dashboard")
    return render_template("admin_login.html")

@app.route("/admin/dashboard")
def admin_dashboard():
    if not session.get("admin"):
        abort(403)

    lan_ip = get_lan_ip()
    port = test_context.get("port", "5000")

    return render_template(
        "admin_dashboard.html",
        lan_ip=lan_ip,
        port=port
    )




@app.route("/admin/upload_rules", methods=["POST"])
def upload_rules():
    if not session.get("admin"):
        abort(403)
    request.files["rules"].save(os.path.join(UPLOAD_FOLDER, "rules.docx"))
    return redirect("/admin/dashboard")

@app.route("/admin/upload_questions", methods=["POST"])
def upload_questions():
    global quiz_time

    if not session.get("admin"):
        abort(403)

    path = os.path.join(UPLOAD_FOLDER, "questions.xlsx")
    request.files["questions"].save(path)

    wb = load_workbook(path)
    ws = wb.active

    questions.clear()
    letters = ["A", "B", "C", "D", "E"]

    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(cell == "~" for cell in row):
            break
        options = {
            letters[i]: str(opt)
            for i, opt in enumerate(row[1:6])
            if opt not in (None, "|")
        }
        questions.append({
            "text": row[0],
            "options": options,
            "correct": str(row[6]).strip().upper()
        })

    wb.close()
    quiz_time = len(questions)
    return redirect("/admin/dashboard")

@app.route("/admin/preview", methods=["GET", "POST"])
def preview():
    global quiz_hosted, quiz_theme, quiz_time

    if not session.get("admin"):
        abort(403)

    if request.method == "POST":
        test_context["name"] = request.form["test_name"].replace(" ", "_")
        test_context["admin"] = app.config["ADMIN_USERNAME"]
        test_context["timestamp"] = datetime.now().strftime("%Y%m%d_%H%M%S")
        test_context["port"] = request.environ.get("SERVER_PORT", "5000")

        quiz_time = min(int(request.form["time"]), 180)
        quiz_theme = request.form.get("theme", "")
        quiz_hosted = True

        rename_uploaded_files()
        return redirect("/admin/dashboard")

    rules_lines = []
    preview_rules = os.path.join(UPLOAD_FOLDER, "rules.docx")
    if os.path.exists(preview_rules):
        doc = Document(preview_rules)
        rules_lines = [p.text for p in doc.paragraphs]
    return render_template("admin_preview.html", questions=questions, quiz_time=quiz_time, rules_lines=rules_lines, theme=quiz_theme )

@app.route("/admin/unhost")
def unhost():
    global quiz_hosted
    quiz_hosted = False
    session.pop("qr_base64", None)
    session.pop("qr_url", None)
    return redirect("/admin/dashboard")

@app.route("/")
def user_entry():
    if not quiz_hosted:
        abort(403)
    return render_template("user_1_details.html", test_name=get_test_name(), theme=quiz_theme)

@app.route("/rules", methods=["GET", "POST"])
def rules():
    if request.method == "POST":
        session["user"] = dict(request.form)
    doc = Document(get_active_rules_file())
    rules_lines = [p.text for p in doc.paragraphs]
    return render_template("user_2_rules.html", rules_lines=rules_lines, theme=quiz_theme, test_name=get_test_name())

@app.route("/quiz")
def quiz():
    if not session.get("user"):
        return redirect("/")
    if not questions:
        load_questions_from_active_file()
    return render_template("user_3_quiz.html", questions=questions, time=quiz_time * 60, theme=quiz_theme, user=session["user"], test_name=get_test_name())

@app.route("/submit", methods=["POST"]) 
def submit(): 
    if not quiz_hosted: 
        abort(403) 
    user = session.get("user") 
    if not user: 
        abort(400) 
    score = 0 
    user_answers = [] 
    for i, q in enumerate(questions): 
        chosen = request.form.get(f"q{i}") 
        user_answers.append(chosen) 
        if chosen == q["correct"]: 
            score += 1 
        prefix = ( f"{test_context['name']}_" f"{test_context['admin']}_" f"{test_context['timestamp']}_" f"{test_context['port']}" ) 
        wb_path = os.path.join(EXPORT_FOLDER, f"{prefix}_results.xlsx") 
        # ---- CREATE FILE & HEADERS IF NEEDED ---- 
        if not os.path.exists(wb_path): 
            wb = Workbook() 
            ws = wb.active 
            headers = [ "Timestamp", "Team Leader Name", "Team Name", "Standard", "Phone Number", "School Name", "School Address", "Total Score/Total", "Time Taken/Time Given" ] 
            for idx, q in enumerate(questions, start=1): 
                opt_text = " | ".join( [f"{k}) {v}" for k, v in q["options"].items()] ) 
                header = f"Q{idx}: {q['text']} | {opt_text} | Correct: {q['correct']}" 
                headers.append(header) 
            ws.append(headers) 
            wb.save(wb_path) 
            wb.close() 
            # ---- APPEND RESULT (RETRY SAFE) ---- 
            for _ in range(3): 
                try: 
                    wb = load_workbook(wb_path) 
                    ws = wb.active 
                    # ---- FORMAT TIME TAKEN ---- 
                    time_taken_sec = int(request.form.get("time_taken", 0)) 
                    taken_min = time_taken_sec // 60 
                    taken_sec = time_taken_sec % 60 
                    total_min = quiz_time 
                    total_sec = 0 
                    time_taken_str = ( f"{taken_min:02d}:{taken_sec:02d}/" f"{total_min:02d}:{total_sec:02d}" ) 
                    row = [ datetime.now().strftime("%Y-%m-%d %H:%M:%S"), user["leader"], user["team"], user["Standard"], user["phone"], user["school"], user["address"], score, time_taken_str ] 
                    # row.append(time_taken_str) 
                    row.extend(user_answers) 
                    ws.append(row) 
                    wb.save(wb_path) 
                    wb.close() 
                    break 
                except PermissionError: 
                    time.sleep(1) 
        else: 
            abort(500, "Results file is locked") 
        session.pop("user", None) 
        return redirect("/thanks")
@app.route("/thanks")
def thanks():
    return render_template("user_4_thanks.html", theme=quiz_theme, test_name=get_test_name())

@app.route("/exports/<path:filename>")
def serve_exports(filename):
    return send_file(os.path.join(EXPORT_FOLDER, filename))

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=5000)
    args = parser.parse_args()
    app.run(host=args.host, port=args.port, debug=True)
