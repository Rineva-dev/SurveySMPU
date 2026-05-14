from flask import Flask, abort, render_template, request, redirect, url_for, session, send_file
from datetime import datetime
from markupsafe import Markup
import json, os
from openpyxl import Workbook
import io, uuid
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv
load_dotenv()

def format_date(date_str):
    return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d-%m-%Y")

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv("SECRET_KEY")

app.jinja_env.globals.update(format_date=format_date)

from flask_sqlalchemy import SQLAlchemy

app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv(
    "DATABASE_URL",
    "sqlite:///survey.db"
)

app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

class Survey(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    start_date = db.Column(db.String(20))
    end_date = db.Column(db.String(20))
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.String(30))
    questions = db.Column(db.Text) 
    unsur_responden = db.Column(db.Text)
    responses = db.relationship(
        'SurveyResponse',
        backref='survey',
        lazy=True
    )
    @property
    def total_responden(self):
        return len(self.responses)
    
    # =========================
    # 🔥 STATUS OTOMATIS SURVEY
    # =========================
    @property
    def status(self):
        if not self.start_date or not self.end_date:
            return "tidak_valid"

        now = datetime.now().date()

        start = datetime.strptime(
            self.start_date, "%Y-%m-%d"
        ).date()

        end = datetime.strptime(
            self.end_date, "%Y-%m-%d"
        ).date()

        if now < start:
            return "belum_mulai"
        elif start <= now <= end:
            return "aktif"
        else:
            return "tutup"

class SurveyResponse(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    survey_id = db.Column(
        db.Integer,
        db.ForeignKey('survey.id'),
        nullable=False
    )

    unsur_responden = db.Column(
        db.String(100)
    )

    answers = db.Column(
        db.Text
    )

    created_at = db.Column(
        db.String(30)
    )

# -----------------------------
# LOGIN ADMIN
# -----------------------------
@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    error = None

    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        if username == "admin" and password == "admin123":
            session["admin"] = True
            return redirect(url_for("admin_dashboard"))
        else:
            error = "Username atau password salah"

    return render_template("admin/login.html", error=error)

@app.route("/")
def home():
    return redirect(url_for("admin_login"))

@app.route("/dashboard")
def admin_dashboard():

    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    surveys = Survey.query.all()

    total_survey = Survey.query.count()

    total_responden = SurveyResponse.query.count()

    survey_aktif = Survey.query.filter_by(
        is_active=True
    ).count()

    return render_template(
        "admin/dashboard.html",

        surveys=surveys,

        total_survey=total_survey,
        total_responden=total_responden,
        survey_aktif=survey_aktif,

        active_page="dashboard"
    )

@app.route("/survey")
def admin_survey():
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    surveys = Survey.query.all()

    return render_template(
        "admin/survey.html",
        surveys=surveys,
        active_page="survey"
    )

@app.route("/admin/survey/tambah", methods=["GET", "POST"])
def admin_survey_tambah():
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    if request.method == "POST":
        title = request.form.get("title")
        description = request.form.get("pembuka")
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")
        unsur_responden = request.form.getlist("unsur_responden[]")

        # =========================
        # SESSION TITLE
        # =========================

        session_titles = request.form.getlist(
            "session_title[]"
        )

        question_sessions = request.form.getlist(
            "question_session[]"
        )

        # container session
        sessions = []

        for s in session_titles:

            sessions.append({
                "title": s,
                "questions": []
            })

        # =========================
        # QUESTIONS
        # =========================

        texts = request.form.getlist("pertanyaan[]")
        types = request.form.getlist("type[]")
        counts = request.form.getlist("count[]")
        modes = request.form.getlist("mode[]")
        required_flags = request.form.getlist("is_required[]")

        for i in range(len(texts)):

            labels = request.form.getlist(
                f"option_label[{i}][]"
            )

            texts_opt = request.form.getlist(
                f"option_text[{i}][]"
            )

            options = []

            for j in range(len(texts_opt)):

                options.append({
                    "label": labels[j]
                        if j < len(labels)
                        else "",

                    "text": texts_opt[j]
                })

            q = {
                "text": texts[i],

                "type":
                    types[i]
                    if i < len(types)
                    else "isian",

                "count":
                    int(counts[i])
                    if i < len(counts)
                    and counts[i].isdigit()
                    else 1,

                "mode":
                    modes[i]
                    if i < len(modes)
                    else "single",

                "options": options,
                "is_required": True if i < len(required_flags) else False
            }

            # ambil session index
            session_index = int(
                question_sessions[i]
            )

            # masukkan ke session terkait
            sessions[session_index]["questions"].append(q)

        # 🔥 BARU BUAT OBJECT SURVEY
        new_survey = Survey(
            title=title,
            description=description,
            start_date=start_date,
            end_date=end_date,
            questions=json.dumps(sessions),
            unsur_responden=json.dumps(unsur_responden),
            is_active=True,
            created_at=datetime.now().strftime("%Y-%m-%d %H:%M")
        )

        db.session.add(new_survey)
        db.session.commit()

        return redirect(url_for("admin_survey"))

    return render_template("admin/survey_form.html", active_page="survey")

@app.route("/admin/survey/<int:survey_id>/toggle", methods=["POST"])
def toggle_survey_status(survey_id):
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    survey = Survey.query.get(survey_id)

    if survey:
        survey.is_active = not survey.is_active
        db.session.commit()

    return redirect(url_for("admin_survey"))

@app.route("/admin/survey/<int:survey_id>/edit", methods=["GET", "POST"])
def edit_survey(survey_id):

    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    survey = Survey.query.get_or_404(survey_id)
    
    if request.method == "POST":

        # =========================
        # BASIC INFO
        # =========================

        survey.title = request.form.get("title")
        survey.description = request.form.get("pembuka")
        survey.start_date = request.form.get("start_date")
        survey.end_date = request.form.get("end_date")
        unsur_responden = request.form.getlist("unsur_responden[]")
        survey.unsur_responden = json.dumps(unsur_responden)
        required_flags = request.form.getlist("is_required[]")

        # =========================
        # SESSION TITLE
        # =========================

        session_titles = request.form.getlist(
            "session_title[]"
        )

        question_sessions = request.form.getlist(
            "question_session[]"
        )

        sessions = []

        for s in session_titles:

            sessions.append({
                "title": s,
                "questions": []
            })

        # =========================
        # QUESTIONS
        # =========================

        texts = request.form.getlist("pertanyaan[]")

        types = request.form.getlist("type[]")

        counts = request.form.getlist("count[]")

        modes = request.form.getlist("mode[]")

        for i in range(len(texts)):

            labels = request.form.getlist(
                f"option_label[{i}][]"
            )

            texts_opt = request.form.getlist(
                f"option_text[{i}][]"
            )

            options = []

            for j in range(len(texts_opt)):

                options.append({
                    "label": labels[j]
                        if j < len(labels)
                        else "",

                    "text": texts_opt[j]
                })

            q = {
                "text": texts[i],

                "type":
                    types[i]
                    if i < len(types)
                    else "isian",

                "count":
                    int(counts[i])
                    if i < len(counts)
                    and counts[i].isdigit()
                    else 1,

                "mode":
                    modes[i]
                    if i < len(modes)
                    else "single",

                "options": options,
                "is_required": True if i < len(required_flags) else False
            }

            session_index = int(
                question_sessions[i]
            )

            sessions[session_index]["questions"].append(q)

        # =========================
        # SAVE QUESTIONS
        # =========================

        survey.questions = json.dumps(sessions)

        db.session.commit()

        return redirect(url_for("admin_survey"))

    return render_template(
        "admin/survey_form.html",
        survey=survey,
        active_page="survey",
        is_edit=True
    )

@app.route("/admin/survey/<int:survey_id>/hapus", methods=["POST"])
def hapus_survey(survey_id):
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    survey = Survey.query.get(survey_id)

    if survey:
        db.session.delete(survey)
        db.session.commit()

    return redirect(url_for("admin_survey"))

@app.route("/survey/<int:survey_id>", methods=["GET", "POST"])
def survey_responden(survey_id):

    survey = Survey.query.get_or_404(survey_id)

    # =========================
    # CEK STATUS SURVEY
    # =========================
    if survey.status != "aktif" or not survey.is_active:
        return render_template(
            "survey/survey_closed.html",
            survey=survey
        )

    # =========================
    # AMBIL TOKEN (GET / POST)
    # =========================
    token = (
        request.args.get("token")
        if request.method == "GET"
        else request.form.get("token")
    )

    session_token = session.get(
        f"survey_{survey_id}_token"
    )

    if not token or token != session_token:
        abort(404)
    
    survey = Survey.query.get_or_404(survey_id)

    # =========================
    # POST (SUBMIT JAWABAN)
    # =========================
    if request.method == "POST":

        token = request.form.get("token")
        session_token = session.get(f"survey_{survey_id}_token")

        if not token or token != session_token:
            abort(404)

        raw_sessions = json.loads(survey.questions or "[]")

        # format session
        if raw_sessions and "questions" not in raw_sessions[0]:
            sessions = [{
                "title": "Survey",
                "questions": raw_sessions
            }]
        else:
            sessions = raw_sessions

        # GLOBAL NUMBERING
        counter = 1
        for s in sessions:
            for q in s["questions"]:
                q["global_index"] = counter
                counter += 1

        unsur_responden = request.form.get("unsur_responden")
        errors = []
        answers = {}

        counter = 1
        error_session_index = 0

        for s_idx, session_item in enumerate(sessions):
            for q in session_item["questions"]:

                if q.get("is_required"):
                    key_prefix = f"q{counter}"

                    found = [
                        v for k, v in request.form.items()
                        if k.startswith(key_prefix)
                    ]

                    if not found or all(v.strip() == "" for v in found):
                        errors.append({
                            "number": counter,
                            "text": q["text"]
                        })
                        error_session_index = s_idx

                counter += 1

        # JIKA ADA ERROR
        if errors:
            show_unsur = (
                survey.unsur_responden
                and len(json.loads(survey.unsur_responden)) > 0
                and "Semua" not in json.loads(survey.unsur_responden)
            )

            if show_unsur:
                error_session_index += 1

            survey.unsur_responden = json.loads(
                survey.unsur_responden or '["Semua"]'
            )

            return render_template(
                "survey/isi_survey.html",
                survey=survey,
                sessions=sessions,
                errors=errors,
                old=request.form,
                selected_unsur=unsur_responden,
                error_session_index=error_session_index,
                token=token
            )

        # SIMPAN JAWABAN
        for k, v in request.form.items():
            if k.startswith("q"):
                answers[k] = v

        response = SurveyResponse(
            survey_id=survey_id,
            unsur_responden=unsur_responden,
            answers=json.dumps(answers),
            created_at=datetime.now().strftime("%Y-%m-%d %H:%M")
        )

        db.session.add(response)
        db.session.commit()

        session.pop(f"survey_{survey_id}_token", None)

        return render_template("survey/thank_you.html")

    raw_questions = json.loads(survey.questions or "[]")

    # =========================
    # FORMAT LAMA
    # =========================
    if raw_questions and "questions" not in raw_questions[0]:

        sessions = [
            {
                "title": "Survey",
                "questions": raw_questions
            }
        ]

    # =========================
    # FORMAT BARU
    # =========================
    else:
        sessions = raw_questions

    # =========================
    # GLOBAL NUMBERING
    # =========================
    counter = 1

    for session_item in sessions:

        for q in session_item["questions"]:

            q["global_index"] = counter
            counter += 1

    # 🔥 PARSE UNSUR
    survey.unsur_responden = json.loads(
        survey.unsur_responden or '["Semua"]'
    )

    return render_template(
        "survey/isi_survey.html",
        survey=survey,
        sessions=sessions,
        token=request.args.get("token")
    )

@app.route("/survey/<int:survey_id>/pembuka")
def survey_intro(survey_id):

    survey = Survey.query.get_or_404(survey_id)

    # =========================
    # CEK STATUS SURVEY
    # =========================
    if survey.status != "aktif" or not survey.is_active:
        return render_template(
            "survey/survey_closed.html",
            survey=survey
        )

    return render_template(
        "survey/pembuka.html",
        survey=survey
    )

@app.route("/survey/<int:survey_id>/start", methods=["POST"])
def survey_start(survey_id):

    survey = Survey.query.get_or_404(survey_id)

    if not survey.is_active:
        abort(404)

    # 🔐 token sekali pakai
    token = str(uuid.uuid4())

    session[f"survey_{survey_id}_token"] = token

    return redirect(
        url_for(
            "survey_responden",
            survey_id=survey_id,
            token=token
        )
    )

@app.template_filter('nl2br')
def nl2br(value):
    if not value:
        return ""
    return Markup(value.replace("\n", "<br>"))

@app.route(
    "/admin/survey/<int:survey_id>/download"
)
def download_survey(survey_id):

    if not session.get("admin"):
        return redirect(
            url_for("admin_login")
        )

    survey = Survey.query.get_or_404(
        survey_id
    )

    responses = SurveyResponse.query.filter_by(
        survey_id=survey_id
    ).all()

    wb = Workbook()

    ws = wb.active
    ws.title = "Respon Survey"

    # =====================================
    # HEADER
    # =====================================

    header_top = [
        "No",
        "Tanggal",
        "Unsur Responden"
    ]

    header_bottom = []

    sessions = json.loads(
        survey.questions or "[]"
    )

    question_keys = []

    current_col = 4

    counter = 1

    for session_item in sessions:

        questions = session_item.get(
            "questions",
            []
        )

        for q in questions:

            q_type = q.get("type")

            # =========================
            # ISIAN MULTI
            # =========================

            if q_type == "isian":

                count = q.get("count", 1)

                start_col = current_col
                end_col = current_col + count - 1

                # merge judul pertanyaan
                ws.merge_cells(
                    start_row=1,
                    start_column=start_col,
                    end_row=1,
                    end_column=end_col
                )

                ws.cell(
                    row=1,
                    column=start_col
                ).value = f"{counter}. {q.get('text')}"

                for i in range(count):

                    header_bottom.append(
                        f"Jawaban {i+1}"
                    )

                    header_top.append("")

                    question_keys.append(
                        f"q{counter}_{i}"
                    )

                    current_col += 1

            # =========================
            # PILIHAN
            # =========================

            else:

                ws.merge_cells(
                    start_row=1,
                    start_column=current_col,
                    end_row=2,
                    end_column=current_col
                )

                ws.cell(
                    row=1,
                    column=current_col
                ).value = f"{counter}. {q.get('text')}"

                question_keys.append(
                    f"q{counter}"
                )

                header_top.append("")
                header_bottom.append("")

                current_col += 1

            counter += 1

    # HEADER AWAL MERGE
    ws.merge_cells("A1:A2")
    ws.merge_cells("B1:B2")
    ws.merge_cells("C1:C2")

    # isi header utama
    ws["A1"] = "No"
    ws["B1"] = "Tanggal"
    ws["C1"] = "Unsur Responden"

    # append header row 2
    for idx, value in enumerate(header_bottom):
        ws.cell(row=2, column=4 + idx).value = value


    # =====================================
    # STYLE HEADER
    # =====================================

    for row in ws.iter_rows(
        min_row=1,
        max_row=2
    ):

        for cell in row:

            cell.font = Font(
                bold=True
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

    # =====================================
    # DATA
    # =====================================

    for idx, r in enumerate(
        responses,
        start=1
    ):

        answer_data = json.loads(
            r.answers or "{}"
        )

        row = [
            idx,
            r.created_at,
            r.unsur_responden
        ]

        for key in question_keys:

            row.append(
                answer_data.get(key, "")
            )

        ws.append(row)

    # =====================================
    # AUTO WIDTH
    # =====================================

    for col_idx in range(1, ws.max_column + 1):
        column_letter = get_column_letter(col_idx)
        length = 0

        for cell in ws[column_letter]:
            if cell.value:
                length = max(length, len(str(cell.value)))

        ws.column_dimensions[column_letter].width = length + 5

    # =====================================
    # SAVE
    # =====================================

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=
            f"survey_{survey.id}.xlsx",
        mimetype=
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/logout")
def logout():

    session.clear()

    return redirect(url_for("admin_login"))

if __name__ == "__main__":
    with app.app_context():
        db.create_all()

    app.run()
