# app.py
from flask import (
    Flask, render_template, request, jsonify, send_file
)
import mysql.connector
import json
import io
import asyncio

# 用 Chromium 列印 PDF（保留 detail 排版）
from playwright.async_api import async_playwright

app = Flask(__name__)

DB_CONFIG = {
    "host": "127.0.0.1",
    "user": "aacsb_user",
    "password": "password",
    "database": "aacsb",
    "charset": "utf8mb4",
    "collation": "utf8mb4_unicode_ci",
}

def get_conn():
    return mysql.connector.connect(**DB_CONFIG)


# =========================
# Home
# =========================
@app.get("/")
def home():
    return render_template("index.html")


# ============================================================
# A) 學生問卷 Student
# ============================================================
@app.get("/student")
def student_form():
    return render_template("form_student.html")


@app.post("/student/submit")
def student_submit():
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "No JSON body received"}), 400

    basic = data.get("basic") or {}
    unit = data.get("internship_unit_satisfaction") or {}
    course = data.get("internship_course_satisfaction") or {}
    feedback = data.get("feedback") or {}
    other = data.get("other_suggestions") or {}

    required_basic = [
        "student_id", "class_name", "student_name",
        "internship_year", "internship_times",
        "internship_type", "internship_org"
    ]
    for k in required_basic:
        if basic.get(k) in (None, "", []):
            return jsonify({"ok": False, "error": f"basic.{k} 必填"}), 400

    if basic.get("internship_type") == "其他":
        if not (basic.get("internship_type_other") or "").strip():
            return jsonify({"ok": False, "error": "選了 internship_type=其他，basic.internship_type_other 必填"}), 400

    if unit.get("salary_system") == "其他":
        if not (unit.get("salary_system_other") or "").strip():
            return jsonify({"ok": False, "error": "選了 salary_system=其他，unit.salary_system_other 必填"}), 400

    must_unit = [
        "unit_content_quality", "unit_environment", "unit_supervisor_guidance",
        "unit_interaction", "unit_overtime_hours", "unit_overall",
        "salary_system", "salary_monthly_equiv", "overtime_pay"
    ]
    for k in must_unit:
        if unit.get(k) in (None, "", []):
            return jsonify({"ok": False, "error": f"internship_unit_satisfaction.{k} 必填"}), 400

    must_course = [
        "course_admin_support", "course_safety_training", "course_advisor_help",
        "course_task_support", "course_goal_match", "course_positive_help", "course_thesis_match"
    ]
    for k in must_course:
        if course.get(k) in (None, "", []):
            return jsonify({"ok": False, "error": f"internship_course_satisfaction.{k} 必填"}), 400

    if feedback.get("org_understanding") in (None, "", []):
        return jsonify({"ok": False, "error": "feedback.org_understanding 必填"}), 400

    certs_arr = feedback.get("certs_to_improve") or []
    exp_arr = feedback.get("experience_use") or []
    if not isinstance(certs_arr, list) or len(certs_arr) == 0:
        return jsonify({"ok": False, "error": "feedback.certs_to_improve 至少選 1 項"}), 400
    if not isinstance(exp_arr, list) or len(exp_arr) == 0:
        return jsonify({"ok": False, "error": "feedback.experience_use 至少選 1 項"}), 400

    for q in ["q1", "q2", "q3", "q4"]:
        if not (other.get(q) or "").strip():
            return jsonify({"ok": False, "error": f"other_suggestions.{q} 必填"}), 400

    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute("""
            INSERT INTO student_satisfaction_survey (
                student_id, class_name, student_name,
                internship_year, internship_times,
                internship_type, internship_type_other,
                internship_org,

                unit_content_quality, unit_environment, unit_supervisor_guidance,
                unit_interaction, unit_overtime_hours, unit_overall,
                salary_system, salary_system_other, salary_monthly_equiv, overtime_pay,

                course_admin_support, course_safety_training, course_advisor_help,
                course_task_support, course_goal_match, course_positive_help, course_thesis_match,

                helpful_abilities, certs_to_improve, experience_use,
                org_understanding,

                suggestion_q1, suggestion_q2, suggestion_q3, suggestion_q4
            ) VALUES (
                %s,%s,%s,
                %s,%s,
                %s,%s,
                %s,

                %s,%s,%s,
                %s,%s,%s,
                %s,%s,%s,%s,

                %s,%s,%s,
                %s,%s,%s,%s,

                %s,%s,%s,
                %s,

                %s,%s,%s,%s
            )
        """, (
            str(basic.get("student_id")),
            str(basic.get("class_name")),
            str(basic.get("student_name")),

            int(basic.get("internship_year")),
            int(basic.get("internship_times")),

            str(basic.get("internship_type")),
            (basic.get("internship_type_other") or None),

            str(basic.get("internship_org")),

            int(unit.get("unit_content_quality")),
            int(unit.get("unit_environment")),
            int(unit.get("unit_supervisor_guidance")),

            int(unit.get("unit_interaction")),
            int(unit.get("unit_overtime_hours")),
            int(unit.get("unit_overall")),

            str(unit.get("salary_system")),
            (unit.get("salary_system_other") if unit.get("salary_system") == "其他" else None),
            str(unit.get("salary_monthly_equiv")),
            str(unit.get("overtime_pay")),

            int(course.get("course_admin_support")),
            int(course.get("course_safety_training")),
            int(course.get("course_advisor_help")),

            int(course.get("course_task_support")),
            int(course.get("course_goal_match")),
            int(course.get("course_positive_help")),
            int(course.get("course_thesis_match")),

            json.dumps(feedback.get("helpful_abilities") or [], ensure_ascii=False),
            json.dumps(feedback.get("certs_to_improve") or [], ensure_ascii=False),
            json.dumps(feedback.get("experience_use") or [], ensure_ascii=False),

            int(feedback.get("org_understanding")),

            str(other.get("q1")),
            str(other.get("q2")),
            str(other.get("q3")),
            str(other.get("q4")),
        ))
        conn.commit()

        sid = cur.lastrowid
        return jsonify({"ok": True, "student_survey_id": sid, "preview_url": f"/admin/student/{sid}"})
    except Exception as e:
        conn.rollback()
        return jsonify({"ok": False, "error": f"{type(e).__name__}: {e}"}), 500
    finally:
        try: cur.close()
        except: pass
        try: conn.close()
        except: pass


@app.get("/admin/student")
def admin_student_list():
    conn = get_conn()
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT student_survey_id, student_id, student_name, class_name, internship_org, submitted_at
            FROM student_satisfaction_survey
            ORDER BY student_survey_id DESC
            LIMIT 200
        """)
        rows = cur.fetchall()
        return render_template("admin_student_list.html", rows=rows)
    finally:
        try: conn.close()
        except: pass


@app.get("/admin/student/<int:survey_id>")
def admin_student_detail(survey_id: int):
    conn = get_conn()
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM student_satisfaction_survey WHERE student_survey_id=%s", (survey_id,))
        survey = cur.fetchone()
        if not survey:
            return "找不到此學生問卷", 404

        for k in ["helpful_abilities", "certs_to_improve", "experience_use"]:
            try:
                survey[k] = json.loads(survey[k]) if survey.get(k) else []
            except Exception:
                survey[k] = []

        return render_template("admin_student_detail.html", survey=survey)
    finally:
        try: conn.close()
        except: pass


# ============================================================
# B) 雇主問卷 Employer
# ============================================================
@app.get("/employer")
def employer_form():
    return render_template("form_employer.html")


@app.post("/employer/submit")
def employer_submit():
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "No JSON body received"}), 400

    industry = data.get("industry")
    industry_other = data.get("industry_other")
    company_name = data.get("company_name")
    job_title = data.get("job_title")
    other_suggestions = data.get("other_suggestions")

    students = data.get("students", [])
    performance = data.get("performance", {})
    course = data.get("course", {})
    improvements = data.get("improvements", [])
    cooperations = data.get("cooperations", [])
    cooperation_note = data.get("cooperation_note")

    if not company_name or not job_title or not other_suggestions:
        return jsonify({"ok": False, "error": "company_name/job_title/other_suggestions 必填"}), 400

    if industry == "其他" and (not industry_other or str(industry_other).strip() == ""):
        return jsonify({"ok": False, "error": "選了產業其他，industry_other 必填"}), 400

    for k in [f"q{i}" for i in range(1, 12)]:
        if k not in performance:
            return jsonify({"ok": False, "error": f"缺少 performance.{k}"}), 400

    for k in ["course_match_industry_needs", "student_meet_core_competency", "internship_admin_satisfaction"]:
        if k not in course:
            return jsonify({"ok": False, "error": f"缺少 course.{k}"}), 400

    conn = get_conn()
    cur = conn.cursor()
    try:
        conn.start_transaction()

        cur.execute("""
            INSERT INTO employer_satisfaction_survey
            (industry, industry_other, company_name, job_title, other_suggestions)
            VALUES (%s, %s, %s, %s, %s)
        """, (
            industry,
            industry_other if industry == "其他" else None,
            company_name,
            job_title,
            other_suggestions
        ))
        employer_survey_id = cur.lastrowid

        if students:
            cur.executemany("""
                INSERT INTO employer_survey_student
                (employer_survey_id, department, student_name)
                VALUES (%s, %s, %s)
            """, [
                (employer_survey_id, s.get("department"), s.get("student_name"))
                for s in students
                if (s.get("department") and s.get("student_name"))
            ])

        cur.execute("""
            INSERT INTO employer_survey_performance
            (employer_survey_id, q1,q2,q3,q4,q5,q6,q7,q8,q9,q10,q11)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (
            employer_survey_id,
            int(performance["q1"]), int(performance["q2"]), int(performance["q3"]),
            int(performance["q4"]), int(performance["q5"]), int(performance["q6"]),
            int(performance["q7"]), int(performance["q8"]), int(performance["q9"]),
            int(performance["q10"]), int(performance["q11"])
        ))

        cur.execute("""
            INSERT INTO employer_survey_course
            (employer_survey_id, course_match_industry_needs, student_meet_core_competency, internship_admin_satisfaction)
            VALUES (%s, %s, %s, %s)
        """, (
            employer_survey_id,
            int(course["course_match_industry_needs"]),
            int(course["student_meet_core_competency"]),
            int(course["internship_admin_satisfaction"])
        ))

        if improvements:
            cur.executemany("""
                INSERT INTO employer_survey_improvement
                (employer_survey_id, improvement_item, improvement_note)
                VALUES (%s, %s, %s)
            """, [
                (employer_survey_id, it.get("improvement_item"), it.get("improvement_note"))
                for it in improvements
                if it.get("improvement_item")
            ])

        if cooperations:
            cur.executemany("""
                INSERT INTO employer_survey_cooperation
                (employer_survey_id, cooperation_item, cooperation_note)
                VALUES (%s, %s, %s)
            """, [
                (employer_survey_id, item, (cooperation_note if item == "其他" else None))
                for item in cooperations
            ])

        conn.commit()
        return jsonify({"ok": True, "employer_survey_id": employer_survey_id, "preview_url": f"/admin/employer/{employer_survey_id}"})

    except Exception as e:
        conn.rollback()
        return jsonify({"ok": False, "error": f"{type(e).__name__}: {e}"}), 500
    finally:
        try: cur.close()
        except: pass
        try: conn.close()
        except: pass


@app.get("/admin/employer")
def admin_employer_list():
    conn = get_conn()
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT employer_survey_id, company_name, job_title, industry, created_at
            FROM employer_satisfaction_survey
            ORDER BY employer_survey_id DESC
            LIMIT 200
        """)
        rows = cur.fetchall()
        return render_template("admin_employer_list.html", rows=rows)
    finally:
        try: conn.close()
        except: pass


@app.get("/admin/employer/<int:survey_id>")
def admin_employer_detail_preview(survey_id: int):
    conn = get_conn()
    try:
        cur = conn.cursor(dictionary=True)

        cur.execute("""
            SELECT employer_survey_id, industry, industry_other, company_name, job_title, other_suggestions, created_at
            FROM employer_satisfaction_survey
            WHERE employer_survey_id = %s
        """, (survey_id,))
        survey = cur.fetchone()
        if not survey:
            return "找不到此雇主問卷", 404

        cur.execute("""
            SELECT department, student_name
            FROM employer_survey_student
            WHERE employer_survey_id = %s
            ORDER BY student_id ASC
        """, (survey_id,))
        students = cur.fetchall()

        cur.execute("""
            SELECT q1,q2,q3,q4,q5,q6,q7,q8,q9,q10,q11
            FROM employer_survey_performance
            WHERE employer_survey_id = %s
        """, (survey_id,))
        performance = cur.fetchone() or {}

        cur.execute("""
            SELECT course_match_industry_needs, student_meet_core_competency, internship_admin_satisfaction
            FROM employer_survey_course
            WHERE employer_survey_id = %s
        """, (survey_id,))
        course = cur.fetchone() or {}

        cur.execute("""
            SELECT improvement_item, improvement_note
            FROM employer_survey_improvement
            WHERE employer_survey_id = %s
            ORDER BY improvement_id ASC
        """, (survey_id,))
        improvements = cur.fetchall()

        cur.execute("""
            SELECT cooperation_item, cooperation_note
            FROM employer_survey_cooperation
            WHERE employer_survey_id = %s
            ORDER BY cooperation_id ASC
        """, (survey_id,))
        cooperations = cur.fetchall()

        return render_template(
            "admin_employer_detail.html",
            survey=survey,
            students=students,
            performance=performance,
            course=course,
            improvements=improvements,
            cooperations=cooperations
        )
    finally:
        try: conn.close()
        except: pass


# ============================================================
# PDF Export（學生/雇主）— 用 Chromium 列印
# ============================================================
async def _render_urls_to_single_pdf(urls):
    """
    把多個 URL 的頁面列印成 PDF，合併為一份（每個 URL 之間自動分頁）
    作法：把每個 detail HTML 以 iframe 包起來，再 print。
    """
    # 用 iframe 組合頁
    frames = []
    for u in urls:
        frames.append(f"""
          <div class="wrap">
            <iframe src="{u}" class="frame"></iframe>
          </div>
          <div class="pb"></div>
        """)

    shell_html = f"""
    <!doctype html>
    <html><head>
      <meta charset="utf-8">
      <style>
        html,body{{margin:0;padding:0;}}
        .frame{{width:100%; height:1300px; border:0;}}
        .pb{{page-break-after:always;}}
        @media print {{
          .pb{{page-break-after:always;}}
        }}
      </style>
    </head>
    <body>
      {''.join(frames)}
    </body></html>
    """

    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.set_content(shell_html, wait_until="load")
        # 等 iframe 載入
        await page.wait_for_timeout(800)
        pdf_bytes = await page.pdf(
            format="A4",
            print_background=True,
            margin={"top": "12mm", "bottom": "12mm", "left": "12mm", "right": "12mm"},
        )
        await browser.close()
        return pdf_bytes


@app.post("/admin/student/export_pdf")
def admin_student_export_pdf():
    ids = request.form.getlist("ids")
    ids = [int(x) for x in ids if str(x).isdigit()]
    if not ids:
        return "請先勾選至少一筆再下載", 400

    # 這些 URL 指向「detail（同排版）」頁
    base = request.host_url.rstrip("/")
    urls = [f"{base}/admin/student/{i}" for i in ids]

    pdf_bytes = asyncio.run(_render_urls_to_single_pdf(urls))
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name="student_surveys_selected.pdf"
    )


@app.post("/admin/employer/export_pdf")
def admin_employer_export_pdf():
    ids = request.form.getlist("ids")
    ids = [int(x) for x in ids if str(x).isdigit()]
    if not ids:
        return "請先勾選至少一筆再下載", 400

    base = request.host_url.rstrip("/")
    urls = [f"{base}/admin/employer/{i}" for i in ids]

    pdf_bytes = asyncio.run(_render_urls_to_single_pdf(urls))
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name="employer_surveys_selected.pdf"
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
