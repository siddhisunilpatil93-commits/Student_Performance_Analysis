from flask import Flask, render_template, request, jsonify
import pandas as pd
import os
import re
import json
import pdfplumber
from docx import Document

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)


# =========================================================
# MSBTE K-SCHEME COMPUTER ENGINEERING SUBJECTS
# =========================================================

SEMESTERS = {

    "Semester 1": {
        "academic_year": "2025-26",
        "subjects": [
            ("311001", "Fundamentals of ICT"),
            ("311002", "Engineering Workshop Practice"),
            ("311003", "Yoga and Meditation"),
            ("311008", "Engineering Graphics"),
            ("311302", "Basic Mathematics"),
            ("311303", "Communication Skills (English)"),
            ("311305", "Basic Science")
        ]
    },

    "Semester 2": {
        "academic_year": "2025-26",
        "subjects": [
            ("312001", "Linux Basics"),
            ("312002", "Professional Communication"),
            ("312003", "Social and Life Skills"),
            ("312004", "Web Page Designing"),
            ("312301", "Applied Mathematics"),
            ("312302", "Basic Electrical and Electronics Engineering"),
            ("312303", "Programming in C")
        ]
    },

    "Semester 3": {
        "academic_year": "2026-27",
        "subjects": [
            ("313301", "Data Structure Using C"),
            ("313302", "Database Management System"),
            ("313303", "Digital Techniques"),
            ("313304", "Object Oriented Programming Using C++")
        ]
    },

    "Semester 4": {
        "academic_year": "2026-27",
        "subjects": [
            ("314301", "Environmental Education and Sustainability"),
            ("314317", "Java Programming"),
            ("314318", "Data Communication and Computer Network"),
            ("314319", "Information Security"),
            ("314320", "Mathematics for Machine Learning"),
            ("314321", "Microprocessor Programming"),
            ("314316", "Probability and Statistics")
        ]
    },

    "Semester 5": {
        "academic_year": "2027-28",
        "subjects": [
            ("315319", "Operating System"),
            ("315323", "Software Engineering"),
            ("315324", "Advanced Database Management"),
            ("315325", "Cloud Computing"),
            ("315326", "Data Analytics"),
            ("315321", "Advanced Computer Network"),
            ("315330", "AI & ML Algorithm"),
            ("315329", "Natural Language Processing"),
            ("315327", "Cloud Computing for Data Science"),
            ("315332", "Software Engineering and Testing"),
            ("315301", "Management")
        ]
    },

    "Semester 6": {
        "academic_year": "2027-28",
        "subjects": [
            ("316313", "Emerging Trends in Computer Engineering and IT"),
            ("316314", "Software Testing"),
            ("316315", "Digital Forensic and Hacking Techniques"),
            ("316316", "Machine Learning"),
            ("316317", "Network and Information Security"),
            ("316318", "Big Data Analytics"),
            ("316319", "Principles of Image Processing"),
            ("316320", "Advanced Algorithm in AI & ML"),
            ("316321", "Data Warehousing with Mining Techniques"),
            ("316322", "Image Processing"),
            ("316323", "Reinforcement Learning"),
            ("316324", "Software Engineering and Testing for Big Data"),
            ("316325", "Wireless and Mobile Network")
        ]
    }
}


# =========================================================
# FILE NAME FOR EACH SEMESTER
# =========================================================

def semester_file(semester):

    safe = semester.lower().replace(" ", "_")

    return os.path.join(
        DATA_DIR,
        safe + ".json"
    )


# =========================================================
# SUBJECT COLUMN MATCHING
# =========================================================

def normalize_text(value):

    if value is None:
        return ""

    value = str(value).strip().lower()

    value = re.sub(
        r"[^a-z0-9]+",
        "_",
        value
    )

    return value.strip("_")


def subject_aliases(code, name):

    aliases = [
        code,
        name,
        normalize_text(name),
        normalize_text(code)
    ]

    words = name.split()

    if words:
        initials = "".join(
            word[0]
            for word in words
            if word
        )

        aliases.append(initials)

    return [
        normalize_text(x)
        for x in aliases
    ]


# =========================================================
# GRADE
# =========================================================

def calculate_grade(percentage):

    percentage = float(percentage)

    if percentage >= 90:
        return "A+"

    if percentage >= 80:
        return "A"

    if percentage >= 70:
        return "B+"

    if percentage >= 60:
        return "B"

    if percentage >= 50:
        return "C"

    if percentage >= 40:
        return "D"

    return "F"


# =========================================================
# PROCESS STUDENT
# =========================================================

def process_student(row, semester):

    subjects = SEMESTERS[semester]["subjects"]

    result = {}

    # -----------------------------------------------------
    # BASIC INFORMATION
    # -----------------------------------------------------

    def get_value(names):

        normalized = {
            normalize_text(k): v
            for k, v in row.items()
        }

        for name in names:

            key = normalize_text(name)

            if key in normalized:
                return normalized[key]

        return ""

    result["Student_ID"] = get_value([
        "Student_ID",
        "Student ID",
        "ID",
        "Roll No",
        "Roll Number",
        "Roll"
    ])

    result["Name"] = get_value([
        "Name",
        "Student Name",
        "Student_Name"
    ])

    result["Gender"] = get_value([
        "Gender",
        "Sex"
    ])

    result["Class"] = get_value([
        "Class",
        "Division"
    ])

    result["Academic_Year"] = SEMESTERS[
        semester
    ]["academic_year"]

    result["Semester"] = semester

    # -----------------------------------------------------
    # SUBJECT MARKS
    # -----------------------------------------------------

    total = 0
    count = 0

    for code, subject_name in subjects:

        aliases = subject_aliases(
            code,
            subject_name
        )

        value = ""

        for key, original_value in row.items():

            normalized_key = normalize_text(key)

            if normalized_key in aliases:

                value = original_value
                break

            # code/name partial matching
            if (
                normalize_text(code) in normalized_key
                or
                normalize_text(subject_name) in normalized_key
            ):
                value = original_value
                break

        try:
            mark = float(value)

        except:
            mark = 0

        mark = max(
            0,
            min(100, mark)
        )

        result[code] = mark

        total += mark
        count += 1

    # -----------------------------------------------------
    # TOTAL / PERCENTAGE
    # -----------------------------------------------------

    result["Total"] = round(
        total,
        2
    )

    if count > 0:

        percentage = (
            total / (count * 100)
        ) * 100

    else:
        percentage = 0

    result["Percentage"] = round(
        percentage,
        2
    )

    # -----------------------------------------------------
    # ATTENDANCE
    # -----------------------------------------------------

    attendance = get_value([
        "Attendance",
        "Attendance %",
        "Attendance Percentage",
        "Attendance_Percentage"
    ])

    try:
        attendance = float(
            attendance
        )
    except:
        attendance = 0

    attendance = max(
        0,
        min(100, attendance)
    )

    result["Attendance"] = round(
        attendance,
        2
    )

    result["Attendance_Status"] = (
        "Good"
        if attendance >= 75
        else "Low"
    )

    result["Grade"] = calculate_grade(
        percentage
    )

    return result


# =========================================================
# LOAD DATA
# =========================================================

def load_data(semester):

    path = semester_file(
        semester
    )

    if not os.path.exists(path):
        return []

    try:

        with open(
            path,
            "r",
            encoding="utf-8"
        ) as f:

            return json.load(f)

    except:
        return []


# =========================================================
# SAVE DATA
# =========================================================

def save_data(
    semester,
    students
):

    path = semester_file(
        semester
    )

    with open(
        path,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            students,
            f,
            indent=4,
            ensure_ascii=False
        )


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():

    return render_template(
        "index.html"
    )


# =========================================================
# SUBJECT API
# =========================================================

@app.route(
    "/api/subjects"
)
def subjects():

    semester = request.args.get(
        "semester",
        "Semester 3"
    )

    if semester not in SEMESTERS:

        return jsonify({
            "success": False,
            "message": "Invalid semester"
        })

    return jsonify({
        "success": True,
        "academic_year":
            SEMESTERS[semester]["academic_year"],
        "subjects": [
            {
                "code": code,
                "name": name
            }
            for code, name
            in SEMESTERS[semester]["subjects"]
        ]
    })


# =========================================================
# GET STUDENTS
# =========================================================

@app.route(
    "/api/students"
)
def students_api():

    semester = request.args.get(
        "semester",
        "Semester 3"
    )

    if semester not in SEMESTERS:

        return jsonify([])

    return jsonify(
        load_data(semester)
    )


# =========================================================
# ADD STUDENT
# =========================================================

@app.route(
    "/api/add_student",
    methods=["POST"]
)
def add_student():

    data = request.get_json()

    semester = data.get(
        "Semester",
        "Semester 3"
    )

    if semester not in SEMESTERS:

        return jsonify({
            "success": False,
            "message": "Invalid semester"
        })

    students = load_data(
        semester
    )

    processed = process_student(
        data,
        semester
    )

    if not processed["Student_ID"]:
        return jsonify({
            "success": False,
            "message": "Student ID is required."
        })

    # duplicate ID
    for student in students:

        if str(
            student.get("Student_ID")
        ).lower() == str(
            processed["Student_ID"]
        ).lower():

            return jsonify({
                "success": False,
                "message": "Student ID already exists."
            })

    students.append(
        processed
    )

    save_data(
        semester,
        students
    )

    return jsonify({
        "success": True,
        "message": "Student added successfully."
    })


# =========================================================
# EXCEL / WORD / PDF UPLOAD
# =========================================================

@app.route(
    "/api/upload",
    methods=["POST"]
)
def upload_file():

    file = request.files.get(
        "file"
    )

    semester = request.form.get(
        "semester",
        "Semester 3"
    )

    if not file:

        return jsonify({
            "success": False,
            "message": "Please select a file."
        })

    if semester not in SEMESTERS:

        return jsonify({
            "success": False,
            "message": "Invalid semester."
        })

    filename = file.filename.lower()

    try:

        # =================================================
        # EXCEL
        # =================================================

        if filename.endswith(
            ".xlsx"
        ):

            df = pd.read_excel(
                file
            )

            rows = df.to_dict(
                orient="records"
            )

        # =================================================
        # WORD
        # =================================================

        elif filename.endswith(
            ".docx"
        ):

            document = Document(
                file
            )

            rows = []

            for table in document.tables:

                if len(table.rows) < 2:
                    continue

                headers = [
                    cell.text.strip()
                    for cell
                    in table.rows[0].cells
                ]

                for row in table.rows[1:]:

                    values = [
                        cell.text.strip()
                        for cell
                        in row.cells
                    ]

                    if len(values) != len(headers):
                        continue

                    rows.append(
                        dict(
                            zip(
                                headers,
                                values
                            )
                        )
                    )

                if rows:
                    break

        # =================================================
        # PDF
        # =================================================

        elif filename.endswith(
            ".pdf"
        ):

            rows = []

            with pdfplumber.open(
                file
            ) as pdf:

                for page in pdf.pages:

                    tables = page.extract_tables()

                    for table in tables:

                        if not table:
                            continue

                        headers = [
                            str(x or "").strip()
                            for x in table[0]
                        ]

                        for values in table[1:]:

                            if not values:
                                continue

                            if len(values) != len(headers):
                                continue

                            rows.append(
                                dict(
                                    zip(
                                        headers,
                                        values
                                    )
                                )
                            )

        else:

            return jsonify({
                "success": False,
                "message":
                    "Only Excel, Word and PDF files are allowed."
            })

        if not rows:

            return jsonify({
                "success": False,
                "message":
                    "No table/student data found in file."
            })

        students = load_data(
            semester
        )

        added = 0
        skipped = 0

        existing_ids = {
            str(
                x.get("Student_ID")
            ).lower()
            for x in students
        }

        for row in rows:

            student = process_student(
                row,
                semester
            )

            sid = str(
                student.get(
                    "Student_ID",
                    ""
                )
            ).strip()

            if not sid:
                skipped += 1
                continue

            if sid.lower() in existing_ids:
                skipped += 1
                continue

            students.append(
                student
            )

            existing_ids.add(
                sid.lower()
            )

            added += 1

        save_data(
            semester,
            students
        )

        return jsonify({
            "success": True,
            "message":
                f"{added} student(s) imported successfully. "
                f"{skipped} skipped."
        })

    except Exception as e:

        return jsonify({
            "success": False,
            "message":
                "File processing error: "
                + str(e)
        })


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            5000
        )
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
