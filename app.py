from flask import Flask, render_template, jsonify, request
import pandas as pd
import os
import json
import re
from werkzeug.utils import secure_filename
from docx import Document
from pypdf import PdfReader

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DATABASE_FILE = os.path.join(DATA_DIR, "students.json")

os.makedirs(DATA_DIR, exist_ok=True)

# =========================================================
# YEAR / SEMESTER
# =========================================================

YEAR_SEMESTERS = {
    "1st Year": ["Semester 1", "Semester 2"],
    "2nd Year": ["Semester 3", "Semester 4"],
    "3rd Year": ["Semester 5", "Semester 6"]
}

SEMESTERS = [
    "Semester 1",
    "Semester 2",
    "Semester 3",
    "Semester 4",
    "Semester 5",
    "Semester 6"
]

# =========================================================
# MSBTE K-SCHEME COMPUTER ENGINEERING
# Core subjects used for performance analysis
# =========================================================

SEMESTER_SUBJECTS = {

    "Semester 1": [
        "Basic Mathematics",
        "Communication Skills",
        "Basic Science",
        "Fundamentals of ICT"
    ],

    "Semester 2": [
        "Applied Mathematics",
        "Basic Electrical and Electronics Engineering",
        "Programming in C",
        "Linux Basics"
    ],

    "Semester 3": [
        "Data Structure Using C",
        "Database Management System",
        "Object Oriented Programming Using C++",
        "Digital Techniques and Microprocessors"
    ],

    "Semester 4": [
        "Java Programming",
        "Data Communication and Computer Network",
        "Information Security",
        "Python Programming"
    ],

    "Semester 5": [
        "Operating System",
        "Software Engineering",
        "Advanced Database Management",
        "Data Analytics"
    ],

    "Semester 6": [
        "Web Based Application Development",
        "Mobile Application Development",
        "Internet of Things",
        "Emerging Technologies"
    ]
}

# Short names for display / CSV compatibility
SUBJECT_ALIASES = {

    "Basic Mathematics": ["BMS", "Basic Mathematics"],
    "Communication Skills": ["ENG", "Communication Skills"],
    "Basic Science": ["BSC", "Basic Science"],
    "Fundamentals of ICT": ["ICT", "Fundamentals of ICT"],

    "Applied Mathematics": ["AMS", "Applied Mathematics"],
    "Basic Electrical and Electronics Engineering":
        ["BEE", "Basic Electrical and Electronics Engineering"],
    "Programming in C": ["PIC", "Programming in C"],
    "Linux Basics": ["BLP", "Linux Basics"],

    "Data Structure Using C": ["DSU", "Data Structure Using C"],
    "Database Management System": ["DMS", "Database Management System"],
    "Object Oriented Programming Using C++":
        ["OOP", "Object Oriented Programming Using C++"],
    "Digital Techniques and Microprocessors":
        ["DTM", "Digital Techniques and Microprocessors"],

    "Java Programming": ["JPR", "Java Programming"],
    "Data Communication and Computer Network":
        ["DCN", "Data Communication and Computer Network"],
    "Information Security": ["INS", "Information Security"],
    "Python Programming": ["PWP", "Python Programming"],

    "Operating System": ["OSY", "Operating System"],
    "Software Engineering": ["STE", "Software Engineering"],
    "Advanced Database Management":
        ["ADM", "Advanced Database Management"],
    "Data Analytics": ["DAN", "Data Analytics"],

    "Web Based Application Development":
        ["WAD", "Web Based Application Development"],
    "Mobile Application Development":
        ["MAD", "Mobile Application Development"],
    "Internet of Things": ["IOT", "Internet of Things"],
    "Emerging Technologies":
        ["ET", "Emerging Technologies"]
}

# =========================================================
# DATABASE
# =========================================================

def load_data():

    if not os.path.exists(DATABASE_FILE):
        return []

    try:
        with open(DATABASE_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)

        return data if isinstance(data, list) else []

    except Exception as e:
        print("DATABASE READ ERROR:", e)
        return []


def save_data(data):

    with open(
        DATABASE_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            data,
            file,
            indent=4,
            ensure_ascii=False
        )


# =========================================================
# YEAR
# =========================================================

def get_year_from_semester(semester):

    for year, semesters in YEAR_SEMESTERS.items():

        if semester in semesters:
            return year

    return ""


# =========================================================
# GRADE
# =========================================================

def calculate_grade(percentage):

    if percentage >= 90:
        return "A+"
    elif percentage >= 80:
        return "A"
    elif percentage >= 70:
        return "B+"
    elif percentage >= 60:
        return "B"
    elif percentage >= 50:
        return "C"
    elif percentage >= 40:
        return "D"
    else:
        return "F"


# =========================================================
# ATTENDANCE
# =========================================================

def calculate_attendance_status(attendance):

    if attendance >= 85:
        return "Good"
    elif attendance >= 75:
        return "Average"
    else:
        return "Low"


# =========================================================
# RESULT
# =========================================================

def calculate_result(student, semester):

    subjects = SEMESTER_SUBJECTS.get(
        semester,
        []
    )

    total = 0

    for subject in subjects:

        try:
            mark = float(
                student.get(
                    subject,
                    0
                )
            )
        except:
            mark = 0

        total += mark

    maximum_marks = len(subjects) * 100

    percentage = 0

    if maximum_marks > 0:

        percentage = round(
            (total / maximum_marks) * 100,
            2
        )

    try:
        attendance = float(
            student.get(
                "Attendance",
                0
            )
        )
    except:
        attendance = 0

    student["Total"] = round(total, 2)
    student["Percentage"] = percentage
    student["Attendance"] = attendance

    student["Attendance Status"] = (
        calculate_attendance_status(
            attendance
        )
    )

    student["Grade"] = calculate_grade(
        percentage
    )

    student["Semester"] = semester
    student["Year"] = get_year_from_semester(
        semester
    )

    return student


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():

    return render_template(
        "index.html"
    )


# =========================================================
# GET STUDENTS
# =========================================================

@app.route("/api/students")
def get_students():

    year = request.args.get(
        "year",
        ""
    ).strip()

    semester = request.args.get(
        "semester",
        ""
    ).strip()

    students = load_data()

    if semester:

        students = [
            s for s in students
            if s.get("Semester") == semester
        ]

    elif year:

        allowed_semesters = YEAR_SEMESTERS.get(
            year,
            []
        )

        students = [
            s for s in students
            if s.get("Semester")
            in allowed_semesters
        ]

    return jsonify(students)


# =========================================================
# SEARCH
# =========================================================

@app.route("/api/search")
def search_student():

    query = request.args.get(
        "q",
        ""
    ).strip().lower()

    year = request.args.get(
        "year",
        ""
    ).strip()

    if query == "":
        return jsonify([])

    students = load_data()

    if year:

        allowed = YEAR_SEMESTERS.get(
            year,
            []
        )

        students = [
            s for s in students
            if s.get("Semester") in allowed
        ]

    result = []

    for student in students:

        student_id = str(
            student.get(
                "Student_ID",
                ""
            )
        ).lower()

        name = str(
            student.get(
                "Name",
                ""
            )
        ).lower()

        if (
            query in student_id
            or query in name
        ):

            result.append(student)

    return jsonify(result)


# =========================================================
# SUBJECT API
# =========================================================

@app.route("/api/subjects")
def get_subjects():

    return jsonify({
        "years": YEAR_SEMESTERS,
        "semesters": SEMESTER_SUBJECTS
    })


# =========================================================
# ADD STUDENT
# =========================================================

@app.route(
    "/api/add_student",
    methods=["POST"]
)
def add_student():

    try:

        data = request.get_json()

        if not data:

            return jsonify({
                "success": False,
                "message": "No data received."
            }), 400

        student_id = str(
            data.get(
                "Student_ID",
                ""
            )
        ).strip()

        name = str(
            data.get(
                "Name",
                ""
            )
        ).strip()

        gender = str(
            data.get(
                "Gender",
                ""
            )
        ).strip()

        student_class = str(
            data.get(
                "Class",
                ""
            )
        ).strip()

        year = str(
            data.get(
                "Year",
                ""
            )
        ).strip()

        semester = str(
            data.get(
                "Semester",
                ""
            )
        ).strip()

        if not student_id:

            return jsonify({
                "success": False,
                "message": "Student ID is required."
            }), 400

        if not name:

            return jsonify({
                "success": False,
                "message": "Student Name is required."
            }), 400

        if year not in YEAR_SEMESTERS:

            return jsonify({
                "success": False,
                "message": "Invalid year."
            }), 400

        if semester not in YEAR_SEMESTERS[year]:

            return jsonify({
                "success": False,
                "message":
                    "Selected semester does not belong to selected year."
            }), 400

        students = load_data()

        # Duplicate check
        for old in students:

            if (
                str(
                    old.get(
                        "Student_ID",
                        ""
                    )
                ).strip() == student_id

                and

                old.get(
                    "Semester"
                ) == semester
            ):

                return jsonify({
                    "success": False,
                    "message":
                        f"Student ID {student_id} already exists in {semester}."
                }), 400

        student = {

            "Student_ID": student_id,
            "Name": name,
            "Gender": gender,
            "Class": student_class,
            "Year": year,
            "Semester": semester
        }

        subjects = SEMESTER_SUBJECTS[
            semester
        ]

        for subject in subjects:

            value = data.get(
                subject,
                0
            )

            try:
                value = float(value)
            except:
                value = 0

            if value < 0 or value > 100:

                return jsonify({
                    "success": False,
                    "message":
                        f"{subject} marks must be between 0 and 100."
                }), 400

            student[subject] = value

        try:

            attendance = float(
                data.get(
                    "Attendance",
                    0
                )
            )

        except:

            attendance = 0

        if attendance < 0 or attendance > 100:

            return jsonify({
                "success": False,
                "message":
                    "Attendance must be between 0 and 100."
            }), 400

        student["Attendance"] = attendance

        student = calculate_result(
            student,
            semester
        )

        students.append(student)

        save_data(students)

        return jsonify({

            "success": True,

            "message":
                f"{name} added successfully.",

            "student": student

        })

    except Exception as e:

        print("ADD ERROR:", e)

        return jsonify({

            "success": False,
            "message": str(e)

        }), 500


# =========================================================
# DELETE
# =========================================================

@app.route(
    "/api/delete_student",
    methods=["POST"]
)
def delete_student():

    try:

        data = request.get_json()

        student_id = str(
            data.get(
                "Student_ID",
                ""
            )
        ).strip()

        semester = str(
            data.get(
                "Semester",
                ""
            )
        ).strip()

        students = load_data()

        new_students = [

            s for s in students

            if not (
                str(
                    s.get(
                        "Student_ID",
                        ""
                    )
                ).strip() == student_id

                and

                s.get(
                    "Semester"
                ) == semester
            )
        ]

        if len(new_students) == len(students):

            return jsonify({
                "success": False,
                "message": "Student not found."
            }), 404

        save_data(new_students)

        return jsonify({
            "success": True,
            "message": "Student deleted successfully."
        })

    except Exception as e:

        return jsonify({
            "success": False,
            "message": str(e)
        }), 500


# =========================================================
# EXCEL PROCESS
# =========================================================

def process_excel(file):

    df = pd.read_excel(file)

    return process_dataframe(df)


# =========================================================
# WORD PROCESS
# =========================================================

def process_word(file):

    document = Document(file)

    rows = []

    for table in document.tables:

        for row in table.rows:

            rows.append([
                cell.text.strip()
                for cell in row.cells
            ])

    if not rows:

        raise Exception(
            "No table found in Word file."
        )

    headers = rows[0]

    data_rows = rows[1:]

    df = pd.DataFrame(
        data_rows,
        columns=headers
    )

    return process_dataframe(df)


# =========================================================
# PDF PROCESS
# =========================================================

def process_pdf(file):

    reader = PdfReader(file)

    text = ""

    for page in reader.pages:

        page_text = page.extract_text()

        if page_text:
            text += "\n" + page_text

    if not text.strip():

        raise Exception(
            "No readable text found in PDF."
        )

    data = {}

    patterns = {

        "Student_ID":
            r"(?:Student\s*ID|Roll\s*No|Roll\s*Number)\s*[:\-]?\s*([A-Za-z0-9]+)",

        "Name":
            r"(?:Student\s*Name|Name)\s*[:\-]?\s*([A-Za-z .]+)",

        "Gender":
            r"Gender\s*[:\-]?\s*([A-Za-z]+)",

        "Class":
            r"Class\s*[:\-]?\s*([A-Za-z0-9 .]+)",

        "Semester":
            r"Semester\s*[:\-]?\s*([0-9]+)",

        "Attendance":
            r"Attendance\s*[:\-]?\s*([0-9.]+)"
    }

    for key, pattern in patterns.items():

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:
            data[key] = match.group(1).strip()

    if "Semester" in data:

        if data["Semester"].isdigit():

            data["Semester"] = (
                "Semester "
                + data["Semester"]
            )

    semester = data.get(
        "Semester",
        ""
    )

    if semester in SEMESTERS:

        subjects = SEMESTER_SUBJECTS[
            semester
        ]

        for subject in subjects:

            pattern = (
                re.escape(subject)
                + r"\s*[:\-]?\s*([0-9.]+)"
            )

            match = re.search(
                pattern,
                text,
                re.IGNORECASE
            )

            if match:

                data[subject] = (
                    match.group(1)
                )

    if not data:

        raise Exception(
            "Student data could not be extracted."
        )

    return [
        data
    ]


# =========================================================
# DATAFRAME
# =========================================================

def process_dataframe(df):

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    results = []

    for _, row in df.iterrows():

        student = {}

        for field in [
            "Student_ID",
            "Name",
            "Gender",
            "Class",
            "Year",
            "Semester"
        ]:

            student[field] = str(
                row.get(
                    field,
                    ""
                )
            ).strip()

        semester = student.get(
            "Semester",
            ""
        )

        # Semester number
        if semester.isdigit():

            semester = (
                "Semester "
                + semester
            )

            student["Semester"] = semester

        # If only Year is supplied
        if semester not in SEMESTERS:

            continue

        student["Year"] = (
            get_year_from_semester(
                semester
            )
        )

        subjects = SEMESTER_SUBJECTS[
            semester
        ]

        for subject in subjects:

            value = row.get(
                subject,
                None
            )

            # Check abbreviation aliases
            if pd.isna(value) or value is None:

                for alias in SUBJECT_ALIASES.get(
                    subject,
                    []
                ):

                    if alias in df.columns:

                        value = row.get(
                            alias,
                            0
                        )

                        break

            try:
                value = float(value)
            except:
                value = 0

            if value < 0 or value > 100:
                value = 0

            student[subject] = value

        attendance = row.get(
            "Attendance",
            0
        )

        try:

            attendance = float(
                str(attendance)
                .replace("%", "")
                .strip()
            )

        except:

            attendance = 0

        student["Attendance"] = attendance

        student = calculate_result(
            student,
            semester
        )

        results.append(student)

    return results


# =========================================================
# DOCUMENT UPLOAD
# =========================================================

@app.route(
    "/api/upload_document",
    methods=["POST"]
)
def upload_document():

    try:

        if "file" not in request.files:

            return jsonify({
                "success": False,
                "message": "Please select a file."
            }), 400

        file = request.files["file"]

        if file.filename == "":

            return jsonify({
                "success": False,
                "message": "No file selected."
            }), 400

        filename = secure_filename(
            file.filename
        )

        extension = os.path.splitext(
            filename
        )[1].lower()

        if extension not in [
            ".xlsx",
            ".docx",
            ".pdf"
        ]:

            return jsonify({
                "success": False,
                "message":
                    "Only Excel, Word and PDF are allowed."
            }), 400

        if extension == ".xlsx":

            imported = process_excel(file)

        elif extension == ".docx":

            imported = process_word(file)

        else:

            imported = process_pdf(file)

        if not imported:

            return jsonify({
                "success": False,
                "message":
                    "No valid student data found."
            }), 400

        students = load_data()

        added = 0
        updated = 0

        for new_student in imported:

            student_id = str(
                new_student.get(
                    "Student_ID",
                    ""
                )
            ).strip()

            semester = new_student.get(
                "Semester",
                ""
            )

            if (
                not student_id
                or semester not in SEMESTERS
            ):
                continue

            new_student = calculate_result(
                new_student,
                semester
            )

            found = False

            for i, old in enumerate(students):

                if (
                    str(
                        old.get(
                            "Student_ID",
                            ""
                        )
                    ).strip()
                    == student_id

                    and

                    old.get(
                        "Semester"
                    )
                    == semester
                ):

                    students[i] = new_student
                    updated += 1
                    found = True
                    break

            if not found:

                students.append(
                    new_student
                )

                added += 1

        save_data(students)

        return jsonify({

            "success": True,

            "message":
                f"Upload successful! Added: {added}, Updated: {updated}",

            "added": added,
            "updated": updated

        })

    except Exception as e:

        print(
            "UPLOAD ERROR:",
            e
        )

        return jsonify({

            "success": False,
            "message": str(e)

        }), 500


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
