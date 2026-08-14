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

os.makedirs(DATA_DIR, exist_ok=True)

DATABASE_FILE = os.path.join(DATA_DIR, "students.json")


# =========================================================
# COLLEGE / COURSE
# =========================================================

COURSE_NAME = "Computer Engineering"
SCHEME = "MSBTE K-Scheme"


# =========================================================
# SEMESTERS
# =========================================================

SEMESTERS = [
    "Semester 1",
    "Semester 2",
    "Semester 3",
    "Semester 4",
    "Semester 5",
    "Semester 6"
]


# =========================================================
# YEAR -> SEMESTER
# =========================================================

YEAR_SEMESTERS = {
    "1st Year": ["Semester 1", "Semester 2"],
    "2nd Year": ["Semester 3", "Semester 4"],
    "3rd Year": ["Semester 5", "Semester 6"]
}


# =========================================================
# MSBTE K-SCHEME COMPUTER ENGINEERING SUBJECTS
# =========================================================

SEMESTER_SUBJECTS = {

    "Semester 1": [
        "Basic Mathematics",
        "Communication Skills (English)",
        "Basic Science - Physics and Chemistry",
        "Fundamentals of ICT",
        "Engineering Workshop Practice",
        "Yoga and Meditation",
        "Engineering Graphics"
    ],

    "Semester 2": [
        "Applied Mathematics",
        "Basic Electrical and Electronics Engineering",
        "Programming in C",
        "Linux Basics",
        "Professional Communication",
        "Social and Life Skills",
        "Web Page Designing"
    ],

    "Semester 3": [
        "Data Structure Using C",
        "Database Management System",
        "Digital Techniques",
        "Object Oriented Programming Using C++",
        "Computer Graphics",
        "Essence of Indian Constitution"
    ],

    "Semester 4": [
        "Environmental Education and Sustainability",
        "Java Programming",
        "Data Communication and Computer Network",
        "Microprocessor Programming",
        "Python Programming",
        "UI/UX Design"
    ],

    "Semester 5": [
        "Operating System",
        "Software Engineering",
        "Entrepreneurship Development and Startups",
        "Seminar and Project Initiation",
        "Advance Computer Network",
        "Cloud Computing",
        "Data Analytics"
    ],

    "Semester 6": [
        "Management",
        "Emerging Trends in Computer Engineering and Information Technology",
        "Software Testing",
        "Client Side Scripting",
        "Mobile Application Development",
        "Capstone Project",
        "Digital Forensic and Hacking Techniques",
        "Machine Learning",
        "Network and Information Security"
    ]
}


# =========================================================
# DATABASE
# =========================================================

def load_data():

    if not os.path.exists(DATABASE_FILE):
        return []

    try:
        with open(
            DATABASE_FILE,
            "r",
            encoding="utf-8"
        ) as file:

            data = json.load(file)

            if isinstance(data, list):
                return data

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

    return "F"


# =========================================================
# ATTENDANCE STATUS
# =========================================================

def calculate_attendance_status(attendance):

    if attendance >= 85:
        return "Good"

    elif attendance >= 75:
        return "Average"

    return "Low"


# =========================================================
# CALCULATE RESULT
# =========================================================

def calculate_result(student, semester):

    subjects = SEMESTER_SUBJECTS.get(
        semester,
        []
    )

    total = 0

    counted_subjects = 0

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
        counted_subjects += 1

    maximum_marks = counted_subjects * 100

    if maximum_marks > 0:

        percentage = round(
            (total / maximum_marks) * 100,
            2
        )

    else:

        percentage = 0


    try:

        attendance = float(
            student.get(
                "Attendance",
                0
            )
        )

    except:

        attendance = 0


    student["Total"] = round(
        total,
        2
    )

    student["Percentage"] = percentage

    student["Attendance"] = attendance

    student["Attendance Status"] = (
        calculate_attendance_status(
            attendance
        )
    )

    student["Grade"] = (
        calculate_grade(
            percentage
        )
    )

    student["Semester"] = semester

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

    semester = request.args.get(
        "semester",
        ""
    ).strip()

    year = request.args.get(
        "year",
        ""
    ).strip()

    data = load_data()


    # Semester filter

    if semester:

        data = [
            student
            for student in data
            if student.get("Semester") == semester
        ]


    # Year filter

    elif year:

        allowed_semesters = YEAR_SEMESTERS.get(
            year,
            []
        )

        data = [
            student
            for student in data
            if student.get("Semester")
            in allowed_semesters
        ]


    return jsonify(data)


# =========================================================
# SEARCH
# =========================================================

@app.route("/api/search")
def search_student():

    query = request.args.get(
        "q",
        ""
    ).strip().lower()

    semester = request.args.get(
        "semester",
        ""
    ).strip()


    if query == "":
        return jsonify([])


    data = load_data()

    results = []


    for student in data:

        if semester and student.get(
            "Semester"
        ) != semester:

            continue


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
            or
            query in name
        ):

            results.append(student)


    return jsonify(results)


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
                "message":
                    "No student data received!"
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


        semester = str(
            data.get(
                "Semester",
                ""
            )
        ).strip()


        # Validation

        if not student_id:

            return jsonify({
                "success": False,
                "message":
                    "Student ID / Roll Number is required!"
            }), 400


        if not name:

            return jsonify({
                "success": False,
                "message":
                    "Student Name is required!"
            }), 400


        if semester not in SEMESTERS:

            return jsonify({
                "success": False,
                "message":
                    "Invalid semester!"
            }), 400


        students = load_data()


        # Duplicate check

        for old_student in students:

            if (
                str(
                    old_student.get(
                        "Student_ID",
                        ""
                    )
                ).strip()
                == student_id

                and

                old_student.get(
                    "Semester"
                )
                == semester
            ):

                return jsonify({

                    "success": False,

                    "message":
                        f"Roll Number {student_id} already exists in {semester}!"

                }), 400


        student = {

            "Student_ID":
                student_id,

            "Name":
                name,

            "Gender":
                gender,

            "Class":
                student_class,

            "Semester":
                semester
        }


        subjects = SEMESTER_SUBJECTS[
            semester
        ]


        # Subject marks

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
                        f"{subject} marks must be between 0 and 100!"

                }), 400


            student[subject] = value


        # Attendance

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
                    "Attendance must be between 0 and 100!"

            }), 400


        student["Attendance"] = attendance


        # Calculate

        calculate_result(
            student,
            semester
        )


        students.append(student)

        save_data(students)


        return jsonify({

            "success": True,

            "message":
                f"{name} added successfully!",

            "student":
                student

        })


    except Exception as e:

        print(
            "ADD ERROR:",
            e
        )

        return jsonify({

            "success": False,

            "message":
                str(e)

        }), 500


# =========================================================
# DELETE STUDENT
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

            student

            for student in students

            if not (

                str(
                    student.get(
                        "Student_ID",
                        ""
                    )
                ).strip()
                == student_id

                and

                student.get(
                    "Semester"
                )
                == semester

            )

        ]


        if len(new_students) == len(students):

            return jsonify({

                "success": False,

                "message":
                    "Student not found!"

            }), 404


        save_data(new_students)


        return jsonify({

            "success": True,

            "message":
                "Student deleted successfully!"

        })


    except Exception as e:

        return jsonify({

            "success": False,

            "message":
                str(e)

        }), 500


# =========================================================
# PROCESS DATAFRAME
# =========================================================

def process_dataframe(df):

    df.columns = [

        str(column).strip()

        for column in df.columns

    ]


    results = []


    for _, row in df.iterrows():

        student = {}


        for field in [

            "Student_ID",
            "Name",
            "Gender",
            "Class",
            "Semester"

        ]:

            value = row.get(
                field,
                ""
            )

            if pd.isna(value):
                value = ""

            student[field] = str(
                value
            ).strip()


        semester = student["Semester"]


        # 1 -> Semester 1

        if semester.isdigit():

            semester = (
                "Semester "
                + semester
            )

            student["Semester"] = semester


        if semester not in SEMESTERS:

            continue


        subjects = SEMESTER_SUBJECTS[
            semester
        ]


        for subject in subjects:

            value = row.get(
                subject,
                0
            )

            try:

                value = float(value)

            except:

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


        calculate_result(
            student,
            semester
        )


        results.append(student)


    return results


# =========================================================
# EXCEL
# =========================================================

def process_excel(file):

    df = pd.read_excel(file)

    return process_dataframe(df)


# =========================================================
# WORD
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
            "No table found in Word document."
        )


    headers = rows[0]

    data_rows = rows[1:]


    df = pd.DataFrame(
        data_rows,
        columns=headers
    )


    return process_dataframe(df)


# =========================================================
# PDF
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
            r"(?:Name|Student\s*Name)\s*[:\-]?\s*([A-Za-z .]+)",

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

            data[key] = (
                match.group(1)
                .strip()
            )


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


    return [data]


# =========================================================
# UPLOAD DOCUMENT
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

                "message":
                    "Please select a file!"

            }), 400


        file = request.files["file"]


        if file.filename == "":

            return jsonify({

                "success": False,

                "message":
                    "No file selected!"

            }), 400


        filename = secure_filename(
            file.filename
        )


        extension = os.path.splitext(
            filename
        )[1].lower()


        allowed = [

            ".xlsx",
            ".docx",
            ".pdf"

        ]


        if extension not in allowed:

            return jsonify({

                "success": False,

                "message":
                    "Only Excel, Word and PDF files are allowed!"

            }), 400


        if extension == ".xlsx":

            imported_students = (
                process_excel(file)
            )

        elif extension == ".docx":

            imported_students = (
                process_word(file)
            )

        else:

            imported_students = (
                process_pdf(file)
            )


        if not imported_students:

            return jsonify({

                "success": False,

                "message":
                    "No valid student data found!"

            }), 400


        students = load_data()

        added = 0

        updated = 0


        for new_student in imported_students:

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
                or
                semester not in SEMESTERS
            ):

                continue


            found = False


            for index, old_student in enumerate(
                students
            ):

                if (

                    str(
                        old_student.get(
                            "Student_ID",
                            ""
                        )
                    ).strip()
                    == student_id

                    and

                    old_student.get(
                        "Semester"
                    )
                    == semester

                ):

                    students[index] = (
                        new_student
                    )

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

            "added":
                added,

            "updated":
                updated

        })


    except Exception as e:

        print(
            "UPLOAD ERROR:",
            e
        )

        return jsonify({

            "success": False,

            "message":
                str(e)

        }), 500


# =========================================================
# SUBJECT API
# =========================================================

@app.route("/api/subjects")
def get_subjects():

    return jsonify({

        "course":
            COURSE_NAME,

        "scheme":
            SCHEME,

        "semesters":
            SEMESTER_SUBJECTS

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

    print(
        "======================================"
    )

    print(
        "STUDENT PERFORMANCE ANALYSIS SYSTEM"
    )

    print(
        "COMPUTER ENGINEERING"
    )

    print(
        "MSBTE K-SCHEME"
    )

    print(
        "======================================"
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
