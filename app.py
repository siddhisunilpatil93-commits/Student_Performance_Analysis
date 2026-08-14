from flask import Flask, render_template, jsonify, request
import pandas as pd
import os
import json
import re

from docx import Document
from pypdf import PdfReader
from werkzeug.utils import secure_filename


app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(
    BASE_DIR,
    "data"
)

os.makedirs(DATA_DIR, exist_ok=True)


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
# MSBTE K-SCHEME SUBJECTS
# COMPUTER ENGINEERING
# =========================================================

SEMESTER_SUBJECTS = {

    "Semester 1": [
        "Subject 1",
        "Subject 2",
        "Subject 3",
        "Subject 4"
    ],

    "Semester 2": [
        "Subject 1",
        "Subject 2",
        "Subject 3",
        "Subject 4"
    ],

    "Semester 3": [
        "Data Structure Using C",
        "Database Management System",
        "Digital Techniques",
        "Object Oriented Programming Using C++"
    ],

    "Semester 4": [
        "Java Programming",
        "Data Communication and Computer Network",
        "Microprocessor Programming",
        "Python Programming"
    ],

    "Semester 5": [
        "Subject 1",
        "Subject 2",
        "Subject 3",
        "Subject 4"
    ],

    "Semester 6": [
        "Subject 1",
        "Subject 2",
        "Subject 3",
        "Subject 4"
    ]
}


# =========================================================
# SUBJECT SHORT NAMES
# =========================================================

SUBJECT_CODES = {

    "Java Programming":
        "JPR",

    "Data Communication and Computer Network":
        "DCN",

    "Microprocessor Programming":
        "MIC",

    "Python Programming":
        "PWP",

    "Data Structure Using C":
        "DSU",

    "Database Management System":
        "DMS",

    "Digital Techniques":
        "DTE",

    "Object Oriented Programming Using C++":
        "OOP"
}


# =========================================================
# DATABASE
# =========================================================

DATABASE_FILE = os.path.join(
    DATA_DIR,
    "students.json"
)


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

            return []

    except Exception as e:

        print(
            "DATABASE READ ERROR:",
            e
        )

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

    else:
        return "F"


# =========================================================
# ATTENDANCE STATUS
# =========================================================

def calculate_attendance_status(attendance):

    if attendance >= 85:
        return "Good"

    elif attendance >= 75:
        return "Average"

    else:
        return "Low"


# =========================================================
# NORMALIZE SEMESTER
# =========================================================

def normalize_semester(value):

    value = str(value).strip().lower()

    value = value.replace(
        "semester",
        ""
    ).strip()

    value = value.replace(
        "sem",
        ""
    ).strip()

    match = re.search(
        r"\b([1-6])\b",
        value
    )

    if match:

        number = match.group(1)

        return "Semester " + number

    return ""


# =========================================================
# FIND COLUMN
# =========================================================

def find_column(columns, possible_names):

    normalized = {}

    for column in columns:

        key = re.sub(
            r"[^a-z0-9]",
            "",
            str(column).lower()
        )

        normalized[key] = column


    for name in possible_names:

        key = re.sub(
            r"[^a-z0-9]",
            "",
            name.lower()
        )

        if key in normalized:

            return normalized[key]

    return None


# =========================================================
# FIND MARKS
# =========================================================

def get_mark(row, subject):

    aliases = {

        "Java Programming": [
            "Java Programming",
            "Java",
            "JPR",
            "314317"
        ],

        "Data Communication and Computer Network": [
            "Data Communication and Computer Network",
            "Data Communication",
            "Computer Network",
            "DCN",
            "DCCN",
            "314318"
        ],

        "Microprocessor Programming": [
            "Microprocessor Programming",
            "Microprocessor",
            "Microprocessor Programming",
            "MIC",
            "314321"
        ],

        "Python Programming": [
            "Python Programming",
            "Python",
            "PWP",
            "314004"
        ],

        "Data Structure Using C": [
            "Data Structure Using C",
            "Data Structure",
            "DSU",
            "313301"
        ],

        "Database Management System": [
            "Database Management System",
            "DBMS",
            "DMS",
            "313302"
        ],

        "Digital Techniques": [
            "Digital Techniques",
            "DTE",
            "313303"
        ],

        "Object Oriented Programming Using C++": [
            "Object Oriented Programming Using C++",
            "OOP",
            "C++",
            "313304"
        ]
    }


    names = aliases.get(
        subject,
        [subject]
    )


    column = find_column(
        row.index,
        names
    )


    if column is None:
        return 0


    try:

        value = row[column]

        if pd.isna(value):
            return 0

        return float(
            str(value)
            .replace("%", "")
            .strip()
        )

    except:

        return 0


# =========================================================
# CALCULATE RESULT
# =========================================================

def calculate_result(
    student,
    semester
):

    subjects = SEMESTER_SUBJECTS[
        semester
    ]

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


    maximum_marks = (
        len(subjects) * 100
    )


    percentage = 0


    if maximum_marks > 0:

        percentage = round(
            (
                total /
                maximum_marks
            ) * 100,
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


    student["Total"] = round(
        total,
        2
    )

    student["Percentage"] = (
        percentage
    )

    student["Attendance"] = (
        attendance
    )

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

    student["Semester"] = (
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

    semester = request.args.get(
        "semester",
        ""
    ).strip()


    data = load_data()


    if semester:

        data = [

            student

            for student in data

            if student.get(
                "Semester"
            ) == semester

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


    if query == "":

        return jsonify([])


    data = load_data()

    results = []


    for student in data:

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

            results.append(
                student
            )


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


        semester = normalize_semester(
            data.get(
                "Semester",
                ""
            )
        )


        if student_id == "":

            return jsonify({
                "success": False,
                "message":
                    "Student ID is required!"
            }), 400


        if name == "":

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


        subjects = SEMESTER_SUBJECTS[
            semester
        ]


        students = load_data()


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
                        "Student already exists!"
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


        for subject in subjects:

            try:

                mark = float(
                    data.get(
                        subject,
                        0
                    )
                )

            except:

                mark = 0


            if mark < 0 or mark > 100:

                return jsonify({
                    "success": False,
                    "message":
                        f"{subject} marks must be between 0 and 100!"
                }), 400


            student[subject] = mark


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


        student["Attendance"] = (
            attendance
        )


        calculate_result(
            student,
            semester
        )


        students.append(
            student
        )


        save_data(
            students
        )


        return jsonify({

            "success": True,

            "message":
                "Student added successfully!",

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
# PROCESS DATAFRAME
# =========================================================

def process_dataframe(df):

    df.columns = [

        str(column).strip()

        for column in df.columns

    ]


    results = []


    # -----------------------------------------
    # Find common columns
    # -----------------------------------------

    id_column = find_column(
        df.columns,
        [
            "Student_ID",
            "Student ID",
            "StudentID",
            "Roll No",
            "Roll Number",
            "ID"
        ]
    )


    name_column = find_column(
        df.columns,
        [
            "Name",
            "Student Name"
        ]
    )


    gender_column = find_column(
        df.columns,
        [
            "Gender"
        ]
    )


    class_column = find_column(
        df.columns,
        [
            "Class"
        ]
    )


    semester_column = find_column(
        df.columns,
        [
            "Semester",
            "Sem"
        ]
    )


    attendance_column = find_column(
        df.columns,
        [
            "Attendance",
            "Attendance %",
            "Attendance Percentage"
        ]
    )


    for _, row in df.iterrows():

        student = {}


        # -----------------------------------------
        # BASIC DETAILS
        # -----------------------------------------

        student["Student_ID"] = str(
            row[id_column]
            if id_column
            else ""
        ).strip()


        student["Name"] = str(
            row[name_column]
            if name_column
            else ""
        ).strip()


        student["Gender"] = str(
            row[gender_column]
            if gender_column
            else ""
        ).strip()


        student["Class"] = str(
            row[class_column]
            if class_column
            else ""
        ).strip()


        # -----------------------------------------
        # SEMESTER
        # -----------------------------------------

        if semester_column:

            semester = normalize_semester(
                row[semester_column]
            )

        else:

            # Since project is 4th semester,
            # if document doesn't contain semester,
            # use Semester 4.

            semester = "Semester 4"


        if semester not in SEMESTERS:

            continue


        student["Semester"] = (
            semester
        )


        subjects = SEMESTER_SUBJECTS[
            semester
        ]


        # -----------------------------------------
        # SUBJECT MARKS
        # -----------------------------------------

        for subject in subjects:

            mark = get_mark(
                row,
                subject
            )

            if mark < 0:
                mark = 0

            if mark > 100:
                mark = 100

            student[subject] = mark


        # -----------------------------------------
        # ATTENDANCE
        # -----------------------------------------

        if attendance_column:

            try:

                attendance = float(
                    str(
                        row[
                            attendance_column
                        ]
                    )
                    .replace(
                        "%",
                        ""
                    )
                    .strip()
                )

            except:

                attendance = 0

        else:

            attendance = 0


        student["Attendance"] = (
            attendance
        )


        # -----------------------------------------
        # CALCULATE
        # -----------------------------------------

        calculate_result(
            student,
            semester
        )


        # -----------------------------------------
        # VALID STUDENT
        # -----------------------------------------

        if (
            student["Student_ID"] != ""
            and
            student["Name"] != ""
        ):

            results.append(
                student
            )


    return results


# =========================================================
# EXCEL
# =========================================================

def process_excel(file):

    df = pd.read_excel(
        file
    )

    return process_dataframe(
        df
    )


# =========================================================
# WORD
# =========================================================

def process_word(file):

    document = Document(
        file
    )


    rows = []


    for table in document.tables:

        for row in table.rows:

            values = [

                cell.text.strip()

                for cell in row.cells

            ]

            rows.append(
                values
            )


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


    return process_dataframe(
        df
    )


# =========================================================
# PDF
# =========================================================

def process_pdf(file):

    reader = PdfReader(
        file
    )


    text = ""


    for page in reader.pages:

        page_text = (
            page.extract_text()
        )

        if page_text:

            text += "\n" + page_text


    if not text.strip():

        raise Exception(
            "No readable text found in PDF."
        )


    data = {}


    # -----------------------------------------
    # Student ID
    # -----------------------------------------

    patterns = {

        "Student_ID": [
            r"Student\s*ID\s*[:\-]?\s*([A-Za-z0-9]+)",
            r"Roll\s*No\s*[:\-]?\s*([A-Za-z0-9]+)",
            r"Roll\s*Number\s*[:\-]?\s*([A-Za-z0-9]+)"
        ],

        "Name": [
            r"Student\s*Name\s*[:\-]?\s*([A-Za-z .]+)",
            r"Name\s*[:\-]?\s*([A-Za-z .]+)"
        ],

        "Gender": [
            r"Gender\s*[:\-]?\s*([A-Za-z]+)"
        ],

        "Class": [
            r"Class\s*[:\-]?\s*([A-Za-z0-9 .]+)"
        ],

        "Semester": [
            r"Semester\s*[:\-]?\s*([0-9]+)",
            r"Sem\s*[:\-]?\s*([0-9]+)"
        ],

        "Attendance": [
            r"Attendance\s*[:\-]?\s*([0-9.]+)"
        ]
    }


    for key, pattern_list in patterns.items():

        for pattern in pattern_list:

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

                break


    # -----------------------------------------
    # Semester
    # -----------------------------------------

    semester = normalize_semester(
        data.get(
            "Semester",
            "4"
        )
    )


    if not semester:

        semester = "Semester 4"


    data["Semester"] = (
        semester
    )


    # -----------------------------------------
    # Subject marks
    # -----------------------------------------

    subjects = SEMESTER_SUBJECTS[
        semester
    ]


    for subject in subjects:

        aliases = [
            subject
        ]


        if subject in SUBJECT_CODES:

            aliases.append(
                SUBJECT_CODES[
                    subject
                ]
            )


        found = False


        for alias in aliases:

            pattern = (
                re.escape(alias)
                +
                r"\s*[:\-]?\s*([0-9.]+)"
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

                found = True

                break


        if not found:

            data[subject] = 0


    # -----------------------------------------
    # Convert to dataframe
    # -----------------------------------------

    df = pd.DataFrame([
        data
    ])


    return process_dataframe(
        df
    )


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
                    "Please select a document!"

            }), 400


        file = request.files[
            "file"
        ]


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
            ".xls",
            ".docx",
            ".pdf",
            ".csv"
        ]


        if extension not in allowed:

            return jsonify({

                "success": False,

                "message":
                    "Only Excel, Word, PDF or CSV files are allowed!"

            }), 400


        # -----------------------------------------
        # PROCESS FILE
        # -----------------------------------------

        if extension in [
            ".xlsx",
            ".xls"
        ]:

            imported_students = (
                process_excel(file)
            )


        elif extension == ".docx":

            imported_students = (
                process_word(file)
            )


        elif extension == ".csv":

            df = pd.read_csv(
                file
            )

            imported_students = (
                process_dataframe(df)
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


        # -----------------------------------------
        # ADD / UPDATE
        # -----------------------------------------

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
                student_id == ""
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
                    ==
                    student_id

                    and

                    old_student.get(
                        "Semester"
                    )
                    ==
                    semester

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


        save_data(
            students
        )


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

@app.route(
    "/api/subjects"
)
def get_subjects():

    return jsonify({

        "subjects":
            SEMESTER_SUBJECTS,

        "codes":
            SUBJECT_CODES

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
