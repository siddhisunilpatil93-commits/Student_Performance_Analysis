from flask import Flask, render_template, jsonify, request
import pandas as pd
import os
import json
import re
from werkzeug.utils import secure_filename

# Word / PDF
from docx import Document
from pypdf import PdfReader

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

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
# SUBJECT NAMES
# =========================================================
# हे temporary names आहेत.
# तुझे actual 4 subjects नंतर इथे बदलायचे.

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
        "Subject 1",
        "Subject 2",
        "Subject 3",
        "Subject 4"
    ],

    "Semester 4": [
        "Subject 1",
        "Subject 2",
        "Subject 3",
        "Subject 4"
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
# JSON DATABASE
# =========================================================
# CSV नाही.
# सर्व student data JSON मध्ये save होईल.

DATABASE_FILE = os.path.join(
    DATA_DIR,
    "students.json"
)


# =========================================================
# DATABASE FUNCTIONS
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

            return []

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
# CALCULATE RESULT
# =========================================================

def calculate_result(student, semester):

    subjects = SEMESTER_SUBJECTS[semester]

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


    student["Total"] = total

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
# GET ALL STUDENTS
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
# SEARCH STUDENT
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


        semester = str(
            data.get(
                "Semester",
                ""
            )
        ).strip()


        # -----------------------------------------
        # Validation
        # -----------------------------------------

        if student_id == "":

            return jsonify({
                "success": False,
                "message":
                    "Student ID / Roll Number is required!"
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


        # -----------------------------------------
        # Read database
        # -----------------------------------------

        students = load_data()


        # -----------------------------------------
        # Same Student + Same Semester check
        # -----------------------------------------

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


        # -----------------------------------------
        # New student
        # -----------------------------------------

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


        # -----------------------------------------
        # Subjects
        # -----------------------------------------

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


        # -----------------------------------------
        # Attendance
        # -----------------------------------------

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


        # -----------------------------------------
        # Calculate
        # -----------------------------------------

        student = calculate_result(
            student,
            semester
        )


        # -----------------------------------------
        # Save
        # -----------------------------------------

        students.append(
            student
        )

        save_data(
            students
        )


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
# EDIT STUDENT
# =========================================================

@app.route(
    "/api/edit_student",
    methods=["POST"]
)
def edit_student():

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


        if student_id == "":
            return jsonify({
                "success": False,
                "message":
                    "Student ID is required!"
            }), 400


        if semester not in SEMESTERS:

            return jsonify({
                "success": False,
                "message":
                    "Invalid semester!"
            }), 400


        students = load_data()

        subjects = SEMESTER_SUBJECTS[
            semester
        ]


        found = False


        for student in students:

            if (
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
            ):

                found = True


                # Basic details

                student["Name"] = str(
                    data.get(
                        "Name",
                        student.get("Name", "")
                    )
                ).strip()


                student["Gender"] = str(
                    data.get(
                        "Gender",
                        student.get("Gender", "")
                    )
                ).strip()


                student["Class"] = str(
                    data.get(
                        "Class",
                        student.get("Class", "")
                    )
                ).strip()


                # Subjects

                for subject in subjects:

                    try:

                        mark = float(
                            data.get(
                                subject,
                                student.get(
                                    subject,
                                    0
                                )
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


                # Attendance

                try:

                    attendance = float(
                        data.get(
                            "Attendance",
                            student.get(
                                "Attendance",
                                0
                            )
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


                # Recalculate

                calculate_result(
                    student,
                    semester
                )


                break


        if not found:

            return jsonify({
                "success": False,
                "message":
                    "Student not found!"
            }), 404


        save_data(
            students
        )


        return jsonify({

            "success": True,

            "message":
                "Student updated successfully!"

        })


    except Exception as e:

        print(
            "EDIT ERROR:",
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


        save_data(
            new_students
        )


        return jsonify({

            "success": True,

            "message":
                "Student deleted successfully!"

        })


    except Exception as e:

        print(
            "DELETE ERROR:",
            e
        )

        return jsonify({

            "success": False,

            "message":
                str(e)

        }), 500


# =========================================================
# EXCEL UPLOAD
# =========================================================

def process_excel(file):

    try:

        df = pd.read_excel(
            file
        )

        return process_dataframe(
            df
        )

    except Exception as e:

        raise Exception(
            "Excel reading error: "
            + str(e)
        )


# =========================================================
# WORD UPLOAD
# =========================================================

def process_word(file):

    try:

        document = Document(
            file
        )

        rows = []


        # -----------------------------------------
        # Read tables
        # -----------------------------------------

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


    except Exception as e:

        raise Exception(
            "Word reading error: "
            + str(e)
        )


# =========================================================
# PDF UPLOAD
# =========================================================

def process_pdf(file):

    try:

        reader = PdfReader(
            file
        )

        text = ""


        for page in reader.pages:

            page_text = page.extract_text()

            if page_text:

                text += "\n" + page_text


        if text.strip() == "":

            raise Exception(
                "No readable text found in PDF."
            )


        # -----------------------------------------
        # Basic key-value extraction
        # -----------------------------------------

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


        # -----------------------------------------
        # Semester conversion
        # -----------------------------------------

        if "Semester" in data:

            sem_number = (
                data["Semester"]
                .strip()
            )

            if sem_number.isdigit():

                data["Semester"] = (
                    "Semester "
                    + sem_number
                )


        # -----------------------------------------
        # Subject marks
        # -----------------------------------------

        for semester in SEMESTERS:

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
                "Student data could not be extracted from PDF."
            )


        return [
            data
        ]


    except Exception as e:

        raise Exception(
            "PDF reading error: "
            + str(e)
        )


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


        # -----------------------------------------
        # Basic details
        # -----------------------------------------

        for field in [

            "Student_ID",
            "Name",
            "Gender",
            "Class",
            "Semester"

        ]:

            student[field] = str(
                row.get(
                    field,
                    ""
                )
            ).strip()


        semester = student["Semester"]


        # Convert formats like 1, 2, 3
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


        # -----------------------------------------
        # Subjects
        # -----------------------------------------

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


        # -----------------------------------------
        # Attendance
        # -----------------------------------------

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


        # -----------------------------------------
        # Calculate
        # -----------------------------------------

        student = calculate_result(
            student,
            semester
        )


        results.append(
            student
        )


    return results


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

        extension = (
            os.path.splitext(
                filename
            )[1]
            .lower()
        )


        # -----------------------------------------
        # Allowed files
        # -----------------------------------------

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


        # -----------------------------------------
        # Process
        # -----------------------------------------

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


        # -----------------------------------------
        # Existing database
        # -----------------------------------------

        students = load_data()


        added = 0

        updated = 0


        # -----------------------------------------
        # Add / Update
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


        # -----------------------------------------
        # Save
        # -----------------------------------------

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
# SEMESTER SUBJECTS API
# =========================================================

@app.route(
    "/api/subjects"
)
def get_subjects():

    return jsonify(
        SEMESTER_SUBJECTS
    )


# =========================================================
# RUN SERVER
# =========================================================

if __name__ == "__main__":

    print()
    print(
        "========================================"
    )
    print(
        " STUDENT PERFORMANCE DASHBOARD"
    )
    print(
        "========================================"
    )

    print(
        "Semesters:",
        ", ".join(SEMESTERS)
    )

    print(
        "Database:",
        DATABASE_FILE
    )

    print(
        "========================================"
    )


    # Render PORT
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
