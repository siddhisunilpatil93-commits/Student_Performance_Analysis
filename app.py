from flask import Flask, render_template, jsonify, request, redirect, url_for, session
import pandas as pd
import os

app = Flask(__name__)

# =========================================================
# SECRET KEY
# =========================================================

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "student-performance-secret-key"
)

BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))


# =========================================================
# YEAR + SEMESTER
# =========================================================

SEMESTERS = {
    "1st Year": ["Semester 1", "Semester 2"],
    "2nd Year": ["Semester 3", "Semester 4"],
    "3rd Year": ["Semester 5", "Semester 6"]
}

SEMESTER_YEAR = {
    "Semester 1": "1st Year",
    "Semester 2": "1st Year",
    "Semester 3": "2nd Year",
    "Semester 4": "2nd Year",
    "Semester 5": "3rd Year",
    "Semester 6": "3rd Year"
}


# =========================================================
# EXCEL FILES
# =========================================================

SEMESTER_FILES = {
    "Semester 1": "semester_1.xlsx",
    "Semester 2": "semester_2.xlsx",
    "Semester 3": "semester_3.xlsx",
    "Semester 4": "semester_4.xlsx",
    "Semester 5": "semester_5.xlsx",
    "Semester 6": "semester_6.xlsx"
}


# =========================================================
# SUBJECTS
# =========================================================

SUBJECTS = {

    "Semester 1": [
        "Basic Mathematics",
        "Communication Skills",
        "Engineering Physics",
        "Engineering Chemistry",
        "Basic Science"
    ],

    "Semester 2": [
        "Applied Mathematics",
        "Engineering Graphics",
        "Basic Electrical Engineering",
        "Programming in C",
        "Web Page Design"
    ],

    "Semester 3": [
        "Object Oriented Programming",
        "Data Structure",
        "Database Management System",
        "Computer Networks",
        "Operating System"
    ],

    "Semester 4": [
        "Java Programming",
        "Data Communication and Network",
        "Microprocessor",
        "Software Engineering",
        "Python Programming"
    ],

    "Semester 5": [
        "Advanced Java",
        "Web Based Application Development",
        "Software Testing",
        "Computer Security",
        "Project Management"
    ],

    "Semester 6": [
        "Mobile Application Development",
        "Cloud Computing",
        "Artificial Intelligence",
        "Internet of Things",
        "Major Project"
    ]
}


# =========================================================
# LOGIN
# =========================================================

LOGIN_USERNAME = os.environ.get(
    "ADMIN_USERNAME",
    "silicon"
)

LOGIN_PASSWORD = os.environ.get(
    "ADMIN_PASSWORD",
    "patil"
)


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def valid_semester(semester):
    return semester in SEMESTER_FILES


def excel_path(semester):
    return os.path.join(
        BASE_FOLDER,
        SEMESTER_FILES[semester]
    )


def logged_in():
    return session.get("logged_in", False) is True


def page_login_check():
    if not logged_in():
        return redirect(url_for("login"))
    return None


def api_login_check():
    if not logged_in():
        return jsonify({
            "success": False,
            "message": "Login required."
        }), 401

    return None


# =========================================================
# GRADE
# =========================================================

def calculate_grade(percentage):

    try:
        percentage = float(percentage)
    except Exception:
        return "F"

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
# READ EXCEL
# =========================================================

def read_excel(semester):

    if not valid_semester(semester):
        return pd.DataFrame()

    file_path = excel_path(semester)

    if not os.path.exists(file_path):
        return pd.DataFrame()

    try:

        df = pd.read_excel(file_path)

        if df is None:
            return pd.DataFrame()

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        return df

    except Exception as error:

        print("Excel read error:", error)

        return pd.DataFrame()


# =========================================================
# PROCESS DATA
# =========================================================

def process_data(df, semester):

    df = df.copy()

    required_columns = [
        "Student_ID",
        "Name",
        "Gender",
        "Class",
        "Attendance"
    ]

    # Add missing basic columns
    for column in required_columns:

        if column not in df.columns:

            if column == "Attendance":
                df[column] = 0
            else:
                df[column] = ""

    # -----------------------------------------------------
    # Student ID
    # -----------------------------------------------------

    df["Student_ID"] = (
        df["Student_ID"]
        .fillna("")
        .astype(str)
        .str.replace(
            r"\.0$",
            "",
            regex=True
        )
        .str.strip()
    )

    # -----------------------------------------------------
    # Name
    # -----------------------------------------------------

    df["Name"] = (
        df["Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # -----------------------------------------------------
    # Gender
    # -----------------------------------------------------

    df["Gender"] = (
        df["Gender"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # -----------------------------------------------------
    # Class
    # -----------------------------------------------------

    df["Class"] = (
        df["Class"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # -----------------------------------------------------
    # Attendance
    # -----------------------------------------------------

    df["Attendance"] = (
        pd.to_numeric(
            df["Attendance"],
            errors="coerce"
        )
        .fillna(0)
        .clip(0, 100)
        .round(2)
    )

    # -----------------------------------------------------
    # Subjects
    # -----------------------------------------------------

    subjects = SUBJECTS[semester]

    for subject in subjects:

        if subject not in df.columns:
            df[subject] = 0

        df[subject] = (
            pd.to_numeric(
                df[subject],
                errors="coerce"
            )
            .fillna(0)
            .clip(0, 100)
            .round(2)
        )

    # -----------------------------------------------------
    # Total
    # -----------------------------------------------------

    df["Total"] = (
        df[subjects]
        .sum(axis=1)
        .round(2)
    )

    # -----------------------------------------------------
    # Percentage
    # -----------------------------------------------------

    max_marks = len(subjects) * 100

    if max_marks > 0:

        df["Percentage"] = (
            df["Total"] / max_marks * 100
        ).round(2)

    else:

        df["Percentage"] = 0

    # -----------------------------------------------------
    # Attendance Status
    # -----------------------------------------------------

    df["Attendance Status"] = df[
        "Attendance"
    ].apply(
        lambda x: "Good" if float(x) >= 75 else "Bad"
    )

    # -----------------------------------------------------
    # Grade
    # -----------------------------------------------------

    df["Grade"] = df[
        "Percentage"
    ].apply(calculate_grade)

    return df


# =========================================================
# SAVE EXCEL
# =========================================================

def save_excel(df, semester):

    processed = process_data(
        df,
        semester
    )

    columns = [
        "Student_ID",
        "Name",
        "Gender",
        "Class"
    ]

    columns += SUBJECTS[semester]

    columns += [
        "Attendance"
    ]

    # Ensure all columns exist
    for column in columns:

        if column not in processed.columns:
            processed[column] = ""

    processed = processed[columns]

    processed.to_excel(
        excel_path(semester),
        index=False
    )


# =========================================================
# LOGIN
# =========================================================

@app.route(
    "/login",
    methods=["GET", "POST"]
)
def login():

    if logged_in():
        return redirect(url_for("home"))

    if request.method == "POST":

        username = (
            request.form
            .get("username", "")
            .strip()
        )

        password = request.form.get(
            "password",
            ""
        )

        if (
            username == LOGIN_USERNAME
            and password == LOGIN_PASSWORD
        ):

            session.clear()

            session["logged_in"] = True

            return redirect(
                url_for("home")
            )

        return render_template(
            "index.html",
            login_page=True,
            login_error="Invalid username or password."
        )

    return render_template(
        "index.html",
        login_page=True
    )


# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():

    session.clear()

    return redirect(
        url_for("login")
    )


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():

    check = page_login_check()

    if check:
        return check

    return render_template(
        "index.html",
        login_page=False
    )


# =========================================================
# SUBJECT API
# =========================================================

@app.route("/api/subjects")
def api_subjects():

    check = api_login_check()

    if check:
        return check

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    if not valid_semester(semester):

        return jsonify({
            "success": False,
            "message": "Invalid semester."
        }), 400

    return jsonify({

        "success": True,

        "branch": "Computer Engineering",

        "year": SEMESTER_YEAR[semester],

        "semester": semester,

        "subjects": [
            {
                "code": subject,
                "name": subject
            }
            for subject in SUBJECTS[semester]
        ]

    })


# =========================================================
# STUDENTS API
# =========================================================

@app.route("/api/students")
def api_students():

    check = api_login_check()

    if check:
        return check

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    if not valid_semester(semester):

        return jsonify([])

    df = read_excel(semester)

    if df.empty:
        return jsonify([])

    df = process_data(
        df,
        semester
    )

    columns = [
        "Student_ID",
        "Name",
        "Gender",
        "Class"
    ]

    columns += SUBJECTS[semester]

    columns += [
        "Total",
        "Percentage",
        "Attendance",
        "Attendance Status",
        "Grade"
    ]

    # Make sure every column exists
    for column in columns:

        if column not in df.columns:
            df[column] = ""

    result = (
        df[columns]
        .fillna("")
        .to_dict("records")
    )

    return jsonify(result)


# =========================================================
# ANALYTICS API
# =========================================================

@app.route("/api/analytics")
def api_analytics():

    check = api_login_check()

    if check:
        return check

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    if not valid_semester(semester):

        return jsonify({
            "total_students": 0,
            "average_percentage": 0,
            "top_performer": "-",
            "average_attendance": 0,
            "subjects": {},
            "grades": {}
        })

    df = read_excel(semester)

    if df.empty:

        return jsonify({
            "total_students": 0,
            "average_percentage": 0,
            "top_performer": "-",
            "average_attendance": 0,
            "subjects": {},
            "grades": {}
        })

    df = process_data(
        df,
        semester
    )

    if df.empty:

        return jsonify({
            "total_students": 0,
            "average_percentage": 0,
            "top_performer": "-",
            "average_attendance": 0,
            "subjects": {},
            "grades": {}
        })

    # -----------------------------------------------------
    # TOP PERFORMER
    # -----------------------------------------------------

    top_index = df["Percentage"].idxmax()

    top_name = str(
        df.loc[
            top_index,
            "Name"
        ]
    )

    # -----------------------------------------------------
    # SUBJECT AVERAGES
    # -----------------------------------------------------

    subject_data = {}

    for subject in SUBJECTS[semester]:

        subject_data[subject] = round(
            float(df[subject].mean()),
            2
        )

    # -----------------------------------------------------
    # GRADES
    # -----------------------------------------------------

    grade_data = (
        df["Grade"]
        .value_counts()
        .to_dict()
    )

    # -----------------------------------------------------
    # RESPONSE
    # -----------------------------------------------------

    return jsonify({

        "total_students": int(len(df)),

        "average_percentage": round(
            float(df["Percentage"].mean()),
            2
        ),

        "top_performer": top_name,

        "average_attendance": round(
            float(df["Attendance"].mean()),
            2
        ),

        "subjects": subject_data,

        "grades": grade_data

    })


# =========================================================
# UPLOAD EXCEL
# =========================================================

@app.route(
    "/api/upload_excel",
    methods=["POST"]
)
def upload_excel():

    check = api_login_check()

    if check:
        return check

    try:

        semester = request.form.get(
            "semester",
            ""
        ).strip()

        file = request.files.get("file")

        # -------------------------------------------------
        # Validate semester
        # -------------------------------------------------

        if not valid_semester(semester):

            return jsonify({
                "success": False,
                "message": "Invalid semester selected."
            }), 400

        # -------------------------------------------------
        # Validate file
        # -------------------------------------------------

        if file is None:

            return jsonify({
                "success": False,
                "message": "Please select Excel file."
            }), 400

        if not file.filename:

            return jsonify({
                "success": False,
                "message": "Please select Excel file."
            }), 400

        filename = file.filename.lower()

        if not filename.endswith(".xlsx"):

            return jsonify({
                "success": False,
                "message": "Only .xlsx Excel file is allowed."
            }), 400

        # -------------------------------------------------
        # Read uploaded Excel
        # -------------------------------------------------

        df = pd.read_excel(file)

        if df is None or df.empty:

            return jsonify({
                "success": False,
                "message": "Uploaded Excel is empty."
            }), 400

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        # -------------------------------------------------
        # Required columns
        # -------------------------------------------------

        if "Student_ID" not in df.columns:

            return jsonify({
                "success": False,
                "message": "Student_ID column is required."
            }), 400

        if "Name" not in df.columns:

            return jsonify({
                "success": False,
                "message": "Name column is required."
            }), 400

        # -------------------------------------------------
        # Save
        # -------------------------------------------------

        save_excel(
            df,
            semester
        )

        return jsonify({

            "success": True,

            "message":
                f"{semester} Excel uploaded successfully."

        })

    except Exception as error:

        print(
            "UPLOAD ERROR:",
            error
        )

        return jsonify({

            "success": False,

            "message":
                f"Excel upload failed: {str(error)}"

        }), 500


# =========================================================
# ADD STUDENT
# =========================================================

@app.route(
    "/api/add_student",
    methods=["POST"]
)
def add_student():

    check = api_login_check()

    if check:
        return check

    try:

        data = request.get_json() or {}

        semester = str(
            data.get(
                "semester",
                "Semester 1"
            )
        ).strip()

        if not valid_semester(semester):

            return jsonify({
                "success": False,
                "message": "Invalid semester."
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

        if not student_id:

            return jsonify({
                "success": False,
                "message": "Student ID is required."
            })

        if not name:

            return jsonify({
                "success": False,
                "message": "Student Name is required."
            })

        df = read_excel(semester)

        # -------------------------------------------------
        # Duplicate ID
        # -------------------------------------------------

        if not df.empty:

            if "Student_ID" not in df.columns:
                df["Student_ID"] = ""

            existing_ids = (
                df["Student_ID"]
                .fillna("")
                .astype(str)
                .str.replace(
                    r"\.0$",
                    "",
                    regex=True
                )
                .str.strip()
                .tolist()
            )

            if student_id in existing_ids:

                return jsonify({
                    "success": False,
                    "message":
                        "Student ID already exists."
                })

        # -------------------------------------------------
        # New row
        # -------------------------------------------------

        row = {

            "Student_ID":
                student_id,

            "Name":
                name,

            "Gender":
                data.get(
                    "Gender",
                    ""
                ),

            "Class":
                data.get(
                    "Class",
                    ""
                ),

            "Attendance":
                data.get(
                    "Attendance",
                    0
                )
        }

        for subject in SUBJECTS[semester]:

            row[subject] = data.get(
                subject,
                0
            )

        # -------------------------------------------------
        # Append
        # -------------------------------------------------

        new_row = pd.DataFrame([row])

        df = pd.concat(
            [
                df,
                new_row
            ],
            ignore_index=True
        )

        save_excel(
            df,
            semester
        )

        return jsonify({

            "success": True,

            "message":
                "Student added successfully."

        })

    except Exception as error:

        print(
            "ADD ERROR:",
            error
        )

        return jsonify({

            "success": False,

            "message":
                str(error)

        }), 500


# =========================================================
# EDIT STUDENT
# =========================================================

@app.route(
    "/api/edit_student",
    methods=["POST"]
)
def edit_student():

    check = api_login_check()

    if check:
        return check

    try:

        data = request.get_json() or {}

        semester = str(
            data.get(
                "semester",
                "Semester 1"
            )
        ).strip()

        student_id = str(
            data.get(
                "Student_ID",
                ""
            )
        ).strip()

        if not valid_semester(semester):

            return jsonify({
                "success": False,
                "message": "Invalid semester."
            })

        df = read_excel(semester)

        if df.empty:

            return jsonify({
                "success": False,
                "message": "Student not found."
            })

        if "Student_ID" not in df.columns:

            return jsonify({
                "success": False,
                "message": "Student_ID column not found."
            })

        df["Student_ID"] = (
            df["Student_ID"]
            .fillna("")
            .astype(str)
            .str.replace(
                r"\.0$",
                "",
                regex=True
            )
            .str.strip()
        )

        matches = df.index[
            df["Student_ID"] == student_id
        ].tolist()

        if not matches:

            return jsonify({
                "success": False,
                "message": "Student not found."
            })

        index = matches[0]

        editable_columns = [
            "Name",
            "Gender",
            "Class",
            "Attendance"
        ]

        editable_columns += SUBJECTS[semester]

        for column in editable_columns:

            if column in data:

                df.loc[
                    index,
                    column
                ] = data[column]

        save_excel(
            df,
            semester
        )

        return jsonify({

            "success": True,

            "message":
                "Student updated successfully."

        })

    except Exception as error:

        print(
            "EDIT ERROR:",
            error
        )

        return jsonify({

            "success": False,

            "message":
                str(error)

        }), 500


# =========================================================
# DELETE STUDENT
# =========================================================

@app.route(
    "/api/delete_student",
    methods=["POST"]
)
def delete_student():

    check = api_login_check()

    if check:
        return check

    try:

        data = request.get_json() or {}

        semester = str(
            data.get(
                "semester",
                "Semester 1"
            )
        ).strip()

        student_id = str(
            data.get(
                "Student_ID",
                ""
            )
        ).strip()

        if not valid_semester(semester):

            return jsonify({
                "success": False,
                "message": "Invalid semester."
            })

        df = read_excel(semester)

        if df.empty:

            return jsonify({
                "success": False,
                "message": "Student not found."
            })

        df["Student_ID"] = (
            df["Student_ID"]
            .fillna("")
            .astype(str)
            .str.replace(
                r"\.0$",
                "",
                regex=True
            )
            .str.strip()
        )

        new_df = df[
            df["Student_ID"] != student_id
        ].copy()

        if len(new_df) == len(df):

            return jsonify({
                "success": False,
                "message": "Student not found."
            })

        save_excel(
            new_df,
            semester
        )

        return jsonify({

            "success": True,

            "message":
                "Student deleted successfully."

        })

    except Exception as error:

        print(
            "DELETE ERROR:",
            error
        )

        return jsonify({

            "success": False,

            "message":
                str(error)

        }), 500


# =========================================================
# HEALTH CHECK
# =========================================================

@app.route("/health")
def health():

    return jsonify({
        "status": "ok",
        "application": "Student Performance Analysis System"
    })


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=int(
            os.environ.get(
                "PORT",
                5000
            )
        )
    )
