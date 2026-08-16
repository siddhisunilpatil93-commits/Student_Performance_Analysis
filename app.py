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
# SEMESTER EXCEL FILES
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
# HELPERS
# =========================================================

def valid_semester(semester):
    return semester in SEMESTER_FILES


def path_for(semester):

    return os.path.join(
        BASE_FOLDER,
        SEMESTER_FILES[semester]
    )


def is_logged_in():

    return session.get("logged_in") is True


def login_required_page():

    if not is_logged_in():
        return redirect(url_for("login"))

    return None


def login_required_api():

    if not is_logged_in():

        return jsonify({
            "success": False,
            "message": "Login required."
        }), 401

    return None


# =========================================================
# CREATE EMPTY EXCEL FILES IF MISSING
# =========================================================

def create_empty_file(semester):

    if not valid_semester(semester):
        return

    file_path = path_for(semester)

    if os.path.exists(file_path):
        return

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

    df = pd.DataFrame(columns=columns)

    try:
        df.to_excel(
            file_path,
            index=False
        )

        print(
            f"Created missing Excel file: {file_path}"
        )

    except Exception as e:

        print(
            f"Could not create {file_path}: {e}"
        )


def create_all_files():

    for semester in SEMESTER_FILES:
        create_empty_file(semester)


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

def read_df(semester):

    if not valid_semester(semester):
        return pd.DataFrame()

    create_empty_file(semester)

    file_path = path_for(semester)

    if not os.path.exists(file_path):
        return pd.DataFrame()

    try:

        df = pd.read_excel(
            file_path
        )

        if df is None:
            return pd.DataFrame()

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        return df

    except Exception as e:

        print(
            "Excel read error:",
            e
        )

        return pd.DataFrame()


# =========================================================
# PROCESS DATA
# =========================================================

def process(df, semester):

    df = df.copy()

    subjects = SUBJECTS[semester]

    # -----------------------------------------------------
    # BASIC COLUMNS
    # -----------------------------------------------------

    basic_columns = [
        "Student_ID",
        "Name",
        "Gender",
        "Class",
        "Attendance"
    ]

    for column in basic_columns:

        if column not in df.columns:

            if column == "Attendance":
                df[column] = 0
            else:
                df[column] = ""

    # -----------------------------------------------------
    # STUDENT ID
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
    # NAME
    # -----------------------------------------------------

    df["Name"] = (
        df["Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # -----------------------------------------------------
    # GENDER
    # -----------------------------------------------------

    df["Gender"] = (
        df["Gender"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # -----------------------------------------------------
    # CLASS
    # -----------------------------------------------------

    df["Class"] = (
        df["Class"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # -----------------------------------------------------
    # ATTENDANCE
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
    # SUBJECTS
    # -----------------------------------------------------

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
    # TOTAL
    # -----------------------------------------------------

    df["Total"] = (
        df[subjects]
        .sum(axis=1)
        .round(2)
    )

    # -----------------------------------------------------
    # PERCENTAGE
    # -----------------------------------------------------

    total_marks = len(subjects) * 100

    if total_marks > 0:

        df["Percentage"] = (
            df["Total"]
            / total_marks
            * 100
        ).round(2)

    else:

        df["Percentage"] = 0

    # -----------------------------------------------------
    # ATTENDANCE STATUS
    # -----------------------------------------------------

    df["Attendance Status"] = (
        df["Attendance"]
        .apply(
            lambda x:
            "Good"
            if float(x) >= 75
            else "Bad"
        )
    )

    # -----------------------------------------------------
    # GRADE
    # -----------------------------------------------------

    df["Grade"] = (
        df["Percentage"]
        .apply(calculate_grade)
    )

    return df


# =========================================================
# SAVE EXCEL
# =========================================================

def save_df(df, semester):

    if not valid_semester(semester):
        return False

    processed = process(
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

    for column in columns:

        if column not in processed.columns:

            if column == "Attendance":
                processed[column] = 0
            else:
                processed[column] = ""

    processed = processed[columns]

    try:

        processed.to_excel(
            path_for(semester),
            index=False
        )

        return True

    except Exception as e:

        print(
            "Excel save error:",
            e
        )

        return False


# =========================================================
# LOGIN
# =========================================================

@app.route(
    "/login",
    methods=["GET", "POST"]
)
def login():

    if is_logged_in():

        return redirect(
            url_for("home")
        )

    if request.method == "POST":

        username = (
            request.form
            .get(
                "username",
                ""
            )
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

    check = login_required_page()

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

    check = login_required_api()

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

        "year":
            SEMESTER_YEAR[semester],

        "semester":
            semester,

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

    check = login_required_api()

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

    df = read_df(semester)

    if df.empty:
        return jsonify([])

    df = process(
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

    return jsonify(
        df[columns]
        .fillna("")
        .to_dict(
            orient="records"
        )
    )


# =========================================================
# ANALYTICS API
# =========================================================

@app.route("/api/analytics")
def api_analytics():

    check = login_required_api()

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

    df = read_df(semester)

    if df.empty:

        return jsonify({

            "total_students": 0,

            "average_percentage": 0,

            "top_performer": "-",

            "average_attendance": 0,

            "subjects": {},

            "grades": {}

        })

    df = process(
        df,
        semester
    )

    # -----------------------------------------------------
    # TOP PERFORMER
    # -----------------------------------------------------

    top_performer = "-"

    if not df.empty:

        top_index = df["Percentage"].idxmax()

        top_performer = str(
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

        if len(df) > 0:

            subject_data[subject] = round(
                float(
                    df[subject].mean()
                ),
                2
            )

        else:

            subject_data[subject] = 0

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

        "total_students":
            int(len(df)),

        "average_percentage":
            round(
                float(
                    df["Percentage"].mean()
                ),
                2
            ),

        "top_performer":
            top_performer,

        "average_attendance":
            round(
                float(
                    df["Attendance"].mean()
                ),
                2
            ),

        "subjects":
            subject_data,

        "grades":
            grade_data

    })


# =========================================================
# UPLOAD EXCEL
# =========================================================

@app.route(
    "/api/upload_excel",
    methods=["POST"]
)
def upload_excel():

    check = login_required_api()

    if check:
        return check

    try:

        semester = request.form.get(
            "semester",
            "Semester 1"
        )

        file = request.files.get(
            "file"
        )

        # -------------------------------------------------
        # VALID SEMESTER
        # -------------------------------------------------

        if not valid_semester(semester):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester."

            }), 400

        # -------------------------------------------------
        # FILE CHECK
        # -------------------------------------------------

        if not file:

            return jsonify({

                "success": False,

                "message":
                    "Excel file select करा."

            }), 400

        filename = (
            file.filename
            .lower()
            .strip()
        )

        if not filename.endswith(
            (".xlsx", ".xls")
        ):

            return jsonify({

                "success": False,

                "message":
                    "Only Excel file (.xlsx / .xls) allowed."

            }), 400

        # -------------------------------------------------
        # READ UPLOADED FILE
        # -------------------------------------------------

        df = pd.read_excel(file)

        if df is None:

            return jsonify({

                "success": False,

                "message":
                    "Excel file empty आहे."

            }), 400

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        # -------------------------------------------------
        # REQUIRED COLUMNS
        # -------------------------------------------------

        if (
            "Student_ID" not in df.columns
            or
            "Name" not in df.columns
        ):

            return jsonify({

                "success": False,

                "message":
                    "Student_ID आणि Name columns required आहेत."

            }), 400

        # -------------------------------------------------
        # SAVE TO SELECTED SEMESTER
        # -------------------------------------------------

        success = save_df(
            df,
            semester
        )

        if not success:

            return jsonify({

                "success": False,

                "message":
                    "Excel save करण्यात problem आला."

            }), 500

        return jsonify({

            "success": True,

            "message":
                f"{semester} Excel successfully uploaded and saved."

        })

    except Exception as e:

        print(
            "Upload error:",
            e
        )

        return jsonify({

            "success": False,

            "message":
                f"Upload error: {str(e)}"

        }), 500


# =========================================================
# ADD STUDENT
# =========================================================

@app.route(
    "/api/add_student",
    methods=["POST"]
)
def add_student():

    check = login_required_api()

    if check:
        return check

    try:

        data = request.get_json() or {}

        semester = data.get(
            "semester",
            "Semester 1"
        )

        if not valid_semester(semester):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester."

            }), 400

        df = read_df(
            semester
        )

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

                "message":
                    "Student ID required आहे."

            })

        if not name:

            return jsonify({

                "success": False,

                "message":
                    "Student Name required आहे."

            })

        # -------------------------------------------------
        # DUPLICATE ID CHECK
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
        # CREATE ROW
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
        # ADD ROW
        # -------------------------------------------------

        df = pd.concat(

            [
                df,
                pd.DataFrame([row])
            ],

            ignore_index=True

        )

        success = save_df(
            df,
            semester
        )

        if not success:

            return jsonify({

                "success": False,

                "message":
                    "Student save करता आला नाही."

            }), 500

        return jsonify({

            "success": True,

            "message":
                f"Student added successfully in {semester}."

        })

    except Exception as e:

        print(
            "Add student error:",
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

    check = login_required_api()

    if check:
        return check

    try:

        data = request.get_json() or {}

        semester = data.get(
            "semester",
            "Semester 1"
        )

        student_id = str(
            data.get(
                "Student_ID",
                ""
            )
        ).strip()

        if not valid_semester(semester):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester."

            }), 400

        if not student_id:

            return jsonify({

                "success": False,

                "message":
                    "Student ID required आहे."

            })

        df = read_df(
            semester
        )

        if df.empty:

            return jsonify({

                "success": False,

                "message":
                    "Student not found."

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
            df["Student_ID"]
            == student_id
        ].tolist()

        if not matches:

            return jsonify({

                "success": False,

                "message":
                    "Student not found."

            })

        index = matches[0]

        editable = [
            "Name",
            "Gender",
            "Class",
            "Attendance"
        ]

        editable += SUBJECTS[semester]

        for column in editable:

            if column in data:

                df.loc[
                    index,
                    column
                ] = data[column]

        success = save_df(
            df,
            semester
        )

        if not success:

            return jsonify({

                "success": False,

                "message":
                    "Student update करता आला नाही."

            }), 500

        return jsonify({

            "success": True,

            "message":
                "Student updated successfully."

        })

    except Exception as e:

        print(
            "Edit student error:",
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

    check = login_required_api()

    if check:
        return check

    try:

        data = request.get_json() or {}

        semester = data.get(
            "semester",
            "Semester 1"
        )

        student_id = str(
            data.get(
                "Student_ID",
                ""
            )
        ).strip()

        if not valid_semester(semester):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester."

            }), 400

        df = read_df(
            semester
        )

        if df.empty:

            return jsonify({

                "success": False,

                "message":
                    "Student not found."

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
            df["Student_ID"]
            != student_id
        ]

        if len(new_df) == len(df):

            return jsonify({

                "success": False,

                "message":
                    "Student not found."

            })

        success = save_df(
            new_df,
            semester
        )

        if not success:

            return jsonify({

                "success": False,

                "message":
                    "Student delete करता आला नाही."

            }), 500

        return jsonify({

            "success": True,

            "message":
                "Student deleted successfully."

        })

    except Exception as e:

        print(
            "Delete student error:",
            e
        )

        return jsonify({

            "success": False,

            "message":
                str(e)

        }), 500


# =========================================================
# HEALTH CHECK
# =========================================================

@app.route("/health")
def health():

    return jsonify({

        "status": "OK",

        "application":
            "Student Performance Analysis System",

        "branch":
            "Computer Engineering",

        "semesters":
            list(SEMESTER_FILES.keys())

    })


# =========================================================
# STARTUP
# =========================================================

create_all_files()


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
