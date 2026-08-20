from flask import Flask, render_template, jsonify, request, redirect, url_for, session
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = os.environ.get(
    "SECRET_KEY",
    "student-performance-secret-key"
)

BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))

SEMESTER_YEAR = {
    "Semester 1": "1st Year",
    "Semester 2": "1st Year",
    "Semester 3": "2nd Year",
    "Semester 4": "2nd Year",
    "Semester 5": "3rd Year",
    "Semester 6": "3rd Year"
}

YEAR_SEMESTERS = {
    "1st Year": ["Semester 1", "Semester 2"],
    "2nd Year": ["Semester 3", "Semester 4"],
    "3rd Year": ["Semester 5", "Semester 6"]
}

SEMESTER_FILES = {
    "Semester 1": "semester_1.xlsx",
    "Semester 2": "semester_2.xlsx",
    "Semester 3": "semester_3.xlsx",
    "Semester 4": "semester_4.xlsx",
    "Semester 5": "semester_5.xlsx",
    "Semester 6": "semester_6.xlsx"
}

SUBJECTS = {
    "Semester 1": [
        "Basic Mathematics",
        "Communication Skills",
        "Basic Science"
    ],

    "Semester 2": [
        "Applied Mathematics",
        "Basic Electrical and Engineering",
        "Programming in C"
    ],

    "Semester 3": [
        "Object Oriented Programming",
        "Data Structure",
        "Digital Techniques",
        "Database Management System"
    ],

    "Semester 4": [
        "Java Programming",
        "Data Communication and Network",
        "Microprocessor Programming",
        "Environmental Education And Sustanability"
    ],

    "Semester 5": [
        "Software Engineering",
        "Opreting System",
        "Data Analytics"
    ],

    "Semester 6": [
        "Mobile Application Development",
        "Machine Learning",
        "Software Testing",
        "Mangement"
    ]
}

LOGIN_USERNAME = os.environ.get(
    "ADMIN_USERNAME",
    "silicon"
)

LOGIN_PASSWORD = os.environ.get(
    "ADMIN_PASSWORD",
    "patil"
)


# =========================================================
# BASIC FUNCTIONS
# =========================================================

def valid_semester(semester):
    return semester in SEMESTER_FILES


def valid_year(year):
    return year in YEAR_SEMESTERS


def excel_path(semester):
    return os.path.join(
        BASE_FOLDER,
        SEMESTER_FILES[semester]
    )


def logged_in():
    return session.get("logged_in") is True


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
# NUMBER CLEANING
# =========================================================

def clean_number(value, default=0):

    if value is None:
        return default

    if isinstance(value, str):

        value = value.strip()

        if value == "":
            return default

        value = value.replace("%", "")

    try:

        number = float(value)

        if pd.isna(number):
            return default

        return round(number, 2)

    except (ValueError, TypeError):

        return default


# =========================================================
# READ EXCEL
# =========================================================

def read_excel(semester):

    if not valid_semester(semester):
        return pd.DataFrame()

    path = excel_path(semester)

    if not os.path.exists(path):
        return pd.DataFrame()

    try:

        df = pd.read_excel(path)

        if df is None:
            return pd.DataFrame()

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        return df

    except Exception as e:

        print("Excel read error:", repr(e))

        return pd.DataFrame()


# =========================================================
# PROCESS DATA
# =========================================================

def process_data(df, semester):

    df = df.copy()

    subjects = SUBJECTS.get(
        semester,
        []
    )

    required_columns = [
        "Student_ID",
        "Name",
        "Gender",
        "Class",
        "Attendance"
    ]

    for column in required_columns:

        if column not in df.columns:

            if column == "Attendance":
                df[column] = 0.0

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
    # TEXT COLUMNS
    # -----------------------------------------------------

    for column in [
        "Name",
        "Gender",
        "Class"
    ]:

        df[column] = (
            df[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )


    # -----------------------------------------------------
    # ATTENDANCE
    # -----------------------------------------------------

    raw_attendance = (
        df["Attendance"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    percent_mask = raw_attendance.str.endswith("%")

    df["Attendance"] = pd.to_numeric(
        raw_attendance.str.replace(
            "%",
            "",
            regex=False
        ),
        errors="coerce"
    )

    fraction_mask = (
        df["Attendance"].notna()
        & df["Attendance"].between(0, 1)
        & ~percent_mask
    )

    df.loc[
        fraction_mask,
        "Attendance"
    ] = (
        df.loc[
            fraction_mask,
            "Attendance"
        ] * 100
    )

    df["Attendance"] = (
        df["Attendance"]
        .fillna(0)
        .clip(0, 100)
        .round(2)
        .astype(float)
    )


    # -----------------------------------------------------
    # SUBJECTS
    # -----------------------------------------------------

    for subject in subjects:

        if subject not in df.columns:
            df[subject] = 0.0

        df[subject] = (
            pd.to_numeric(
                df[subject],
                errors="coerce"
            )
            .fillna(0)
            .clip(0, 100)
            .round(2)
            .astype(float)
        )


    # -----------------------------------------------------
    # TOTAL
    # -----------------------------------------------------

    if subjects:

        df["Total"] = (
            df[subjects]
            .sum(axis=1)
            .round(2)
        )

        df["Percentage"] = (
            df["Total"]
            / (len(subjects) * 100)
            * 100
        ).round(2)

    else:

        df["Total"] = 0.0
        df["Percentage"] = 0.0


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

def save_excel(df, semester):

    processed = process_data(
        df,
        semester
    )

    columns = (
        [
            "Student_ID",
            "Name",
            "Gender",
            "Class"
        ]
        + SUBJECTS[semester]
        + [
            "Attendance"
        ]
    )


    for column in columns:

        if column not in processed.columns:

            if column == "Attendance":
                processed[column] = 0.0

            else:
                processed[column] = ""


    # IMPORTANT FIX
    # All marks and attendance become float
    # before writing to Excel.

    numeric_columns = (
        SUBJECTS[semester]
        + ["Attendance"]
    )

    for column in numeric_columns:

        processed[column] = (
            pd.to_numeric(
                processed[column],
                errors="coerce"
            )
            .fillna(0)
            .clip(0, 100)
            .round(2)
        )


    processed[columns].to_excel(
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
        return redirect(
            url_for("home")
        )


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


    semester = (
        request.args
        .get(
            "semester",
            "Semester 1"
        )
        .strip()
    )


    if not valid_semester(semester):

        return jsonify({
            "success": False,
            "message": "Invalid semester"
        }), 400


    return jsonify({

        "success": True,

        "branch":
            "Computer Engineering",

        "year":
            SEMESTER_YEAR[semester],

        "semester":
            semester,

        "subjects": [
            {
                "code": subject,
                "name": subject
            }

            for subject in SUBJECTS[
                semester
            ]
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


    semester = (
        request.args
        .get(
            "semester",
            "Semester 1"
        )
        .strip()
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


    columns = (
        [
            "Student_ID",
            "Name",
            "Gender",
            "Class"
        ]
        + SUBJECTS[semester]
        + [
            "Total",
            "Percentage",
            "Attendance",
            "Attendance Status",
            "Grade"
        ]
    )


    return jsonify(
        df[columns]
        .fillna("")
        .to_dict("records")
    )


# =========================================================
# ANALYTICS
# =========================================================

@app.route("/api/analytics")
def api_analytics():

    check = api_login_check()

    if check:
        return check


    semester = (
        request.args
        .get(
            "semester",
            "Semester 1"
        )
        .strip()
    )


    empty = {

        "total_students": 0,

        "average_percentage": 0,

        "top_performer": "-",

        "average_attendance": 0,

        "subjects": {},

        "grades": {}
    }


    if not valid_semester(semester):
        return jsonify(empty)


    df = read_excel(semester)

    if df.empty:
        return jsonify(empty)


    df = process_data(
        df,
        semester
    )


    if df.empty:
        return jsonify(empty)


    top = "-"


    try:

        top_index = (
            df["Percentage"]
            .idxmax()
        )

        top = str(
            df.loc[
                top_index,
                "Name"
            ]
        )

    except Exception:
        pass


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
            top,

        "average_attendance":
            round(
                float(
                    df["Attendance"].mean()
                ),
                2
            ),

        "subjects": {

            subject:
                round(
                    float(
                        df[subject].mean()
                    ),
                    2
                )

            for subject in SUBJECTS[
                semester
            ]
        },

        "grades":
            df["Grade"]
            .value_counts()
            .to_dict()
    })


# =========================================================
# COMBINED ANALYTICS
# =========================================================

@app.route(
    "/api/combined_analytics"
)
def combined_analytics():

    check = api_login_check()

    if check:
        return check


    year = (
        request.args
        .get(
            "year",
            "1st Year"
        )
        .strip()
    )


    if not valid_year(year):

        return jsonify({

            "success": False,

            "message":
                "Invalid year"

        }), 400


    result = {}


    for semester in YEAR_SEMESTERS[year]:

        df = read_excel(
            semester
        )


        result[semester] = {

            "subjects": {},

            "grades": {},

            "students": 0,

            "percentage": 0,

            "attendance": 0
        }


        if df.empty:
            continue


        df = process_data(
            df,
            semester
        )


        if df.empty:
            continue


        result[semester][
            "students"
        ] = int(len(df))


        result[semester][
            "percentage"
        ] = round(
            float(
                df["Percentage"].mean()
            ),
            2
        )


        result[semester][
            "attendance"
        ] = round(
            float(
                df["Attendance"].mean()
            ),
            2
        )


        result[semester][
            "subjects"
        ] = {

            subject:
                round(
                    float(
                        df[subject].mean()
                    ),
                    2
                )

            for subject in SUBJECTS[
                semester
            ]
        }


        result[semester][
            "grades"
        ] = (
            df["Grade"]
            .value_counts()
            .to_dict()
        )


    return jsonify({

        "success": True,

        "year": year,

        "semesters": result

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

        semester = (
            request.form
            .get(
                "semester",
                ""
            )
            .strip()
        )

        file = request.files.get(
            "file"
        )


        if not valid_semester(
            semester
        ):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester selected."

            }), 400


        if not file or not file.filename:

            return jsonify({

                "success": False,

                "message":
                    "Please select Excel file."

            }), 400


        if not file.filename.lower().endswith(
            (".xlsx", ".xls")
        ):

            return jsonify({

                "success": False,

                "message":
                    "Only Excel files are allowed."

            }), 400


        df = pd.read_excel(
            file
        )


        if df.empty:

            return jsonify({

                "success": False,

                "message":
                    "Uploaded Excel is empty."

            }), 400


        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )


        for column in [
            "Student_ID",
            "Name"
        ]:

            if column not in df.columns:

                return jsonify({

                    "success": False,

                    "message":
                        f"{column} column is required."

                }), 400


        save_excel(
            df,
            semester
        )


        return jsonify({

            "success": True,

            "message":
                f"{semester} Excel uploaded successfully."

        })


    except Exception as e:

        print(
            "Excel upload error:",
            repr(e)
        )

        return jsonify({

            "success": False,

            "message":
                f"Excel upload failed: {e}"

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


        if not valid_semester(
            semester
        ):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester."

            }), 400


        if not student_id or not name:

            return jsonify({

                "success": False,

                "message":
                    "Student ID and Name are required."

            })


        df = read_excel(
            semester
        )


        if (
            not df.empty
            and "Student_ID" in df.columns
        ):

            ids = (
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


            if student_id in ids.tolist():

                return jsonify({

                    "success": False,

                    "message":
                        "Student ID already exists."

                })


        row = {

            "Student_ID":
                student_id,

            "Name":
                name,

            "Gender":
                str(
                    data.get(
                        "Gender",
                        ""
                    )
                ).strip(),

            "Class":
                str(
                    data.get(
                        "Class",
                        ""
                    )
                ).strip(),

            "Attendance":
                clean_number(
                    data.get(
                        "Attendance",
                        0
                    )
                )
        }


        for subject in SUBJECTS[
            semester
        ]:

            row[subject] = clean_number(
                data.get(
                    subject,
                    0
                )
            )


        df = pd.concat(
            [
                df,
                pd.DataFrame([row])
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


    except Exception as e:

        print(
            "Add error:",
            repr(e)
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

    check = api_login_check()

    if check:
        return check


    try:

        data = request.get_json() or {}


        semester = str(
            data.get(
                "semester",
                ""
            )
        ).strip()


        student_id = str(
            data.get(
                "Student_ID",
                ""
            )
        ).strip()


        if not valid_semester(
            semester
        ):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester."

            }), 400


        if not student_id:

            return jsonify({

                "success": False,

                "message":
                    "Student ID is required."

            }), 400


        df = read_excel(
            semester
        )


        if (
            df.empty
            or "Student_ID"
            not in df.columns
        ):

            return jsonify({

                "success": False,

                "message":
                    "Student not found."

            })


        # IMPORTANT:
        # Excel Student ID like 2301.0
        # becomes 2301

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


        # -------------------------------------------------
        # TEXT VALUES
        # -------------------------------------------------

        for column in [
            "Name",
            "Gender",
            "Class"
        ]:

            if column in data:

                df.at[
                    index,
                    column
                ] = str(
                    data[column]
                ).strip()


        # -------------------------------------------------
        # NUMERIC VALUES
        # -------------------------------------------------
        # THIS FIXES:
        # Invalid value '92' for dtype int64
        #
        # Every edited mark/attendance is converted
        # to numeric float before saving.
        # -------------------------------------------------

        for column in (
            SUBJECTS[semester]
            + ["Attendance"]
        ):

            if column in data:

                value = clean_number(
                    data[column],
                    0
                )

                df.at[
                    index,
                    column
                ] = float(value)


        # -------------------------------------------------
        # SAVE AND RECALCULATE
        # -------------------------------------------------

        save_excel(
            df,
            semester
        )


        return jsonify({

            "success": True,

            "message":
                "Student updated successfully."

        })


    except Exception as e:

        print(
            "Edit error:",
            repr(e)
        )

        return jsonify({

            "success": False,

            "message":
                f"Update failed: {e}"

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
                ""
            )
        ).strip()


        student_id = str(
            data.get(
                "Student_ID",
                ""
            )
        ).strip()


        if not valid_semester(
            semester
        ):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester."

            }), 400


        df = read_excel(
            semester
        )


        if (
            df.empty
            or "Student_ID"
            not in df.columns
        ):

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
        ].copy()


        if len(new_df) == len(df):

            return jsonify({

                "success": False,

                "message":
                    "Student not found."

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


    except Exception as e:

        print(
            "Delete error:",
            repr(e)
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

        "status": "ok",

        "application":
            "Student Performance Analysis System"

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
