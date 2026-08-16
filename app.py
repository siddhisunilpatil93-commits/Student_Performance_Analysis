from flask import Flask, render_template, jsonify, request, redirect, url_for, session
import pandas as pd
import os

app = Flask(__name__)

# ============================================================
# LOGIN
# ============================================================

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "student-performance-secret-key"
)

LOGIN_USERNAME = os.environ.get("ADMIN_USERNAME", "silicon")
LOGIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "patil")


# ============================================================
# BASE FOLDER
# ============================================================

BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))


# ============================================================
# YEAR / SEMESTER
# ============================================================

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

SEMESTER_FILES = {
    "Semester 1": "semester_1.xlsx",
    "Semester 2": "semester_2.xlsx",
    "Semester 3": "semester_3.xlsx",
    "Semester 4": "semester_4.xlsx",
    "Semester 5": "semester_5.xlsx",
    "Semester 6": "semester_6.xlsx"
}


# ============================================================
# COMPUTER ENGINEERING SUBJECTS
# ============================================================

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


# ============================================================
# FILE PATH
# ============================================================

def excel_path(semester):

    return os.path.join(
        BASE_FOLDER,
        SEMESTER_FILES[semester]
    )


# ============================================================
# VALIDATE SEMESTER
# ============================================================

def valid_semester(semester):

    return semester in SEMESTER_FILES


# ============================================================
# GRADE
# ============================================================

def calculate_grade(percentage):

    try:
        percentage = float(percentage)
    except:
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

    return "F"


# ============================================================
# READ EXCEL
# ============================================================

def read_excel(semester):

    if not valid_semester(semester):
        return pd.DataFrame()

    file_path = excel_path(semester)

    if not os.path.exists(file_path):
        return pd.DataFrame()

    try:

        df = pd.read_excel(file_path)

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        return df

    except Exception as e:

        print("Excel read error:", e)

        return pd.DataFrame()


# ============================================================
# PROCESS DATA
# ============================================================

def process_data(df, semester):

    df = df.copy()

    # --------------------------------------------------------
    # Required columns
    # --------------------------------------------------------

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
                df[column] = 0
            else:
                df[column] = ""


    # --------------------------------------------------------
    # Student ID
    # --------------------------------------------------------

    df["Student_ID"] = (
        df["Student_ID"]
        .fillna("")
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
    )


    # --------------------------------------------------------
    # Name
    # --------------------------------------------------------

    df["Name"] = (
        df["Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    # --------------------------------------------------------
    # Attendance
    # --------------------------------------------------------

    df["Attendance"] = (
        pd.to_numeric(
            df["Attendance"],
            errors="coerce"
        )
        .fillna(0)
        .clip(0, 100)
        .round(2)
    )


    # --------------------------------------------------------
    # Subjects
    # --------------------------------------------------------

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


    # --------------------------------------------------------
    # Total
    # --------------------------------------------------------

    df["Total"] = (
        df[subjects]
        .sum(axis=1)
        .round(2)
    )


    # --------------------------------------------------------
    # Percentage
    # --------------------------------------------------------

    df["Percentage"] = (
        df["Total"]
        / (len(subjects) * 100)
        * 100
    ).round(2)


    # --------------------------------------------------------
    # Attendance Status
    # --------------------------------------------------------

    df["Attendance Status"] = df["Attendance"].apply(
        lambda x: "Good" if x >= 75 else "Bad"
    )


    # --------------------------------------------------------
    # Grade
    # --------------------------------------------------------

    df["Grade"] = df["Percentage"].apply(
        calculate_grade
    )


    return df


# ============================================================
# SAVE EXCEL
# ============================================================

def save_excel(df, semester):

    if not valid_semester(semester):
        return False

    try:

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
            + ["Attendance"]
        )

        for column in columns:

            if column not in processed.columns:
                processed[column] = ""

        final_df = processed[columns].copy()

        final_df.to_excel(
            excel_path(semester),
            index=False
        )

        print(
            f"Excel updated successfully: "
            f"{excel_path(semester)}"
        )

        return True

    except Exception as e:

        print("Excel save error:", e)

        return False


# ============================================================
# LOGIN CHECK
# ============================================================

def logged_in():

    return session.get("logged_in") is True


# ============================================================
# LOGIN PAGE
# ============================================================

@app.route(
    "/login",
    methods=["GET", "POST"]
)
def login():

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


# ============================================================
# LOGOUT
# ============================================================

@app.route("/logout")
def logout():

    session.clear()

    return redirect(
        url_for("login")
    )


# ============================================================
# PROTECT ALL PAGES
# ============================================================

@app.before_request
def protect_pages():

    allowed = {
        "login",
        "static"
    }

    if request.endpoint in allowed:
        return None

    if not logged_in():

        if request.path.startswith("/api/"):

            return jsonify({
                "success": False,
                "message": "Login required."
            }), 401

        return redirect(
            url_for("login")
        )

    return None


# ============================================================
# HOME
# ============================================================

@app.route("/")
def home():

    return render_template(
        "index.html",
        login_page=False
    )


# ============================================================
# SUBJECT API
# ============================================================

@app.route("/api/subjects")
def api_subjects():

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
            for subject in SUBJECTS[semester]
        ]

    })


# ============================================================
# STUDENTS API
# ============================================================

@app.route("/api/students")
def api_students():

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

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


# ============================================================
# ANALYTICS API
# ============================================================

@app.route("/api/analytics")
def api_analytics():

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

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

    top_performer = "-"

    if len(df) > 0:

        index = df[
            "Percentage"
        ].idxmax()

        top_performer = str(
            df.loc[index, "Name"]
        )


    return jsonify({

        "total_students":
            int(len(df)),

        "average_percentage":
            round(
                df["Percentage"].mean(),
                2
            ),

        "top_performer":
            top_performer,

        "average_attendance":
            round(
                df["Attendance"].mean(),
                2
            ),

        "subjects": {

            subject:
                round(
                    df[subject].mean(),
                    2
                )

            for subject
            in SUBJECTS[semester]
        },

        "grades":
            df["Grade"]
            .value_counts()
            .to_dict()

    })


# ============================================================
# UPLOAD EXCEL
# ============================================================

@app.route(
    "/api/upload_excel",
    methods=["POST"]
)
def upload_excel():

    try:

        semester = request.form.get(
            "semester",
            "Semester 1"
        )

        file = request.files.get(
            "file"
        )

        if not valid_semester(semester):

            return jsonify({
                "success": False,
                "message": "Invalid semester."
            })


        if not file:

            return jsonify({
                "success": False,
                "message": "Please select Excel file."
            })


        filename = (
            file.filename
            .lower()
        )

        if not filename.endswith(
            (".xlsx", ".xls")
        ):

            return jsonify({
                "success": False,
                "message":
                    "Only Excel files allowed."
            })


        df = pd.read_excel(file)

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )


        if (
            "Student_ID" not in df.columns
            or
            "Name" not in df.columns
        ):

            return jsonify({
                "success": False,
                "message":
                    "Student_ID and Name columns required."
            })


        if not save_excel(
            df,
            semester
        ):

            return jsonify({
                "success": False,
                "message":
                    "Excel could not be saved."
            })


        return jsonify({

            "success": True,

            "message":
                f"{semester} Excel uploaded and updated successfully."

        })


    except Exception as e:

        return jsonify({

            "success": False,

            "message": str(e)

        })


# ============================================================
# ADD STUDENT
# ============================================================

@app.route(
    "/api/add_student",
    methods=["POST"]
)
def add_student():

    try:

        data = request.get_json() or {}

        semester = data.get(
            "semester",
            "Semester 1"
        )

        if not valid_semester(
            semester
        ):

            return jsonify({
                "success": False,
                "message": "Invalid semester."
            })


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
                    "Student ID required."
            })


        if not name:

            return jsonify({
                "success": False,
                "message":
                    "Student Name required."
            })


        df = read_excel(
            semester
        )


        if not df.empty:

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

            if student_id in df[
                "Student_ID"
            ].tolist():

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


        for subject in SUBJECTS[
            semester
        ]:

            row[subject] = data.get(
                subject,
                0
            )


        if df.empty:

            df = pd.DataFrame(
                [row]
            )

        else:

            df = pd.concat(
                [
                    df,
                    pd.DataFrame([row])
                ],
                ignore_index=True
            )


        if not save_excel(
            df,
            semester
        ):

            return jsonify({
                "success": False,
                "message":
                    "Excel update failed."
            })


        return jsonify({

            "success": True,

            "message":
                "Student added successfully and Excel updated."

        })


    except Exception as e:

        return jsonify({

            "success": False,

            "message": str(e)

        })


# ============================================================
# EDIT STUDENT
# ============================================================

@app.route(
    "/api/edit_student",
    methods=["POST"]
)
def edit_student():

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


        df = read_excel(
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
        ] + SUBJECTS[semester]


        for column in editable:

            if column in data:

                df.loc[
                    index,
                    column
                ] = data[column]


        if not save_excel(
            df,
            semester
        ):

            return jsonify({
                "success": False,
                "message":
                    "Excel update failed."
            })


        return jsonify({

            "success": True,

            "message":
                "Student updated successfully and Excel updated."

        })


    except Exception as e:

        return jsonify({

            "success": False,

            "message": str(e)

        })


# ============================================================
# DELETE STUDENT
# ============================================================

@app.route(
    "/api/delete_student",
    methods=["POST"]
)
def delete_student():

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


        df = read_excel(
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


        original_length = len(df)


        df = df[
            df["Student_ID"]
            != student_id
        ]


        if len(df) == original_length:

            return jsonify({
                "success": False,
                "message":
                    "Student not found."
            })


        if not save_excel(
            df,
            semester
        ):

            return jsonify({
                "success": False,
                "message":
                    "Excel update failed."
            })


        return jsonify({

            "success": True,

            "message":
                "Student deleted successfully and Excel updated."

        })


    except Exception as e:

        return jsonify({

            "success": False,

            "message": str(e)

        })


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    app.run(

        host="0.0.0.0",

        port=int(
            os.environ.get(
                "PORT",
                5000
            )
        ),

        debug=True

    )
