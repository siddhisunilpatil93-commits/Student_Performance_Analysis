from flask import Flask, render_template, jsonify, request, redirect, url_for, session
import pandas as pd
import os

app = Flask(__name__)

# =========================================================
# SECRET KEY
# =========================================================

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "student-performance-change-this-secret"
)

BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))

DATA_FOLDER = os.path.join(
    BASE_FOLDER,
    "data"
)

os.makedirs(DATA_FOLDER, exist_ok=True)


# =========================================================
# YEAR + SEMESTER
# =========================================================

SEMESTERS = {

    "1st Year": [
        "Semester 1",
        "Semester 2"
    ],

    "2nd Year": [
        "Semester 3",
        "Semester 4"
    ],

    "3rd Year": [
        "Semester 5",
        "Semester 6"
    ]

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
# HELPER
# =========================================================

def valid_semester(semester):

    return semester in SEMESTER_FILES


def is_logged_in():

    return session.get("logged_in") is True


def login_required_page():

    if not is_logged_in():

        return redirect(
            url_for("login")
        )

    return None


def login_required_api():

    if not is_logged_in():

        return jsonify({

            "success": False,

            "message":
                "Login required."

        }), 401

    return None


# =========================================================
# EXCEL PATH
# =========================================================

def path_for(semester):

    filename = SEMESTER_FILES[semester]

    # First check data folder
    data_path = os.path.join(
        DATA_FOLDER,
        filename
    )

    if os.path.exists(data_path):

        return data_path

    # Then check project root
    root_path = os.path.join(
        BASE_FOLDER,
        filename
    )

    if os.path.exists(root_path):

        return root_path

    # New files will be created inside data folder
    return data_path


# =========================================================
# GRADE
# =========================================================

def grade(percentage):

    try:

        percentage = float(percentage)

    except Exception:

        return "F"

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
# READ EXCEL
# =========================================================

def read_df(semester):

    if not valid_semester(semester):

        return pd.DataFrame()

    file_path = path_for(semester)

    if not os.path.exists(file_path):

        return pd.DataFrame()

    try:

        df = pd.read_excel(
            file_path
        )

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

    basic_columns = [

        "Student_ID",
        "Name",
        "Gender",
        "Class",
        "Attendance"

    ]

    # Add missing basic columns

    for column in basic_columns:

        if column not in df.columns:

            if column == "Attendance":

                df[column] = 0

            else:

                df[column] = ""


    # Student ID

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


    # Name

    df["Name"] = (

        df["Name"]
        .fillna("")
        .astype(str)
        .str.strip()

    )


    # Gender

    df["Gender"] = (

        df["Gender"]
        .fillna("")
        .astype(str)
        .str.strip()

    )


    # Class

    df["Class"] = (

        df["Class"]
        .fillna("")
        .astype(str)
        .str.strip()

    )


    # Attendance

    df["Attendance"] = (

        pd.to_numeric(
            df["Attendance"],
            errors="coerce"
        )
        .fillna(0)
        .clip(0, 100)
        .round(2)

    )


    # Subjects

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


    # Total

    df["Total"] = (

        df[subjects]
        .sum(axis=1)
        .round(2)

    )


    # Percentage

    df["Percentage"] = (

        df["Total"]
        / (len(subjects) * 100)
        * 100

    ).round(2)


    # Attendance Status

    df["Attendance Status"] = (

        df["Attendance"]
        .apply(
            lambda x:
            "Good"
            if float(x) >= 75
            else "Bad"
        )

    )


    # Grade

    df["Grade"] = (

        df["Percentage"]
        .apply(grade)

    )


    return df


# =========================================================
# SAVE EXCEL
# =========================================================

def save_df(df, semester):

    processed = process(
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

            processed[column] = ""


    processed[columns].to_excel(

        path_for(semester),

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

    if is_logged_in():

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
            and
            password == LOGIN_PASSWORD
        ):

            session["logged_in"] = True

            return redirect(
                url_for("home")
            )


        return render_template(

            "index.html",

            login_page=True,

            login_error=
                "Invalid username or password."

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
    ).strip()


    if not valid_semester(semester):

        return jsonify({

            "success": False,

            "message":
                "Invalid semester."

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

            for subject
            in SUBJECTS[semester]

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
    ).strip()


    if not valid_semester(semester):

        return jsonify([])


    df = read_df(
        semester
    )


    if df.empty:

        return jsonify([])


    df = process(
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
    ).strip()


    if not valid_semester(semester):

        return jsonify({

            "total_students": 0,
            "average_percentage": 0,
            "top_performer": "-",
            "average_attendance": 0,
            "subjects": {},
            "grades": {}

        })


    df = read_df(
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


    df = process(
        df,
        semester
    )


    top = "-"


    if not df.empty:

        top_index =
            df["Percentage"].idxmax()

        top = str(
            df.loc[
                top_index,
                "Name"
            ]
        )


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

            for subject
            in SUBJECTS[semester]

        },

        "grades":
            df["Grade"]
            .value_counts()
            .to_dict()

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
        ).strip()


        file = request.files.get(
            "file"
        )


        # Check semester

        if not valid_semester(semester):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester."

            })


        # Check file

        if not file:

            return jsonify({

                "success": False,

                "message":
                    "Excel file select करा."

            })


        filename =
            file.filename.lower().strip()


        if not (
            filename.endswith(".xlsx")
            or
            filename.endswith(".xls")
        ):

            return jsonify({

                "success": False,

                "message":
                    "Only Excel file (.xlsx / .xls) allowed."

            })


        # Read uploaded Excel

        df = pd.read_excel(
            file
        )


        # Clean column names

        df.columns = (

            df.columns
            .astype(str)
            .str.strip()

        )


        # Required columns

        if (
            "Student_ID"
            not in df.columns
            or
            "Name"
            not in df.columns
        ):

            return jsonify({

                "success": False,

                "message":
                    "Excel मध्ये Student_ID आणि Name columns असणे आवश्यक आहे."

            })


        # Process and save

        save_df(
            df,
            semester
        )


        return jsonify({

            "success": True,

            "message":
                f"{semester} Excel successfully uploaded."

        })


    except Exception as e:

        print(
            "UPLOAD ERROR:",
            repr(e)
        )


        return jsonify({

            "success": False,

            "message":
                "Excel upload error: " + str(e)

        })


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


        semester = str(
            data.get(
                "semester",
                "Semester 1"
            )
        ).strip()


        if not valid_semester(semester):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester."

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


        if not student_id or not name:

            return jsonify({

                "success": False,

                "message":
                    "Student ID आणि Name required आहेत."

            })


        df = read_df(
            semester
        )


        # Duplicate check

        if not df.empty:

            df = process(
                df,
                semester
            )


            existing_ids = (

                df["Student_ID"]
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


        # Create row

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


        save_df(
            df,
            semester
        )


        return jsonify({

            "success": True,

            "message":
                f"Student added to {semester} Excel."

        })


    except Exception as e:

        print(
            "ADD ERROR:",
            repr(e)
        )


        return jsonify({

            "success": False,

            "message":
                "Add student error: " + str(e)

        })


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

                "message":
                    "Invalid semester."

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


        editable = (

            [
                "Name",
                "Gender",
                "Class",
                "Attendance"
            ]

            + SUBJECTS[semester]

        )


        for column in editable:

            if column in data:

                df.loc[
                    index,
                    column
                ] = data[column]


        save_df(
            df,
            semester
        )


        return jsonify({

            "success": True,

            "message":
                "Student updated in Excel."

        })


    except Exception as e:

        print(
            "EDIT ERROR:",
            repr(e)
        )


        return jsonify({

            "success": False,

            "message":
                "Edit student error: " + str(e)

        })


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

                "message":
                    "Invalid semester."

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


        save_df(
            new_df,
            semester
        )


        return jsonify({

            "success": True,

            "message":
                f"Student deleted from {semester} Excel."

        })


    except Exception as e:

        print(
            "DELETE ERROR:",
            repr(e)
        )


        return jsonify({

            "success": False,

            "message":
                "Delete student error: " + str(e)

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
