from flask import Flask, render_template, jsonify, request
import pandas as pd
import os

app = Flask(__name__)

# =========================================================
# BASE FOLDER
# =========================================================

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
# SEMESTER FILES
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
# COMPUTER ENGINEERING SUBJECTS
# =========================================================

SUBJECTS = {

    "Semester 1": [
        ("Basic Mathematics", "Basic Mathematics"),
        ("Communication Skills", "Communication Skills"),
        ("Engineering Physics", "Engineering Physics"),
        ("Engineering Chemistry", "Engineering Chemistry"),
        ("Basic Science", "Basic Science")
    ],

    "Semester 2": [
        ("Applied Mathematics", "Applied Mathematics"),
        ("Engineering Graphics", "Engineering Graphics"),
        ("Basic Electrical Engineering", "Basic Electrical Engineering"),
        ("Programming in C", "Programming in C"),
        ("Web Page Design", "Web Page Design")
    ],

    "Semester 3": [
        ("Object Oriented Programming", "Object Oriented Programming"),
        ("Data Structure", "Data Structure"),
        ("Database Management System", "Database Management System"),
        ("Computer Networks", "Computer Networks"),
        ("Operating System", "Operating System")
    ],

    "Semester 4": [
        ("Java Programming", "Java Programming"),
        ("Data Communication and Network", "Data Communication and Network"),
        ("Microprocessor", "Microprocessor"),
        ("Software Engineering", "Software Engineering"),
        ("Python Programming", "Python Programming")
    ],

    "Semester 5": [
        ("Advanced Java", "Advanced Java"),
        ("Web Based Application Development", "Web Based Application Development"),
        ("Software Testing", "Software Testing"),
        ("Computer Security", "Computer Security"),
        ("Project Management", "Project Management")
    ],

    "Semester 6": [
        ("Mobile Application Development", "Mobile Application Development"),
        ("Cloud Computing", "Cloud Computing"),
        ("Artificial Intelligence", "Artificial Intelligence"),
        ("Internet of Things", "Internet of Things"),
        ("Major Project", "Major Project")
    ]
}


# =========================================================
# GET SEMESTER FILE
# =========================================================

def semester_file(semester):

    filename = SEMESTER_FILES.get(semester)

    if not filename:
        return None

    return os.path.join(
        BASE_FOLDER,
        filename
    )


# =========================================================
# VALIDATE SEMESTER
# =========================================================

def valid_semester(semester):

    return semester in SEMESTER_FILES


# =========================================================
# GRADE
# =========================================================

def get_grade(percentage):

    try:
        p = float(percentage)
    except:
        return "F"

    if p >= 90:
        return "A+"

    elif p >= 80:
        return "A"

    elif p >= 70:
        return "B+"

    elif p >= 60:
        return "B"

    elif p >= 50:
        return "C"

    elif p >= 40:
        return "D"

    else:
        return "F"


# =========================================================
# ATTENDANCE STATUS
# =========================================================

def attendance_status(value):

    try:
        attendance = float(value)
    except:
        return "Bad"

    if attendance >= 75:
        return "Good"

    return "Bad"


# =========================================================
# READ EXCEL
# =========================================================

def read_excel(semester):

    if not valid_semester(semester):
        return pd.DataFrame()

    file = semester_file(semester)

    if not file:
        return pd.DataFrame()

    if not os.path.exists(file):
        return pd.DataFrame()

    try:

        df = pd.read_excel(file)

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        return df

    except Exception as e:

        print("Excel read error:", e)

        return pd.DataFrame()


# =========================================================
# GET SUBJECT COLUMNS
# =========================================================

def get_subject_columns(df, semester):

    predefined = [
        code
        for code, name in SUBJECTS.get(
            semester,
            []
        )
    ]

    if df.empty:
        return predefined

    found = []

    # First use the predefined semester subjects
    for subject in predefined:

        if subject in df.columns:
            found.append(subject)

    # If Excel contains different subject columns,
    # detect numeric columns automatically.
    if not found:

        ignored = {
            "Student_ID",
            "Name",
            "Gender",
            "Class",
            "Attendance",
            "Total",
            "Percentage",
            "Grade",
            "Attendance Status",
            "Semester",
            "Year"
        }

        for col in df.columns:

            if col in ignored:
                continue

            numeric = pd.to_numeric(
                df[col],
                errors="coerce"
            )

            if numeric.notna().sum() > 0:
                found.append(col)

    return found


# =========================================================
# PROCESS DATA
# =========================================================

def process_data(df, semester):

    if df.empty:
        return df

    df = df.copy()

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # -----------------------------------------------------
    # Required columns
    # -----------------------------------------------------

    required = [
        "Student_ID",
        "Name",
        "Gender",
        "Class",
        "Attendance"
    ]

    for column in required:

        if column not in df.columns:
            df[column] = ""

    # -----------------------------------------------------
    # Student ID
    # -----------------------------------------------------

    df["Student_ID"] = (
        df["Student_ID"]
        .astype(str)
        .str.replace(
            ".0",
            "",
            regex=False
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
    # Attendance
    # -----------------------------------------------------

    df["Attendance"] = pd.to_numeric(
        df["Attendance"],
        errors="coerce"
    ).fillna(0)

    df["Attendance"] = (
        df["Attendance"]
        .clip(0, 100)
        .round(2)
    )

    # -----------------------------------------------------
    # Subjects
    # -----------------------------------------------------

    subjects = get_subject_columns(
        df,
        semester
    )

    # If no subjects found, use semester subjects
    if not subjects:

        subjects = [
            code
            for code, name
            in SUBJECTS.get(
                semester,
                []
            )
        ]

    # Make missing subject columns
    # with 0 marks
    for subject in subjects:

        if subject not in df.columns:
            df[subject] = 0

        df[subject] = pd.to_numeric(
            df[subject],
            errors="coerce"
        ).fillna(0)

        df[subject] = (
            df[subject]
            .clip(0, 100)
        )

    # -----------------------------------------------------
    # TOTAL
    # -----------------------------------------------------

    if subjects:

        df["Total"] = (
            df[subjects]
            .sum(axis=1)
        )

        max_marks = (
            len(subjects) * 100
        )

        df["Percentage"] = (
            df["Total"] /
            max_marks
        ) * 100

    else:

        df["Total"] = 0
        df["Percentage"] = 0

    df["Total"] = (
        df["Total"]
        .round(2)
    )

    df["Percentage"] = (
        df["Percentage"]
        .round(2)
    )

    # -----------------------------------------------------
    # GRADE
    # -----------------------------------------------------

    df["Grade"] = (
        df["Percentage"]
        .apply(get_grade)
    )

    # -----------------------------------------------------
    # ATTENDANCE STATUS
    # -----------------------------------------------------

    df["Attendance Status"] = (
        df["Attendance"]
        .apply(attendance_status)
    )

    return df


# =========================================================
# SAVE EXCEL
# =========================================================

def save_excel(df, semester):

    file = semester_file(semester)

    if not file:
        return False

    try:

        save_df = df.copy()

        calculated = [
            "Total",
            "Percentage",
            "Grade",
            "Attendance Status"
        ]

        for column in calculated:

            if column in save_df.columns:

                save_df = save_df.drop(
                    columns=[column]
                )

        save_df.to_excel(
            file,
            index=False
        )

        return True

    except Exception as e:

        print("Excel save error:", e)

        return False


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():

    return render_template(
        "index.html"
    )


# =========================================================
# NO CACHE
# =========================================================

@app.after_request
def no_cache(response):

    response.headers["Cache-Control"] = (
        "no-store, no-cache, "
        "must-revalidate, max-age=0"
    )

    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"

    return response


# =========================================================
# SUBJECT API
# =========================================================

@app.route("/api/subjects")
def subject_api():

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    if not valid_semester(semester):

        return jsonify({
            "success": False,
            "message": "Invalid semester",
            "subjects": []
        }), 400

    subjects = [
        {
            "code": code,
            "name": name
        }

        for code, name
        in SUBJECTS.get(
            semester,
            []
        )
    ]

    return jsonify({

        "success": True,

        "branch":
            "Computer Engineering",

        "semester":
            semester,

        "year":
            SEMESTER_YEAR.get(
                semester
            ),

        "subjects":
            subjects
    })


# =========================================================
# STUDENTS API
# =========================================================

@app.route("/api/students")
def students_api():

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    if not valid_semester(semester):
        return jsonify([])

    df = read_excel(
        semester
    )

    if df.empty:
        return jsonify([])

    df = process_data(
        df,
        semester
    )

    subjects = get_subject_columns(
        df,
        semester
    )

    columns = [
        "Student_ID",
        "Name",
        "Gender",
        "Class"
    ]

    columns += subjects

    columns += [
        "Total",
        "Percentage",
        "Attendance",
        "Attendance Status",
        "Grade"
    ]

    columns = [
        column
        for column in columns
        if column in df.columns
    ]

    result = (
        df[columns]
        .fillna("")
        .to_dict(
            orient="records"
        )
    )

    return jsonify(result)


# =========================================================
# ANALYTICS API
# =========================================================

@app.route("/api/analytics")
def analytics():

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

    df = read_excel(
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

    df = process_data(
        df,
        semester
    )

    subjects = get_subject_columns(
        df,
        semester
    )

    # -----------------------------------------------------
    # TOP PERFORMER
    # -----------------------------------------------------

    top_name = "-"

    if not df.empty:

        index = df[
            "Percentage"
        ].idxmax()

        top_name = str(
            df.loc[
                index,
                "Name"
            ]
        )

    # -----------------------------------------------------
    # SUBJECT AVERAGE
    # -----------------------------------------------------

    subject_average = {}

    for subject in subjects:

        if subject in df.columns:

            subject_average[
                subject
            ] = round(
                float(
                    df[subject].mean()
                ),
                2
            )

    # -----------------------------------------------------
    # GRADES
    # -----------------------------------------------------

    grades = (
        df["Grade"]
        .value_counts()
        .to_dict()
    )

    return jsonify({

        "total_students":
            int(len(df)),

        "average_percentage":
            round(
                float(
                    df[
                        "Percentage"
                    ].mean()
                ),
                2
            ),

        "top_performer":
            top_name,

        "average_attendance":
            round(
                float(
                    df[
                        "Attendance"
                    ].mean()
                ),
                2
            ),

        "subjects":
            subject_average,

        "grades":
            grades
    })


# =========================================================
# EXCEL UPLOAD
# =========================================================

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

        if not valid_semester(semester):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester selected."

            })

        if "file" not in request.files:

            return jsonify({

                "success": False,

                "message":
                    "Please select an Excel file."

            })

        file = request.files["file"]

        if not file.filename:

            return jsonify({

                "success": False,

                "message":
                    "No file selected."

            })

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
                    "Only .xlsx or .xls files are allowed."

            })

        df = pd.read_excel(file)

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

                "message":
                    "Excel मध्ये Student_ID column पाहिजे."

            })

        if "Name" not in df.columns:

            return jsonify({

                "success": False,

                "message":
                    "Excel मध्ये Name column पाहिजे."

            })

        if "Gender" not in df.columns:
            df["Gender"] = ""

        if "Class" not in df.columns:
            df["Class"] = ""

        if "Attendance" not in df.columns:
            df["Attendance"] = 0

        # -------------------------------------------------
        # Remove calculated columns
        # -------------------------------------------------

        calculated = [
            "Total",
            "Percentage",
            "Grade",
            "Attendance Status"
        ]

        for column in calculated:

            if column in df.columns:

                df = df.drop(
                    columns=[column]
                )

        # -------------------------------------------------
        # Validate / process
        # -------------------------------------------------

        processed = process_data(
            df,
            semester
        )

        # Save only raw + required data
        save_df = processed.copy()

        for column in calculated:

            if column in save_df.columns:

                save_df = save_df.drop(
                    columns=[column]
                )

        file_path = semester_file(
            semester
        )

        save_df.to_excel(
            file_path,
            index=False
        )

        return jsonify({

            "success": True,

            "message":
                f"{semester} Excel successfully uploaded."

        })

    except Exception as e:

        print(
            "Upload error:",
            e
        )

        return jsonify({

            "success": False,

            "message":
                f"Upload failed: {str(e)}"

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

                "message":
                    "Invalid student data."

            })

        semester = data.get(
            "semester",
            "Semester 1"
        )

        if not valid_semester(semester):

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester."

            })

        df = read_excel(
            semester
        )

        # -------------------------------------------------
        # Use semester subjects
        # -------------------------------------------------

        subjects = [
            code
            for code, name
            in SUBJECTS.get(
                semester,
                []
            )
        ]

        # If existing Excel has
        # different numeric subjects,
        # use them.
        if not df.empty:

            existing_subjects = (
                get_subject_columns(
                    df,
                    semester
                )
            )

            if existing_subjects:
                subjects = existing_subjects

        # -------------------------------------------------
        # Student
        # -------------------------------------------------

        student = {

            "Student_ID":
                str(
                    data.get(
                        "Student_ID",
                        ""
                    )
                ).strip(),

            "Name":
                str(
                    data.get(
                        "Name",
                        ""
                    )
                ).strip(),

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
                data.get(
                    "Attendance",
                    0
                )
        }

        # -------------------------------------------------
        # Subject marks
        # -------------------------------------------------

        for subject in subjects:

            student[subject] = data.get(
                subject,
                0
            )

        new_df = pd.DataFrame(
            [student]
        )

        # -------------------------------------------------
        # Existing data
        # -------------------------------------------------

        if not df.empty:

            calculated = [
                "Total",
                "Percentage",
                "Grade",
                "Attendance Status"
            ]

            for column in calculated:

                if column in df.columns:

                    df = df.drop(
                        columns=[column]
                    )

            df = pd.concat(
                [
                    df,
                    new_df
                ],
                ignore_index=True
            )

        else:

            df = new_df

        # -------------------------------------------------
        # Save
        # -------------------------------------------------

        success = save_excel(
            df,
            semester
        )

        if not success:

            return jsonify({

                "success": False,

                "message":
                    "Student save failed."

            })

        return jsonify({

            "success": True,

            "message":
                "Student added successfully."

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
