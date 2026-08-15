from flask import Flask, render_template, jsonify, request
import pandas as pd
import os

app = Flask(__name__)

DATA_FOLDER = "data"
os.makedirs(DATA_FOLDER, exist_ok=True)

# ---------------------------------------------------------
# YEAR + SEMESTER
# ---------------------------------------------------------

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

# Default Computer Engineering subjects.
# If Excel contains these columns, their actual marks are used.
SUBJECTS = {
    "Semester 1": [
        ("BMS", "Basic Mathematics"),
        ("BEE", "Basic Electrical Engineering"),
        ("PCC", "Programming in C"),
        ("EG", "Engineering Graphics")
    ],
    "Semester 2": [
        ("AMS", "Applied Mathematics"),
        ("PWP", "Python Programming"),
        ("DMS", "Data Management System"),
        ("DTE", "Digital Techniques")
    ],
    "Semester 3": [
        ("DSU", "Data Structure"),
        ("DBMS", "Database Management System"),
        ("CG", "Computer Graphics"),
        ("DTE", "Digital Techniques")
    ],
    "Semester 4": [
        ("JPR", "Java Programming"),
        ("DCN", "Data Communication and Computer Network"),
        ("MIC", "Microprocessor"),
        ("DAD", "Database Management")
    ],
    "Semester 5": [
        ("AJP", "Advanced Java Programming"),
        ("OS", "Operating System"),
        ("ETI", "Emerging Trends in IT"),
        ("ML", "Machine Learning")
    ],
    "Semester 6": [
        ("WBP", "Web Based Application Development"),
        ("MAD", "Mobile Application Development"),
        ("CSS", "Cloud and Security"),
        ("PWP", "Project Work")
    ]
}


# ---------------------------------------------------------
# FILE NAME
# IMPORTANT: semester-wise files prevent marks mixing
# ---------------------------------------------------------

def semester_file(semester):
    names = {
        "Semester 1": "semester_1.xlsx",
        "Semester 2": "semester_2.xlsx",
        "Semester 3": "semester_3.xlsx",
        "Semester 4": "semester_4.xlsx",
        "Semester 5": "semester_5.xlsx",
        "Semester 6": "semester_6.xlsx"
    }

    return os.path.join(
        DATA_FOLDER,
        names[semester]
    )


# ---------------------------------------------------------
# GRADE
# ---------------------------------------------------------

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


# ---------------------------------------------------------
# ATTENDANCE STATUS
# ---------------------------------------------------------

def attendance_status(value):

    try:
        a = float(value)
    except:
        return "Bad"

    if a >= 75:
        return "Good"
    else:
        return "Bad"


# ---------------------------------------------------------
# READ EXCEL
# ---------------------------------------------------------

def read_excel(semester):

    file = semester_file(semester)

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


# ---------------------------------------------------------
# FIND SUBJECT COLUMNS
# ---------------------------------------------------------

def get_subject_columns(df, semester):

    default_subjects = [
        code for code, name
        in SUBJECTS.get(semester, [])
    ]

    found = []

    # First preference: semester's proper subject codes
    for code in default_subjects:

        if code in df.columns:
            found.append(code)

    # If Excel has different subject columns,
    # use those numeric columns instead of showing 0.
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


# ---------------------------------------------------------
# PROCESS DATA
# ---------------------------------------------------------

def process_data(df, semester):

    if df.empty:
        return df

    df = df.copy()

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # Required columns
    for col in [
        "Student_ID",
        "Name",
        "Gender",
        "Class",
        "Attendance"
    ]:

        if col not in df.columns:
            df[col] = ""

    # Student ID
    df["Student_ID"] = (
        df["Student_ID"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    # Attendance
    df["Attendance"] = pd.to_numeric(
        df["Attendance"],
        errors="coerce"
    ).fillna(0)

    subjects = get_subject_columns(
        df,
        semester
    )

    # Convert subject marks
    for subject in subjects:

        df[subject] = pd.to_numeric(
            df[subject],
            errors="coerce"
        ).fillna(0)

    # Total
    if subjects:

        df["Total"] = df[subjects].sum(
            axis=1
        )

        max_marks = len(subjects) * 100

        df["Percentage"] = (
            df["Total"] /
            max_marks
        ) * 100

    else:

        df["Total"] = 0
        df["Percentage"] = 0

    df["Percentage"] = df[
        "Percentage"
    ].round(2)

    df["Grade"] = df[
        "Percentage"
    ].apply(get_grade)

    df["Attendance Status"] = df[
        "Attendance"
    ].apply(attendance_status)

    return df


# ---------------------------------------------------------
# SAVE
# ---------------------------------------------------------

def save_excel(df, semester):

    file = semester_file(semester)

    # Don't permanently store calculated columns
    save_df = df.copy()

    for col in [
        "Total",
        "Percentage",
        "Grade",
        "Attendance Status"
    ]:

        if col in save_df.columns:
            save_df = save_df.drop(
                columns=[col]
            )

    save_df.to_excel(
        file,
        index=False
    )


# ---------------------------------------------------------
# HOME
# ---------------------------------------------------------

@app.route("/")
def home():

    response = render_template(
        "index.html"
    )

    return response


# ---------------------------------------------------------
# NO CACHE
# ---------------------------------------------------------

@app.after_request
def no_cache(response):

    response.headers[
        "Cache-Control"
    ] = "no-store, no-cache, must-revalidate, max-age=0"

    response.headers[
        "Pragma"
    ] = "no-cache"

    response.headers[
        "Expires"
    ] = "0"

    return response


# ---------------------------------------------------------
# SUBJECT API
# ---------------------------------------------------------

@app.route("/api/subjects")
def subject_api():

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

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
        "branch": "Computer Engineering",
        "semester": semester,
        "year": SEMESTER_YEAR.get(
            semester,
            "1st Year"
        ),
        "subjects": subjects
    })


# ---------------------------------------------------------
# STUDENTS API
# ---------------------------------------------------------

@app.route("/api/students")
def students_api():

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
        c for c in columns
        if c in df.columns
    ]

    return jsonify(
        df[columns].fillna("").to_dict(
            orient="records"
        )
    )


# ---------------------------------------------------------
# ANALYTICS
# ---------------------------------------------------------

@app.route("/api/analytics")
def analytics():

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

    subjects = get_subject_columns(
        df,
        semester
    )

    top_name = "-"

    if len(df) > 0:

        index = df[
            "Percentage"
        ].idxmax()

        top_name = str(
            df.loc[index, "Name"]
        )

    subject_average = {}

    for subject in subjects:

        subject_average[
            subject
        ] = round(
            float(
                df[subject].mean()
            ),
            2
        )

    grades = (
        df["Grade"]
        .value_counts()
        .to_dict()
    )

    return jsonify({

        "total_students":
            len(df),

        "average_percentage":
            round(
                float(
                    df["Percentage"].mean()
                ),
                2
            ),

        "top_performer":
            top_name,

        "average_attendance":
            round(
                float(
                    df["Attendance"].mean()
                ),
                2
            ),

        "subjects":
            subject_average,

        "grades":
            grades
    })


# ---------------------------------------------------------
# EXCEL UPLOAD
# ---------------------------------------------------------

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

        if "file" not in request.files:

            return jsonify({
                "success": False,
                "message":
                    "Excel file select करा."
            })

        file = request.files["file"]

        if not file.filename:

            return jsonify({
                "success": False,
                "message":
                    "No file selected."
            })

        filename = file.filename.lower()

        if not filename.endswith(
            (".xlsx", ".xls")
        ):

            return jsonify({
                "success": False,
                "message":
                    "फक्त Excel .xlsx किंवा .xls file upload करा."
            })

        df = pd.read_excel(file)

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        # Required fields
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

        # Save exactly to selected semester
        save_excel(
            df,
            semester
        )

        return jsonify({
            "success": True,
            "message":
                f"{semester} Excel successfully uploaded."
        })

    except Exception as e:

        return jsonify({
            "success": False,
            "message": str(e)
        })


# ---------------------------------------------------------
# ADD STUDENT
# ---------------------------------------------------------

@app.route(
    "/api/add_student",
    methods=["POST"]
)
def add_student():

    try:

        data = request.get_json()

        semester = data.get(
            "semester",
            "Semester 1"
        )

        df = read_excel(
            semester
        )

        subjects = get_subject_columns(
            df,
            semester
        )

        student = {

            "Student_ID":
                data.get(
                    "Student_ID",
                    ""
                ),

            "Name":
                data.get(
                    "Name",
                    ""
                ),

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

        # If no Excel exists yet,
        # use default semester subjects
        if not subjects:

            subjects = [
                code
                for code, name
                in SUBJECTS.get(
                    semester,
                    []
                )
            ]

        for subject in subjects:

            student[subject] = data.get(
                subject,
                0
            )

        new_df = pd.DataFrame(
            [student]
        )

        if not df.empty:

            # Remove calculated columns
            for col in [
                "Total",
                "Percentage",
                "Grade",
                "Attendance Status"
            ]:

                if col in df.columns:
                    df = df.drop(
                        columns=[col]
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

        return jsonify({
            "success": False,
            "message": str(e)
        })


# ---------------------------------------------------------
# RUN
# ---------------------------------------------------------

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=int(
            os.environ.get(
                "PORT",
                5000
            )
        ),
        debug=False
    )
