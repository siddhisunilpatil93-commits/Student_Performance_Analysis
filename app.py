from flask import Flask, render_template, jsonify, request
import pandas as pd
import os

app = Flask(__name__)

DATA_FOLDER = "data"
os.makedirs(DATA_FOLDER, exist_ok=True)

YEAR_FILES = {
    "1st Year": "first_year.xlsx",
    "2nd Year": "second_year.xlsx",
    "3rd Year": "third_year.xlsx"
}

# =========================================================
# SEMESTER-WISE SUBJECTS
# =========================================================

SUBJECTS_BY_SEMESTER = {

    "Semester 1": [
        {"code": "BMS", "name": "Basic Mathematics"},
        {"code": "BEE", "name": "Basic Electrical Engineering"},
        {"code": "PCC", "name": "Programming in C"},
        {"code": "EG", "name": "Engineering Graphics"}
    ],

    "Semester 2": [
        {"code": "AMS", "name": "Applied Mathematics"},
        {"code": "PWP", "name": "Python Programming"},
        {"code": "DMS", "name": "Data Management System"},
        {"code": "DTE", "name": "Digital Techniques"}
    ],

    "Semester 3": [
        {"code": "CG", "name": "Computer Graphics"},
        {"code": "DBMS", "name": "Database Management System"},
        {"code": "DSU", "name": "Data Structure"},
        {"code": "DTE", "name": "Digital Techniques"}
    ],

    "Semester 4": [
        {"code": "JPR", "name": "Java Programming"},
        {"code": "DCN", "name": "Data Communication and Computer Network"},
        {"code": "MIC", "name": "Microprocessor"},
        {"code": "DAD", "name": "Database Management"}
    ],

    "Semester 5": [
        {"code": "AJP", "name": "Advanced Java Programming"},
        {"code": "OS", "name": "Operating System"},
        {"code": "ETI", "name": "Emerging Trends in IT"},
        {"code": "ML", "name": "Machine Learning"}
    ],

    "Semester 6": [
        {"code": "WBP", "name": "Web Based Application Development"},
        {"code": "MAD", "name": "Mobile Application Development"},
        {"code": "CSS", "name": "Cloud and Security"},
        {"code": "PWP", "name": "Project Work"}
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


def get_subject_codes(semester):
    return [x["code"] for x in SUBJECTS_BY_SEMESTER.get(semester, [])]


def get_year_file(year):
    return os.path.join(DATA_FOLDER, YEAR_FILES[year])


# =========================================================
# GRADE
# =========================================================

def calculate_grade(p):
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
    return "F"


# =========================================================
# LOAD EXCEL
# =========================================================

def load_students(year, semester):

    if year not in YEAR_FILES:
        return pd.DataFrame()

    filepath = get_year_file(year)

    if not os.path.exists(filepath):
        return pd.DataFrame()

    try:
        df = pd.read_excel(filepath)
        df.columns = df.columns.astype(str).str.strip()

        # Basic columns
        for col in ["Student_ID", "Name", "Gender", "Class", "Attendance"]:
            if col not in df.columns:
                df[col] = ""

        subjects = get_subject_codes(semester)

        # Missing subjects = 0
        for subject in subjects:
            if subject not in df.columns:
                df[subject] = 0

            df[subject] = pd.to_numeric(
                df[subject],
                errors="coerce"
            ).fillna(0)

        # Attendance
        df["Attendance"] = pd.to_numeric(
            df["Attendance"],
            errors="coerce"
        ).fillna(0)

        # Calculations
        df["Total"] = df[subjects].sum(axis=1)

        if subjects:
            df["Percentage"] = (
                df["Total"] / (len(subjects) * 100)
            ) * 100
        else:
            df["Percentage"] = 0

        df["Percentage"] = df["Percentage"].round(2)

        df["Attendance Status"] = df["Attendance"].apply(
            lambda x: "Eligible" if x >= 75 else "Shortage"
        )

        df["Grade"] = df["Percentage"].apply(calculate_grade)

        # Student ID cleanup
        df["Student_ID"] = (
            df["Student_ID"]
            .astype(str)
            .str.replace(".0", "", regex=False)
            .str.strip()
        )

        return df.fillna("")

    except Exception as e:
        print("EXCEL ERROR:", e)
        return pd.DataFrame()


# =========================================================
# SAVE EXCEL
# =========================================================

def save_students(year, df):

    try:
        filepath = get_year_file(year)

        # Remove calculated columns before saving
        calculated = [
            "Total",
            "Percentage",
            "Attendance Status",
            "Grade"
        ]

        for col in calculated:
            if col in df.columns:
                df = df.drop(columns=[col])

        df.to_excel(filepath, index=False)

        return True

    except Exception as e:
        print("SAVE ERROR:", e)
        return False


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():
    return render_template("index.html")


# =========================================================
# SUBJECTS API
# =========================================================

@app.route("/api/subjects")
def subjects():

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    return jsonify({
        "success": True,
        "semester": semester,
        "academic_year": "2026-27",
        "year": SEMESTER_YEAR.get(
            semester,
            "1st Year"
        ),
        "subjects": SUBJECTS_BY_SEMESTER.get(
            semester,
            []
        )
    })


# =========================================================
# STUDENTS API
# =========================================================

@app.route("/api/students")
def students_api():

    year = request.args.get(
        "year",
        "1st Year"
    )

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    df = load_students(
        year,
        semester
    )

    if df.empty:
        return jsonify([])

    subjects = get_subject_codes(semester)

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
        df[columns].to_dict(
            orient="records"
        )
    )


# =========================================================
# ANALYTICS
# =========================================================

@app.route("/api/analytics")
def analytics():

    year = request.args.get(
        "year",
        "1st Year"
    )

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    df = load_students(
        year,
        semester
    )

    if df.empty:
        return jsonify({
            "total_students": 0,
            "average_percentage": 0,
            "top_performer": "-",
            "average_attendance": 0,
            "grades": {},
            "subjects": {}
        })

    subjects = get_subject_codes(semester)

    top = "-"

    if "Percentage" in df.columns:
        top_row = df.loc[
            df["Percentage"].idxmax()
        ]
        top = str(top_row["Name"])

    grades = {}

    if "Grade" in df.columns:
        grades = (
            df["Grade"]
            .value_counts()
            .to_dict()
        )

    subject_average = {}

    for subject in subjects:
        if subject in df.columns:
            subject_average[subject] = round(
                float(
                    pd.to_numeric(
                        df[subject],
                        errors="coerce"
                    ).mean()
                ),
                2
            )

    return jsonify({
        "total_students": len(df),

        "average_percentage": round(
            float(df["Percentage"].mean()),
            2
        ),

        "top_performer": top,

        "average_attendance": round(
            float(df["Attendance"].mean()),
            2
        ),

        "grades": grades,

        "subjects": subject_average
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

        data = request.get_json() or {}

        year = data.get(
            "year",
            "1st Year"
        )

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
                "Student ID and Name are required."
            })

        df = load_students(
            year,
            semester
        )

        if not df.empty:

            if student_id in (
                df["Student_ID"]
                .astype(str)
                .values
            ):

                return jsonify({
                    "success": False,
                    "message":
                    "Student ID already exists."
                })

        subjects = get_subject_codes(
            semester
        )

        new_student = {

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
            float(
                data.get(
                    "Attendance",
                    0
                )
            )
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

            new_student[subject] = mark

        if df.empty:

            raw_df = pd.DataFrame(
                [new_student]
            )

        else:

            raw_df = df.copy()

            # remove calculated columns
            for col in [
                "Total",
                "Percentage",
                "Attendance Status",
                "Grade"
            ]:
                if col in raw_df.columns:
                    raw_df = raw_df.drop(
                        columns=[col]
                    )

            raw_df = pd.concat(
                [
                    raw_df,
                    pd.DataFrame(
                        [new_student]
                    )
                ],
                ignore_index=True
            )

        save_students(
            year,
            raw_df
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


# =========================================================
# EXCEL UPLOAD ONLY
# =========================================================

@app.route(
    "/api/upload_excel",
    methods=["POST"]
)
def upload_excel():

    try:

        year = request.form.get(
            "year",
            "1st Year"
        )

        semester = request.form.get(
            "semester",
            "Semester 1"
        )

        if "file" not in request.files:

            return jsonify({
                "success": False,
                "message":
                "Please select an Excel file."
            })

        file = request.files["file"]

        if file.filename == "":

            return jsonify({
                "success": False,
                "message":
                "No file selected."
            })

        if not file.filename.lower().endswith(
            (".xlsx", ".xls")
        ):

            return jsonify({
                "success": False,
                "message":
                "Only Excel files (.xlsx/.xls) are allowed."
            })

        df = pd.read_excel(file)

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        required = [
            "Student_ID",
            "Name"
        ]

        missing = [
            x for x in required
            if x not in df.columns
        ]

        if missing:

            return jsonify({
                "success": False,
                "message":
                "Missing columns: "
                + ", ".join(missing)
            })

        # Save uploaded Excel
        filepath = get_year_file(year)

        df.to_excel(
            filepath,
            index=False
        )

        return jsonify({
            "success": True,
            "message":
            f"Excel uploaded successfully for {year} - {semester}."
        })

    except Exception as e:

        return jsonify({
            "success": False,
            "message": str(e)
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
        ),
        debug=False
    )
