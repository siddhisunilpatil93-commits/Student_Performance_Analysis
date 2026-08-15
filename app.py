from flask import Flask, render_template, jsonify, request
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

DATA_FOLDER = "data"
os.makedirs(DATA_FOLDER, exist_ok=True)

# ============================================================
# SEMESTER → FILE + ACADEMIC YEAR
# ============================================================

SEMESTER_CONFIG = {
    "Semester 1": {
        "file": "first_year.xlsx",
        "academic_year": "2025-26"
    },
    "Semester 2": {
        "file": "first_year.xlsx",
        "academic_year": "2025-26"
    },
    "Semester 3": {
        "file": "second_year.xlsx",
        "academic_year": "2026-27"
    },
    "Semester 4": {
        "file": "second_year.xlsx",
        "academic_year": "2026-27"
    },
    "Semester 5": {
        "file": "third_year.xlsx",
        "academic_year": "2027-28"
    },
    "Semester 6": {
        "file": "third_year.xlsx",
        "academic_year": "2027-28"
    }
}

STANDARD_COLUMNS = [
    "Student_ID",
    "Name",
    "Gender",
    "Class",
    "Total",
    "Percentage",
    "Attendance",
    "Attendance Status",
    "Attendance_Status",
    "Grade",
    "Semester"
]

KNOWN_SUBJECTS = [
    "OSY",
    "STE",
    "ACN",
    "DAN"
]


# ============================================================
# FIND HEADER ROW
# ============================================================

def find_header_row(filepath):

    try:
        raw = pd.read_excel(
            filepath,
            header=None,
            sheet_name=0
        )

        for index, row in raw.iterrows():

            values = [
                str(value).strip()
                for value in row.tolist()
                if pd.notna(value)
            ]

            if "Student_ID" in values:
                return index

    except Exception as e:
        print("HEADER SEARCH ERROR:", e)

    return 0


# ============================================================
# READ EXCEL
# ============================================================

def read_excel_file(filepath):

    header_row = find_header_row(filepath)

    df = pd.read_excel(
        filepath,
        header=header_row,
        sheet_name=0
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # Remove completely empty columns
    df = df.dropna(axis=1, how="all")

    # Remove completely empty rows
    df = df.dropna(axis=0, how="all")

    # Fix common column names
    rename_map = {}

    for column in df.columns:

        clean = str(column).strip()

        if clean.lower() in ["student id", "studentid", "roll no", "roll number"]:
            rename_map[column] = "Student_ID"

        elif clean.lower() == "attendance status":
            rename_map[column] = "Attendance Status"

    df = df.rename(columns=rename_map)

    return df


# ============================================================
# SUBJECT DETECTION
# ============================================================

def get_subject_columns(df):

    subjects = []

    for column in df.columns:

        column_name = str(column).strip()

        if column_name in STANDARD_COLUMNS:
            continue

        if column_name in ["Unnamed: 0", "Unnamed: 1"]:
            continue

        # Numeric subject columns
        if column_name not in [
            "Student_ID",
            "Name",
            "Gender",
            "Class"
        ]:
            subjects.append(column_name)

    # Prefer actual known subjects
    known = [
        subject
        for subject in KNOWN_SUBJECTS
        if subject in df.columns
    ]

    if known:
        return known

    return subjects


# ============================================================
# CALCULATE DATA
# ============================================================

def process_dataframe(df):

    if df.empty:
        return df

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    subjects = get_subject_columns(df)

    # Convert subject marks
    for subject in subjects:

        df[subject] = pd.to_numeric(
            df[subject],
            errors="coerce"
        ).fillna(0)

    # Total
    if subjects:

        df["Total"] = df[subjects].sum(axis=1)

        df["Percentage"] = (
            df["Total"] /
            (len(subjects) * 100)
        ) * 100

    # Attendance
    if "Attendance" in df.columns:

        df["Attendance"] = pd.to_numeric(
            df["Attendance"],
            errors="coerce"
        ).fillna(0)

        df["Attendance Status"] = df[
            "Attendance"
        ].apply(
            lambda value:
            "Good" if value >= 75 else "Low"
        )

    # Grade
    if "Percentage" in df.columns:

        def calculate_grade(value):

            try:
                value = float(value)
            except:
                return "F"

            if value >= 90:
                return "A+"
            elif value >= 80:
                return "A"
            elif value >= 70:
                return "B"
            elif value >= 60:
                return "C"
            elif value >= 50:
                return "D"
            else:
                return "F"

        df["Grade"] = df[
            "Percentage"
        ].apply(calculate_grade)

    df = df.fillna("")

    return df


# ============================================================
# LOAD STUDENTS
# ============================================================

def load_students(semester):

    config = SEMESTER_CONFIG.get(semester)

    if not config:
        return pd.DataFrame()

    filepath = os.path.join(
        DATA_FOLDER,
        config["file"]
    )

    print("Loading:", filepath)

    if not os.path.exists(filepath):

        print("FILE NOT FOUND:", filepath)

        return pd.DataFrame()

    try:

        df = read_excel_file(filepath)

        df = process_dataframe(df)

        return df

    except Exception as error:

        print("EXCEL ERROR:", error)

        return pd.DataFrame()


# ============================================================
# SAVE DATA
# ============================================================

def save_students(semester, df):

    config = SEMESTER_CONFIG.get(semester)

    if not config:
        return False

    filepath = os.path.join(
        DATA_FOLDER,
        config["file"]
    )

    try:

        df.to_excel(
            filepath,
            index=False
        )

        return True

    except Exception as error:

        print("SAVE ERROR:", error)

        return False


# ============================================================
# HOME
# ============================================================

@app.route("/")
def home():

    return render_template("index.html")


# ============================================================
# SEMESTER SUBJECT API
# ============================================================

@app.route("/api/subjects")
def subjects_api():

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    config = SEMESTER_CONFIG.get(
        semester,
        SEMESTER_CONFIG["Semester 1"]
    )

    df = load_students(semester)

    subjects = get_subject_columns(df)

    subject_data = []

    for subject in subjects:

        subject_data.append({
            "code": str(subject),
            "name": str(subject)
        })

    return jsonify({
        "success": True,
        "semester": semester,
        "academic_year": config["academic_year"],
        "subjects": subject_data
    })


# ============================================================
# STUDENTS API
# ============================================================

@app.route("/api/students")
def students_api():

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    df = load_students(semester)

    if df.empty:
        return jsonify([])

    return jsonify(
        df.to_dict(
            orient="records"
        )
    )


# ============================================================
# SEARCH API
# ============================================================

@app.route("/api/search")
def search_api():

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    query = request.args.get(
        "q",
        ""
    ).strip().lower()

    df = load_students(semester)

    if df.empty:
        return jsonify([])

    if not query:

        return jsonify(
            df.to_dict(
                orient="records"
            )
        )

    result = df[
        df.astype(str)
        .apply(
            lambda row:
            row.str.lower()
            .str.contains(
                query,
                na=False
            ).any(),
            axis=1
        )
    ]

    return jsonify(
        result.to_dict(
            orient="records"
        )
    )


# ============================================================
# DASHBOARD STATS
# ============================================================

@app.route("/api/stats")
def stats_api():

    semester = request.args.get(
        "semester",
        "Semester 1"
    )

    df = load_students(semester)

    if df.empty:

        return jsonify({
            "total_students": 0,
            "average_percentage": 0,
            "top_performer": "-",
            "average_attendance": 0
        })

    total_students = len(df)

    # Percentage
    if "Percentage" in df.columns:

        percentage = pd.to_numeric(
            df["Percentage"],
            errors="coerce"
        )

        average_percentage = round(
            percentage.mean(),
            2
        )

    else:

        average_percentage = 0

    # Top performer
    top_performer = "-"

    if (
        "Percentage" in df.columns
        and "Name" in df.columns
    ):

        temp = df.copy()

        temp["Percentage"] = pd.to_numeric(
            temp["Percentage"],
            errors="coerce"
        )

        temp = temp.dropna(
            subset=["Percentage"]
        )

        if not temp.empty:

            index = temp[
                "Percentage"
            ].idxmax()

            top_performer = str(
                temp.loc[
                    index,
                    "Name"
                ]
            )

    # Attendance
    if "Attendance" in df.columns:

        attendance = pd.to_numeric(
            df["Attendance"],
            errors="coerce"
        )

        average_attendance = round(
            attendance.mean(),
            2
        )

    else:

        average_attendance = 0

    return jsonify({

        "total_students":
            total_students,

        "average_percentage":
            average_percentage,

        "top_performer":
            top_performer,

        "average_attendance":
            average_attendance
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

        data = request.get_json()

        semester = data.get(
            "Semester",
            "Semester 1"
        )

        df = load_students(semester)

        config = SEMESTER_CONFIG.get(
            semester
        )

        if not config:

            return jsonify({
                "success": False,
                "message": "Invalid semester."
            })

        subjects = get_subject_columns(df)

        new_student = {}

        new_student["Student_ID"] = data.get(
            "Student_ID",
            ""
        )

        new_student["Name"] = data.get(
            "Name",
            ""
        )

        new_student["Gender"] = data.get(
            "Gender",
            ""
        )

        new_student["Class"] = data.get(
            "Class",
            "Computer Engineering"
        )

        for subject in subjects:

            value = data.get(
                subject,
                0
            )

            try:
                value = float(value)
            except:
                value = 0

            new_student[subject] = value

        attendance = data.get(
            "Attendance",
            0
        )

        try:
            attendance = float(attendance)
        except:
            attendance = 0

        new_student["Attendance"] = attendance

        # Calculate total
        total = sum(
            new_student.get(
                subject,
                0
            )
            for subject in subjects
        )

        new_student["Total"] = total

        if subjects:

            percentage = (
                total /
                (len(subjects) * 100)
            ) * 100

        else:

            percentage = 0

        new_student["Percentage"] = round(
            percentage,
            2
        )

        new_student["Attendance Status"] = (
            "Good"
            if attendance >= 75
            else "Low"
        )

        if percentage >= 90:
            grade = "A+"
        elif percentage >= 80:
            grade = "A"
        elif percentage >= 70:
            grade = "B"
        elif percentage >= 60:
            grade = "C"
        elif percentage >= 50:
            grade = "D"
        else:
            grade = "F"

        new_student["Grade"] = grade

        result = pd.concat(
            [
                df,
                pd.DataFrame([new_student])
            ],
            ignore_index=True
        )

        if save_students(
            semester,
            result
        ):

            return jsonify({
                "success": True,
                "message":
                    "Student added successfully."
            })

        return jsonify({
            "success": False,
            "message":
                "Unable to save student."
        })

    except Exception as error:

        print("ADD STUDENT ERROR:", error)

        return jsonify({
            "success": False,
            "message": str(error)
        })


# ============================================================
# UPLOAD EXCEL / CSV
# ============================================================

@app.route(
    "/api/upload",
    methods=["POST"]
)
def upload_file():

    try:

        if "file" not in request.files:

            return jsonify({
                "success": False,
                "message":
                    "Please select a file."
            })

        file = request.files["file"]

        if file.filename == "":

            return jsonify({
                "success": False,
                "message":
                    "No file selected."
            })

        semester = request.form.get(
            "semester",
            "Semester 1"
        )

        filename = secure_filename(
            file.filename
        )

        extension = os.path.splitext(
            filename
        )[1].lower()

        # Excel
        if extension in [
            ".xlsx",
            ".xls"
        ]:

            uploaded_df = pd.read_excel(
                file,
                header=None
            )

            header_row = None

            for index, row in uploaded_df.iterrows():

                values = [
                    str(x).strip()
                    for x in row.tolist()
                    if pd.notna(x)
                ]

                if "Student_ID" in values:

                    header_row = index
                    break

            if header_row is None:
                header_row = 0

            file.stream.seek(0)

            uploaded_df = pd.read_excel(
                file,
                header=header_row
            )

        # CSV
        elif extension == ".csv":

            uploaded_df = pd.read_csv(
                file
            )

        else:

            return jsonify({
                "success": False,
                "message":
                    "Please upload Excel or CSV file."
            })

        uploaded_df.columns = (
            uploaded_df.columns
            .astype(str)
            .str.strip()
        )

        uploaded_df = process_dataframe(
            uploaded_df
        )

        existing_df = load_students(
            semester
        )

        if existing_df.empty:

            final_df = uploaded_df

        else:

            final_df = pd.concat(
                [
                    existing_df,
                    uploaded_df
                ],
                ignore_index=True
            )

        if save_students(
            semester,
            final_df
        ):

            return jsonify({
                "success": True,
                "message":
                    "Student data imported successfully."
            })

        return jsonify({
            "success": False,
            "message":
                "Unable to save uploaded data."
        })

    except Exception as error:

        print("UPLOAD ERROR:", error)

        return jsonify({
            "success": False,
            "message": str(error)
        })


# ============================================================
# RUN
# ============================================================

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
