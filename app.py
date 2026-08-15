```python
from flask import Flask, render_template, jsonify, request
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

DATA_FOLDER = "data"
os.makedirs(DATA_FOLDER, exist_ok=True)

# Your existing Excel files
YEAR_FILES = {
    "1st Year": "first_year.xlsx",
    "2nd Year": "second_year.xlsx",
    "3rd Year": "third_year.xlsx"
}

# ---------------------------------------------------------
# SEMESTER + ACADEMIC YEAR + SUBJECTS
# ---------------------------------------------------------

SEMESTER_INFO = {
    "Semester 1": {
        "academic_year": "2025-26",
        "year": "1st Year",
        "subjects": [
            {"code": "BMS", "name": "Basic Mathematics"},
            {"code": "BEE", "name": "Basic Electrical Engineering"},
            {"code": "PPS", "name": "Programming in C"},
            {"code": "PCI", "name": "Professional Communication"},
            {"code": "WPD", "name": "Web Page Designing"}
        ]
    },

    "Semester 2": {
        "academic_year": "2025-26",
        "year": "1st Year",
        "subjects": [
            {"code": "AMS", "name": "Applied Mathematics"},
            {"code": "OOP", "name": "Object Oriented Programming"},
            {"code": "DMS", "name": "Digital and Microprocessor System"},
            {"code": "DBMS", "name": "Database Management System"},
            {"code": "CGR", "name": "Computer Graphics"}
        ]
    },

    "Semester 3": {
        "academic_year": "2026-27",
        "year": "2nd Year",
        "subjects": [
            {"code": "DSU", "name": "Data Structures Using C"},
            {"code": "DBMS", "name": "Database Management System"},
            {"code": "DCO", "name": "Digital Communication and Networking"},
            {"code": "OOP", "name": "Object Oriented Programming"},
            {"code": "JPR", "name": "Java Programming"}
        ]
    },

    "Semester 4": {
        "academic_year": "2026-27",
        "year": "2nd Year",
        "subjects": [
            {"code": "JAVA", "name": "Java Programming"},
            {"code": "DCN", "name": "Data Communication and Networking"},
            {"code": "MIC", "name": "Microprocessor Programming"},
            {"code": "GUI", "name": "Graphical User Interface"},
            {"code": "MAD", "name": "Mobile Application Development"}
        ]
    },

    "Semester 5": {
        "academic_year": "2027-28",
        "year": "3rd Year",
        "subjects": [
            {"code": "STE", "name": "Software Testing"},
            {"code": "DAN", "name": "Data Analytics"},
            {"code": "OSY", "name": "Operating System"},
            {"code": "ACN", "name": "Advanced Computer Network"},
            {"code": "WBP", "name": "Web Based Programming"}
        ]
    },

    "Semester 6": {
        "academic_year": "2027-28",
        "year": "3rd Year",
        "subjects": [
            {"code": "AI", "name": "Artificial Intelligence"},
            {"code": "ML", "name": "Machine Learning"},
            {"code": "CPE", "name": "Computer Project"},
            {"code": "MAD", "name": "Mobile Application Development"},
            {"code": "ENT", "name": "Entrepreneurship Development"}
        ]
    }
}


# ---------------------------------------------------------
# HELPERS
# ---------------------------------------------------------

def get_year_from_semester(semester):
    info = SEMESTER_INFO.get(semester)

    if not info:
        return None

    return info["year"]


def get_subject_codes(semester):
    info = SEMESTER_INFO.get(semester)

    if not info:
        return []

    return [subject["code"] for subject in info["subjects"]]


def find_column(df, possible_names):
    """
    Finds a column even if Excel has small naming differences.
    """
    normalized = {
        str(col).strip().lower().replace(" ", "_"): col
        for col in df.columns
    }

    for name in possible_names:
        key = name.strip().lower().replace(" ", "_")

        if key in normalized:
            return normalized[key]

    return None


# ---------------------------------------------------------
# LOAD STUDENTS
# ---------------------------------------------------------

def load_students(semester):

    info = SEMESTER_INFO.get(semester)

    if not info:
        return pd.DataFrame()

    year = info["year"]
    filename = YEAR_FILES.get(year)

    if not filename:
        return pd.DataFrame()

    filepath = os.path.join(DATA_FOLDER, filename)

    print("Loading:", filepath)

    if not os.path.exists(filepath):
        print("FILE NOT FOUND:", filepath)
        return pd.DataFrame()

    try:

        df = pd.read_excel(filepath)

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        if df.empty:
            return pd.DataFrame()

        # -------------------------------------------------
        # Standard columns
        # -------------------------------------------------

        id_col = find_column(
            df,
            ["Student_ID", "Student ID", "ID", "Roll_No", "Roll No"]
        )

        name_col = find_column(
            df,
            ["Name", "Student_Name", "Student Name"]
        )

        gender_col = find_column(
            df,
            ["Gender", "Sex"]
        )

        class_col = find_column(
            df,
            ["Class", "Division"]
        )

        attendance_col = find_column(
            df,
            ["Attendance", "Attendance %", "Attendance_Percentage"]
        )

        if id_col and id_col != "Student_ID":
            df["Student_ID"] = df[id_col]

        if name_col and name_col != "Name":
            df["Name"] = df[name_col]

        if gender_col and gender_col != "Gender":
            df["Gender"] = df[gender_col]

        if class_col and class_col != "Class":
            df["Class"] = df[class_col]

        if attendance_col and attendance_col != "Attendance":
            df["Attendance"] = df[attendance_col]

        # -------------------------------------------------
        # SUBJECTS
        # -------------------------------------------------

        subjects = get_subject_codes(semester)

        existing_subjects = []

        for subject in subjects:

            col = find_column(
                df,
                [subject]
            )

            if col:

                if col != subject:
                    df[subject] = df[col]

                df[subject] = pd.to_numeric(
                    df[subject],
                    errors="coerce"
                ).fillna(0)

                existing_subjects.append(subject)

            else:
                # Keep missing subject as 0
                df[subject] = 0

        # -------------------------------------------------
        # If semester-specific subjects are not present,
        # use existing marks columns from the Excel.
        # -------------------------------------------------

        if not existing_subjects:

            fallback_subjects = [
                col for col in
                ["OSY", "STE", "ACN", "DAN"]
                if col in df.columns
            ]

            existing_subjects = fallback_subjects

        # -------------------------------------------------
        # TOTAL
        # -------------------------------------------------

        if existing_subjects:

            df["Total"] = df[
                existing_subjects
            ].sum(axis=1)

            df["Percentage"] = (
                df["Total"] /
                (len(existing_subjects) * 100)
            ) * 100

        else:

            if "Total" not in df.columns:
                df["Total"] = 0

            if "Percentage" not in df.columns:
                df["Percentage"] = 0

        # -------------------------------------------------
        # ATTENDANCE
        # -------------------------------------------------

        if "Attendance" in df.columns:

            df["Attendance"] = pd.to_numeric(
                df["Attendance"],
                errors="coerce"
            ).fillna(0)

            df["Attendance_Status"] = df[
                "Attendance"
            ].apply(
                lambda x:
                "Good" if x >= 75 else "Low"
            )

        else:

            df["Attendance"] = 0
            df["Attendance_Status"] = "Low"

        # -------------------------------------------------
        # GRADE
        # -------------------------------------------------

        df["Percentage"] = pd.to_numeric(
            df["Percentage"],
            errors="coerce"
        ).fillna(0)

        def grade(p):

            if p >= 90:
                return "A+"
            elif p >= 80:
                return "A"
            elif p >= 70:
                return "B"
            elif p >= 60:
                return "C"
            elif p >= 50:
                return "D"
            else:
                return "F"

        df["Grade"] = df[
            "Percentage"
        ].apply(grade)

        # -------------------------------------------------
        # Clean NaN
        # -------------------------------------------------

        df = df.fillna("")

        return df

    except Exception as e:

        print("ERROR:", str(e))

        return pd.DataFrame()


# ---------------------------------------------------------
# SAVE EXCEL
# ---------------------------------------------------------

def save_students(semester, df):

    info = SEMESTER_INFO.get(semester)

    if not info:
        return False

    year = info["year"]
    filename = YEAR_FILES.get(year)

    if not filename:
        return False

    filepath = os.path.join(
        DATA_FOLDER,
        filename
    )

    try:

        df.to_excel(
            filepath,
            index=False
        )

        return True

    except Exception as e:

        print("SAVE ERROR:", e)

        return False


# ---------------------------------------------------------
# HOME
# ---------------------------------------------------------

@app.route("/")
def home():

    return render_template(
        "index.html"
    )


# ---------------------------------------------------------
# SUBJECTS
# ---------------------------------------------------------

@app.route("/api/subjects")
def subjects_api():

    semester = request.args.get(
        "semester",
        "Semester 3"
    )

    info = SEMESTER_INFO.get(
        semester
    )

    if not info:

        return jsonify({
            "success": False,
            "message": "Invalid semester"
        })

    return jsonify({
        "success": True,
        "semester": semester,
        "academic_year": info["academic_year"],
        "year": info["year"],
        "subjects": info["subjects"]
    })


# ---------------------------------------------------------
# STUDENTS
# ---------------------------------------------------------

@app.route("/api/students")
def students_api():

    semester = request.args.get(
        "semester",
        "Semester 3"
    )

    df = load_students(
        semester
    )

    return jsonify(
        df.to_dict(
            orient="records"
        )
    )


# ---------------------------------------------------------
# SEARCH
# ---------------------------------------------------------

@app.route("/api/search")
def search_student():

    semester = request.args.get(
        "semester",
        "Semester 3"
    )

    query = request.args.get(
        "q",
        ""
    ).strip().lower()

    df = load_students(
        semester
    )

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
            )
            .any(),
            axis=1
        )
    ]

    return jsonify(
        result.to_dict(
            orient="records"
        )
    )


# ---------------------------------------------------------
# DASHBOARD STATS
# ---------------------------------------------------------

@app.route("/api/stats")
def stats():

    semester = request.args.get(
        "semester",
        "Semester 3"
    )

    df = load_students(
        semester
    )

    if df.empty:

        return jsonify({
            "total_students": 0,
            "average_percentage": 0,
            "top_performer": "-",
            "average_attendance": 0
        })

    total_students = len(df)

    percentage = pd.to_numeric(
        df["Percentage"],
        errors="coerce"
    )

    average_percentage = round(
        percentage.mean(),
        2
    )

    top_performer = "-"

    if "Name" in df.columns:

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
                temp.loc[index, "Name"]
            )

    attendance = pd.to_numeric(
        df["Attendance"],
        errors="coerce"
    )

    average_attendance = round(
        attendance.mean(),
        2
    )

    return jsonify({
        "total_students": total_students,
        "average_percentage": average_percentage,
        "top_performer": top_performer,
        "average_attendance": average_attendance
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

        if not data:

            return jsonify({
                "success": False,
                "message": "No student data received."
            })

        semester = data.get(
            "Semester",
            "Semester 3"
        )

        df = load_students(
            semester
        )

        # -------------------------------------------------
        # Create empty dataframe if Excel is empty
        # -------------------------------------------------

        if df.empty:

            df = pd.DataFrame()

        new_row = {}

        new_row["Student_ID"] = data.get(
            "Student_ID",
            ""
        )

        new_row["Name"] = data.get(
            "Name",
            ""
        )

        new_row["Gender"] = data.get(
            "Gender",
            ""
        )

        new_row["Class"] = data.get(
            "Class",
            ""
        )

        subjects = get_subject_codes(
            semester
        )

        total = 0

        for subject in subjects:

            value = pd.to_numeric(
                data.get(subject, 0),
                errors="coerce"
            )

            if pd.isna(value):
                value = 0

            new_row[subject] = float(
                value
            )

            total += float(value)

        new_row["Total"] = total

        if subjects:

            percentage = (
                total /
                (len(subjects) * 100)
            ) * 100

        else:

            percentage = 0

        new_row["Percentage"] = round(
            percentage,
            2
        )

        attendance = pd.to_numeric(
            data.get("Attendance", 0),
            errors="coerce"
        )

        if pd.isna(attendance):
            attendance = 0

        new_row["Attendance"] = float(
            attendance
        )

        new_row["Attendance_Status"] = (
            "Good"
            if attendance >= 75
            else "Low"
        )

        if percentage >= 90:
            new_row["Grade"] = "A+"
        elif percentage >= 80:
            new_row["Grade"] = "A"
        elif percentage >= 70:
            new_row["Grade"] = "B"
        elif percentage >= 60:
            new_row["Grade"] = "C"
        elif percentage >= 50:
            new_row["Grade"] = "D"
        else:
            new_row["Grade"] = "F"

        new_df = pd.DataFrame(
            [new_row]
        )

        # -------------------------------------------------
        # Merge
        # -------------------------------------------------

        if df.empty:

            final_df = new_df

        else:

            final_df = pd.concat(
                [df, new_df],
                ignore_index=True
            )

        if save_students(
            semester,
            final_df
        ):

            return jsonify({
                "success": True,
                "message": "Student added successfully."
            })

        return jsonify({
            "success": False,
            "message": "Unable to save student."
        })

    except Exception as e:

        print(
            "ADD STUDENT ERROR:",
            str(e)
        )

        return jsonify({
            "success": False,
            "message": str(e)
        })


# ---------------------------------------------------------
# UPLOAD EXCEL / CSV
# ---------------------------------------------------------

@app.route(
    "/api/upload",
    methods=["POST"]
)
def upload_file():

    try:

        if "file" not in request.files:

            return jsonify({
                "success": False,
                "message": "Please select a file."
            })

        file = request.files["file"]

        if file.filename == "":

            return jsonify({
                "success": False,
                "message": "No file selected."
            })

        semester = request.form.get(
            "semester",
            "Semester 3"
        )

        filename = secure_filename(
            file.filename
        )

        extension = os.path.splitext(
            filename
        )[1].lower()

        if extension not in [
            ".xlsx",
            ".xls",
            ".csv"
        ]:

            return jsonify({
                "success": False,
                "message": "Only Excel or CSV files are supported."
            })

        # Read uploaded file
        if extension == ".csv":

            uploaded_df = pd.read_csv(
                file
            )

        else:

            uploaded_df = pd.read_excel(
                file
            )

        if uploaded_df.empty:

            return jsonify({
                "success": False,
                "message": "Uploaded file is empty."
            })

        uploaded_df.columns = (
            uploaded_df.columns
            .astype(str)
            .str.strip()
        )

        existing_df = load_students(
            semester
        )

        if existing_df.empty:

            final_df = uploaded_df

        else:

            final_df = pd.concat(
                [existing_df, uploaded_df],
                ignore_index=True
            )

        # Remove duplicate Student IDs
        if "Student_ID" in final_df.columns:

            final_df = final_df.drop_duplicates(
                subset=["Student_ID"],
                keep="last"
            )

        # Recalculate data
        subjects = get_subject_codes(
            semester
        )

        existing_subjects = []

        for subject in subjects:

            if subject in final_df.columns:

                final_df[subject] = pd.to_numeric(
                    final_df[subject],
                    errors="coerce"
                ).fillna(0)

                existing_subjects.append(
                    subject
                )

        if existing_subjects:

            final_df["Total"] = final_df[
                existing_subjects
            ].sum(axis=1)

            final_df["Percentage"] = (
                final_df["Total"] /
                (len(existing_subjects) * 100)
            ) * 100

        if "Attendance" in final_df.columns:

            final_df["Attendance"] = pd.to_numeric(
                final_df["Attendance"],
                errors="coerce"
            ).fillna(0)

            final_df["Attendance_Status"] = (
                final_df["Attendance"]
                .apply(
                    lambda x:
                    "Good"
                    if x >= 75
                    else "Low"
                )
            )

        if "Percentage" in final_df.columns:

            final_df["Grade"] = (
                final_df["Percentage"]
                .apply(
                    lambda p:
                    "A+" if p >= 90
                    else "A" if p >= 80
                    else "B" if p >= 70
                    else "C" if p >= 60
                    else "D" if p >= 50
                    else "F"
                )
            )

        if save_students(
            semester,
            final_df
        ):

            return jsonify({
                "success": True,
                "message": "Student data imported successfully."
            })

        return jsonify({
            "success": False,
            "message": "Could not save uploaded data."
        })

    except Exception as e:

        print(
            "UPLOAD ERROR:",
            str(e)
        )

        return jsonify({
            "success": False,
            "message": str(e)
        })


# ---------------------------------------------------------
# RUN
# ---------------------------------------------------------

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
```
