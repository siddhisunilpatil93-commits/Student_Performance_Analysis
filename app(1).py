from flask import Flask, render_template, jsonify, request
import pandas as pd
import os

app = Flask(__name__)

BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))

SEMESTER_YEAR = {
    "Semester 1": "1st Year",
    "Semester 2": "1st Year",
    "Semester 3": "2nd Year",
    "Semester 4": "2nd Year",
    "Semester 5": "3rd Year",
    "Semester 6": "3rd Year",
}

SEMESTER_FILES = {
    "Semester 1": "semester_1.xlsx",
    "Semester 2": "semester_2.xlsx",
    "Semester 3": "semester_3.xlsx",
    "Semester 4": "semester_4.xlsx",
    "Semester 5": "semester_5.xlsx",
    "Semester 6": "semester_6.xlsx",
}

FALLBACK_SUBJECTS = {
    "Semester 1": [
        "Basic Mathematics", "Communication Skills", "Engineering Physics",
        "Engineering Chemistry", "Basic Science"
    ],
    "Semester 2": [
        "Applied Mathematics", "Engineering Graphics",
        "Basic Electrical Engineering", "Programming in C", "Web Page Design"
    ],
    "Semester 3": [
        "Object Oriented Programming", "Data Structure",
        "Database Management System", "Computer Networks", "Operating System"
    ],
    "Semester 4": [
        "Java Programming", "Data Communication and Network",
        "Microprocessor", "Software Engineering", "Python Programming"
    ],
    "Semester 5": [
        "Advanced Java", "Web Based Application Development",
        "Software Testing", "Computer Security", "Project Management"
    ],
    "Semester 6": [
        "Mobile Application Development", "Cloud Computing",
        "Artificial Intelligence", "Internet of Things", "Major Project"
    ],
}

META_COLUMNS = {
    "Student_ID", "Name", "Gender", "Class", "Attendance",
    "Total", "Percentage", "Attendance Status", "Grade",
    "Semester", "Year"
}


def semester_file(semester):
    filename = SEMESTER_FILES.get(semester)
    return os.path.join(BASE_FOLDER, filename) if filename else None


def valid_semester(semester):
    return semester in SEMESTER_FILES


def clean_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def get_grade(value):
    try:
        p = float(value)
    except Exception:
        return "F"
    if p >= 90: return "A+"
    if p >= 80: return "A"
    if p >= 70: return "B+"
    if p >= 60: return "B"
    if p >= 50: return "C"
    if p >= 40: return "D"
    return "F"


def attendance_status(value):
    try:
        return "Good" if float(value) >= 75 else "Bad"
    except Exception:
        return "Bad"


def fallback_subjects(semester):
    return FALLBACK_SUBJECTS.get(semester, [])


def get_subject_columns(df, semester):
    if df is None or df.empty:
        return fallback_subjects(semester)

    subjects = []

    # Prefer the known Computer Engineering subjects when present.
    for subject in fallback_subjects(semester):
        if subject in df.columns:
            subjects.append(subject)

    # Also accept any other numeric columns from the uploaded Excel.
    for col in df.columns:
        if col in META_COLUMNS or col in subjects:
            continue
        numeric = pd.to_numeric(df[col], errors="coerce")
        if numeric.notna().sum() > 0:
            subjects.append(col)

    return subjects


def process_data(df, semester):
    if df is None or df.empty:
        return pd.DataFrame(columns=["Student_ID", "Name", "Gender", "Class", "Attendance"])

    df = clean_columns(df)

    # Make standard columns if absent.
    for col in ["Student_ID", "Name", "Gender", "Class", "Attendance"]:
        if col not in df.columns:
            df[col] = ""

    df["Student_ID"] = (
        df["Student_ID"].fillna("").astype(str)
        .str.replace(r"\.0$", "", regex=True).str.strip()
    )
    df["Name"] = df["Name"].fillna("").astype(str).str.strip()
    df["Gender"] = df["Gender"].fillna("").astype(str).str.strip()
    df["Class"] = df["Class"].fillna("").astype(str).str.strip()

    df["Attendance"] = pd.to_numeric(df["Attendance"], errors="coerce").fillna(0)
    df["Attendance"] = df["Attendance"].clip(0, 100).round(2)

    subjects = get_subject_columns(df, semester)

    for subject in subjects:
        if subject not in df.columns:
            df[subject] = 0
        df[subject] = pd.to_numeric(df[subject], errors="coerce").fillna(0).clip(0, 100)

    if subjects:
        df["Total"] = df[subjects].sum(axis=1).round(2)
        df["Percentage"] = (df["Total"] / (len(subjects) * 100) * 100).round(2)
    else:
        df["Total"] = 0
        df["Percentage"] = 0

    df["Grade"] = df["Percentage"].apply(get_grade)
    df["Attendance Status"] = df["Attendance"].apply(attendance_status)

    return df


def read_excel(semester):
    if not valid_semester(semester):
        return pd.DataFrame()

    path = semester_file(semester)
    if not path or not os.path.exists(path):
        return pd.DataFrame()

    try:
        # If workbook has a sheet with the exact semester name, use it.
        book = pd.ExcelFile(path)
        target_sheet = semester if semester in book.sheet_names else book.sheet_names[0]
        return clean_columns(pd.read_excel(path, sheet_name=target_sheet))
    except Exception as e:
        print("Excel read error:", e)
        return pd.DataFrame()


def split_combined_dataframe(df):
    """
    Handles an Excel sheet where Semester 1, 2, 3... are placed side-by-side
    and Student_ID/Name/... are repeated for every semester.
    Returns a list of semester dataframes.
    """
    df = clean_columns(df)
    cols = list(df.columns)
    starts = [i for i, c in enumerate(cols) if c == "Student_ID"]

    if len(starts) <= 1:
        return [df]

    blocks = []
    for pos, start in enumerate(starts):
        end = starts[pos + 1] if pos + 1 < len(starts) else len(cols)
        block_cols = cols[start:end]
        block = df.loc[:, block_cols].copy()
        # Remove duplicate column names inside a block if any.
        block = block.loc[:, ~block.columns.duplicated()]
        blocks.append(block)

    return blocks


def extract_uploaded_semester(file_storage, semester):
    """
    Accepts:
    1) a normal one-semester Excel,
    2) a workbook with Semester 1..6 sheets,
    3) one sheet containing semester blocks side-by-side.
    """
    raw = file_storage.read()
    from io import BytesIO

    xls = pd.ExcelFile(BytesIO(raw))
    desired_sheet = semester

    if desired_sheet in xls.sheet_names:
        return clean_columns(pd.read_excel(BytesIO(raw), sheet_name=desired_sheet))

    # If workbook has six semester sheets but exact names differ, use index.
    semester_index = int(semester.split()[-1]) - 1
    if len(xls.sheet_names) >= 6 and semester_index < len(xls.sheet_names):
        return clean_columns(pd.read_excel(BytesIO(raw), sheet_name=xls.sheet_names[semester_index]))

    # Otherwise read first sheet and split repeated Student_ID blocks.
    first = clean_columns(pd.read_excel(BytesIO(raw), sheet_name=xls.sheet_names[0]))
    blocks = split_combined_dataframe(first)

    if len(blocks) >= semester_index + 1:
        return blocks[semester_index]

    return first


def save_excel(df, semester):
    path = semester_file(semester)
    if not path:
        return False
    try:
        save_df = process_data(df, semester)
        calculated = ["Total", "Percentage", "Grade", "Attendance Status"]
        raw_df = save_df.drop(columns=[c for c in calculated if c in save_df.columns], errors="ignore")
        raw_df.to_excel(path, index=False)
        return True
    except Exception as e:
        print("Excel save error:", e)
        return False


@app.route("/")
def home():
    return render_template("index.html")


@app.after_request
def no_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.route("/api/subjects")
def subjects_api():
    semester = request.args.get("semester", "Semester 1")
    if not valid_semester(semester):
        return jsonify({"success": False, "subjects": []}), 400

    df = read_excel(semester)
    processed = process_data(df, semester) if not df.empty else df
    subjects = get_subject_columns(processed, semester)

    return jsonify({
        "success": True,
        "branch": "Computer Engineering",
        "scheme": "MSBTE K-Scheme",
        "semester": semester,
        "year": SEMESTER_YEAR[semester],
        "subjects": [{"code": s, "name": s} for s in subjects]
    })


@app.route("/api/students")
def students_api():
    semester = request.args.get("semester", "Semester 1")
    if not valid_semester(semester):
        return jsonify([])

    df = process_data(read_excel(semester), semester)
    if df.empty:
        return jsonify([])

    subjects = get_subject_columns(df, semester)
    columns = ["Student_ID", "Name", "Gender", "Class"] + subjects + [
        "Total", "Percentage", "Attendance", "Attendance Status", "Grade"
    ]
    columns = [c for c in columns if c in df.columns]

    return jsonify(df[columns].fillna("").to_dict(orient="records"))


@app.route("/api/analytics")
def analytics():
    semester = request.args.get("semester", "Semester 1")
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

    df = process_data(read_excel(semester), semester)
    if df.empty:
        return jsonify(empty)

    subjects = get_subject_columns(df, semester)
    top_name = "-"
    if not df.empty:
        idx = df["Percentage"].idxmax()
        top_name = str(df.loc[idx, "Name"]) or "-"

    subject_average = {
        subject: round(float(df[subject].mean()), 2)
        for subject in subjects if subject in df.columns
    }

    grades = {str(k): int(v) for k, v in df["Grade"].value_counts().items()}

    return jsonify({
        "total_students": int(len(df)),
        "average_percentage": round(float(df["Percentage"].mean()), 2),
        "top_performer": top_name,
        "average_attendance": round(float(df["Attendance"].mean()), 2),
        "subjects": subject_average,
        "grades": grades
    })


@app.route("/api/upload_excel", methods=["POST"])
def upload_excel():
    try:
        semester = request.form.get("semester", "Semester 1")

        if not valid_semester(semester):
            return jsonify({"success": False, "message": "Invalid semester selected."})

        if "file" not in request.files:
            return jsonify({"success": False, "message": "Please select an Excel file."})

        file = request.files["file"]
        if not file.filename:
            return jsonify({"success": False, "message": "No file selected."})

        if not file.filename.lower().endswith((".xlsx", ".xls")):
            return jsonify({"success": False, "message": "Only .xlsx or .xls files are allowed."})

        df = extract_uploaded_semester(file, semester)
        df = clean_columns(df)

        if "Student_ID" not in df.columns or "Name" not in df.columns:
            return jsonify({
                "success": False,
                "message": "Excel मध्ये Student_ID आणि Name columns पाहिजेत."
            })

        for col in ["Gender", "Class", "Attendance"]:
            if col not in df.columns:
                df[col] = "" if col != "Attendance" else 0

        # Ignore previously calculated values; system recalculates them.
        df = df.drop(
            columns=["Total", "Percentage", "Grade", "Attendance Status"],
            errors="ignore"
        )

        processed = process_data(df, semester)
        raw_df = processed.drop(
            columns=["Total", "Percentage", "Grade", "Attendance Status"],
            errors="ignore"
        )

        path = semester_file(semester)
        raw_df.to_excel(path, index=False)

        return jsonify({
            "success": True,
            "message": f"{semester} Excel successfully uploaded.",
            "students": int(len(raw_df)),
            "file": SEMESTER_FILES[semester]
        })

    except Exception as e:
        print("Upload error:", e)
        return jsonify({"success": False, "message": f"Upload failed: {str(e)}"})


@app.route("/api/add_student", methods=["POST"])
def add_student():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"success": False, "message": "Invalid student data."})

        semester = data.get("semester", "Semester 1")
        if not valid_semester(semester):
            return jsonify({"success": False, "message": "Invalid semester."})

        df = read_excel(semester)
        existing_subjects = get_subject_columns(df, semester)
        subjects = existing_subjects or fallback_subjects(semester)

        student = {
            "Student_ID": str(data.get("Student_ID", "")).strip(),
            "Name": str(data.get("Name", "")).strip(),
            "Gender": str(data.get("Gender", "")).strip(),
            "Class": str(data.get("Class", "")).strip(),
            "Attendance": data.get("Attendance", 0),
        }

        for subject in subjects:
            student[subject] = data.get(subject, 0)

        new_df = pd.DataFrame([student])

        if not df.empty:
            df = df.drop(
                columns=["Total", "Percentage", "Grade", "Attendance Status"],
                errors="ignore"
            )
            # Ensure both dataframes have the same subject/meta columns.
            for col in new_df.columns:
                if col not in df.columns:
                    df[col] = 0
            for col in df.columns:
                if col not in new_df.columns:
                    new_df[col] = 0
            new_df = new_df[df.columns]
            df = pd.concat([df, new_df], ignore_index=True)
        else:
            df = new_df

        if not save_excel(df, semester):
            return jsonify({"success": False, "message": "Student save failed."})

        return jsonify({"success": True, "message": "Student added successfully."})

    except Exception as e:
        print("Add student error:", e)
        return jsonify({"success": False, "message": str(e)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
