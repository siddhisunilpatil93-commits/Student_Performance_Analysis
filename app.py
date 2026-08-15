from flask import Flask, render_template, jsonify, request
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

# ==========================================================
# SETTINGS
# ==========================================================

DATA_FOLDER = "data"
os.makedirs(DATA_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {"xlsx", "xls", "docx", "pdf"}

# ==========================================================
# SEMESTER CONFIGURATION
# ==========================================================

SEMESTERS = {

    "Semester 1": {
        "academic_year": "2025-26",
        "subjects": [
            {"code": "BMS", "name": "Basic Mathematics"},
            {"code": "BEE", "name": "Basic Electrical Engineering"},
            {"code": "PPS", "name": "Programming in Python"},
            {"code": "FPL", "name": "Fundamentals of Programming Languages"},
            {"code": "BSC", "name": "Basic Science"}
        ]
    },

    "Semester 2": {
        "academic_year": "2025-26",
        "subjects": [
            {"code": "AMS", "name": "Applied Mathematics"},
            {"code": "DCS", "name": "Digital Communication System"},
            {"code": "PDS", "name": "Programming and Data Structure"},
            {"code": "DBMS", "name": "Database Management System"},
            {"code": "DTE", "name": "Digital Techniques"}
        ]
    },

    "Semester 3": {
        "academic_year": "2026-27",
        "subjects": [
            {"code": "OSY", "name": "Operating System"},
            {"code": "STE", "name": "Software Testing"},
            {"code": "ACN", "name": "Advanced Computer Network"},
            {"code": "DAN", "name": "Data Analytics"}
        ]
    },

    "Semester 4": {
        "academic_year": "2026-27",
        "subjects": [
            {"code": "JAVA", "name": "Java Programming"},
            {"code": "DCN", "name": "Data Communication and Network"},
            {"code": "MIC", "name": "Microprocessor"},
            {"code": "GUI", "name": "Graphical User Interface"}
        ]
    },

    "Semester 5": {
        "academic_year": "2027-28",
        "subjects": [
            {"code": "WAD", "name": "Web Application Development"},
            {"code": "MAD", "name": "Mobile Application Development"},
            {"code": "CNS", "name": "Computer Network Security"},
            {"code": "AI", "name": "Artificial Intelligence"}
        ]
    },

    "Semester 6": {
        "academic_year": "2027-28",
        "subjects": [
            {"code": "ML", "name": "Machine Learning"},
            {"code": "CC", "name": "Cloud Computing"},
            {"code": "IOT", "name": "Internet of Things"},
            {"code": "PWP", "name": "Project Work"}
        ]
    }
}


# ==========================================================
# FILE NAME FOR EACH SEMESTER
# ==========================================================

SEMESTER_FILES = {

    "Semester 1": "semester_1.xlsx",
    "Semester 2": "semester_2.xlsx",
    "Semester 3": "semester_3.xlsx",
    "Semester 4": "semester_4.xlsx",
    "Semester 5": "semester_5.xlsx",
    "Semester 6": "semester_6.xlsx"

}


# ==========================================================
# HELPER FUNCTIONS
# ==========================================================

def allowed_file(filename):

    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower()
        in ALLOWED_EXTENSIONS
    )


def get_subject_codes(semester):

    if semester not in SEMESTERS:
        return []

    return [
        subject["code"]
        for subject in SEMESTERS[semester]["subjects"]
    ]


def get_excel_path(semester):

    filename = SEMESTER_FILES.get(semester)

    if not filename:
        return None

    return os.path.join(
        DATA_FOLDER,
        filename
    )


# ==========================================================
# CREATE EMPTY EXCEL FILE
# ==========================================================

def create_empty_excel(semester):

    filepath = get_excel_path(semester)

    if not filepath:
        return

    if os.path.exists(filepath):
        return

    subjects = get_subject_codes(semester)

    columns = [
        "Student_ID",
        "Name",
        "Gender",
        "Class"
    ]

    columns.extend(subjects)

    columns.extend([
        "Total",
        "Percentage",
        "Attendance",
        "Attendance_Status",
        "Grade"
    ])

    df = pd.DataFrame(columns=columns)

    df.to_excel(
        filepath,
        index=False
    )


# ==========================================================
# LOAD STUDENTS
# ==========================================================

def load_students(semester):

    if semester not in SEMESTERS:

        return pd.DataFrame()

    filepath = get_excel_path(semester)

    if not filepath:

        return pd.DataFrame()

    # Automatically create Excel file
    create_empty_excel(semester)

    if not os.path.exists(filepath):

        return pd.DataFrame()

    try:

        df = pd.read_excel(filepath)

        # Clean column names
        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        subjects = get_subject_codes(
            semester
        )

        # Make sure subject columns exist
        for subject in subjects:

            if subject not in df.columns:

                df[subject] = 0

            df[subject] = pd.to_numeric(
                df[subject],
                errors="coerce"
            ).fillna(0)

        # ==================================================
        # TOTAL
        # ==================================================

        df["Total"] = df[
            subjects
        ].sum(axis=1)

        # ==================================================
        # PERCENTAGE
        # ==================================================

        if len(subjects) > 0:

            df["Percentage"] = (
                df["Total"]
                /
                (len(subjects) * 100)
            ) * 100

        else:

            df["Percentage"] = 0

        df["Percentage"] = df[
            "Percentage"
        ].round(2)

        # ==================================================
        # ATTENDANCE
        # ==================================================

        if "Attendance" not in df.columns:

            df["Attendance"] = 0

        df["Attendance"] = pd.to_numeric(
            df["Attendance"],
            errors="coerce"
        ).fillna(0)

        df["Attendance"] = df[
            "Attendance"
        ].clip(0, 100)

        # ==================================================
        # ATTENDANCE STATUS
        # ==================================================

        df["Attendance_Status"] = df[
            "Attendance"
        ].apply(
            lambda x:
            "Good" if x >= 75 else "Low"
        )

        # ==================================================
        # GRADE
        # ==================================================

        def calculate_grade(percent):

            if percent >= 90:
                return "A+"

            elif percent >= 80:
                return "A"

            elif percent >= 70:
                return "B+"

            elif percent >= 60:
                return "B"

            elif percent >= 50:
                return "C"

            elif percent >= 40:
                return "D"

            else:
                return "F"

        df["Grade"] = df[
            "Percentage"
        ].apply(
            calculate_grade
        )

        # Replace NaN
        df = df.fillna("")

        return df

    except Exception as e:

        print(
            "ERROR READING EXCEL:",
            e
        )

        return pd.DataFrame()


# ==========================================================
# SAVE STUDENTS
# ==========================================================

def save_students(
    semester,
    df
):

    filepath = get_excel_path(
        semester
    )

    if not filepath:
        return False

    try:

        df.to_excel(
            filepath,
            index=False
        )

        return True

    except Exception as e:

        print(
            "ERROR SAVING EXCEL:",
            e
        )

        return False


# ==========================================================
# HOME
# ==========================================================

@app.route("/")
def home():

    return render_template(
        "index.html"
    )


# ==========================================================
# SUBJECT API
# ==========================================================

@app.route("/api/subjects")
def subjects_api():

    semester = request.args.get(
        "semester",
        "Semester 3"
    )

    if semester not in SEMESTERS:

        return jsonify({
            "success": False,
            "message": "Invalid semester"
        })

    return jsonify({

        "success": True,

        "semester": semester,

        "academic_year":
            SEMESTERS[
                semester
            ]["academic_year"],

        "subjects":
            SEMESTERS[
                semester
            ]["subjects"]

    })


# ==========================================================
# STUDENTS API
# ==========================================================

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


# ==========================================================
# SEARCH API
# ==========================================================

@app.route("/api/search")
def search_api():

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


# ==========================================================
# DASHBOARD STATS
# ==========================================================

@app.route("/api/stats")
def stats_api():

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

    # ==================================================
    # AVERAGE PERCENTAGE
    # ==================================================

    percentage = pd.to_numeric(
        df["Percentage"],
        errors="coerce"
    )

    average_percentage = round(
        percentage.mean(),
        2
    )

    # ==================================================
    # TOP PERFORMER
    # ==================================================

    top_performer = "-"

    if "Name" in df.columns:

        temp = df.copy()

        temp["Percentage"] = pd.to_numeric(
            temp["Percentage"],
            errors="coerce"
        )

        temp = temp.dropna(
            subset=[
                "Percentage"
            ]
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

    # ==================================================
    # ATTENDANCE
    # ==================================================

    attendance = pd.to_numeric(
        df["Attendance"],
        errors="coerce"
    )

    average_attendance = round(
        attendance.mean(),
        2
    )

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


# ==========================================================
# ADD STUDENT
# ==========================================================

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
                    "No student data received."

            })

        semester = data.get(
            "Semester",
            "Semester 3"
        )

        if semester not in SEMESTERS:

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

        if not student_id:

            return jsonify({

                "success": False,

                "message":
                    "Student ID is required."

            })

        if not name:

            return jsonify({

                "success": False,

                "message":
                    "Student Name is required."

            })

        df = load_students(
            semester
        )

        # ==================================================
        # DUPLICATE ID CHECK
        # ==================================================

        if (
            not df.empty
            and "Student_ID" in df.columns
        ):

            existing_ids = (
                df["Student_ID"]
                .astype(str)
                .str.strip()
            )

            if student_id in (
                existing_ids.values
            ):

                return jsonify({

                    "success": False,

                    "message":
                        "Student ID already exists."

                })

        # ==================================================
        # CREATE NEW ROW
        # ==================================================

        new_row = {

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

        subjects = get_subject_codes(
            semester
        )

        for subject in subjects:

            new_row[
                subject
            ] = data.get(
                subject,
                0
            )

        # ==================================================
        # ADD ROW
        # ==================================================

        if df.empty:

            df = pd.DataFrame(
                [new_row]
            )

        else:

            df = pd.concat(
                [
                    df,
                    pd.DataFrame(
                        [new_row]
                    )
                ],
                ignore_index=True
            )

        # Recalculate everything
        temp_path = get_excel_path(
            semester
        )

        df.to_excel(
            temp_path,
            index=False
        )

        # Reload calculated dataframe
        final_df = load_students(
            semester
        )

        final_df.to_excel(
            temp_path,
            index=False
        )

        return jsonify({

            "success": True,

            "message":
                "Student added successfully."

        })

    except Exception as e:

        print(
            "ADD STUDENT ERROR:",
            e
        )

        return jsonify({

            "success": False,

            "message":
                "Error adding student."

        })


# ==========================================================
# UPLOAD EXCEL / WORD / PDF
# ==========================================================

@app.route(
    "/api/upload",
    methods=["POST"]
)
def upload_file():

    try:

        semester = request.form.get(
            "semester",
            "Semester 3"
        )

        if semester not in SEMESTERS:

            return jsonify({

                "success": False,

                "message":
                    "Invalid semester."

            })

        if "file" not in request.files:

            return jsonify({

                "success": False,

                "message":
                    "No file selected."

            })

        file = request.files[
            "file"
        ]

        if file.filename == "":

            return jsonify({

                "success": False,

                "message":
                    "Please select a file."

            })

        if not allowed_file(
            file.filename
        ):

            return jsonify({

                "success": False,

                "message":
                    "Only Excel, Word or PDF files are allowed."

            })

        extension = (
            file.filename
            .rsplit(".", 1)[1]
            .lower()
        )

        # ==================================================
        # EXCEL
        # ==================================================

        if extension in {
            "xlsx",
            "xls"
        }:

            uploaded_df = pd.read_excel(
                file
            )

            if uploaded_df.empty:

                return jsonify({

                    "success": False,

                    "message":
                        "Excel file is empty."

                })

            uploaded_df.columns = (
                uploaded_df.columns
                .astype(str)
                .str.strip()
            )

            subjects = get_subject_codes(
                semester
            )

            required = [
                "Student_ID",
                "Name"
            ]

            missing = [
                column
                for column in required
                if column not in
                uploaded_df.columns
            ]

            if missing:

                return jsonify({

                    "success": False,

                    "message":
                        "Excel must contain Student_ID and Name."

                })

            current_df = load_students(
                semester
            )

            if current_df.empty:

                combined = uploaded_df

            else:

                combined = pd.concat(
                    [
                        current_df,
                        uploaded_df
                    ],
                    ignore_index=True
                )

                if "Student_ID" in combined.columns:

                    combined = (
                        combined
                        .drop_duplicates(
                            subset=[
                                "Student_ID"
                            ],
                            keep="last"
                        )
                    )

            path = get_excel_path(
                semester
            )

            combined.to_excel(
                path,
                index=False
            )

            # Recalculate
            final_df = load_students(
                semester
            )

            final_df.to_excel(
                path,
                index=False
            )

            return jsonify({

                "success": True,

                "message":
                    "Excel data imported successfully."

            })

        # ==================================================
        # WORD / PDF
        # ==================================================

        if extension in {
            "docx",
            "pdf"
        }:

            return jsonify({

                "success": False,

                "message":
                    "Excel import is fully supported. Word/PDF upload is accepted but student-table extraction requires a structured table format."

            })

    except Exception as e:

        print(
            "UPLOAD ERROR:",
            e
        )

        return jsonify({

            "success": False,

            "message":
                "File upload failed."

        })


# ==========================================================
# RUN
# ==========================================================

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
