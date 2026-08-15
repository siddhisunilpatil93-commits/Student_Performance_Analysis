from flask import Flask, render_template, jsonify, request
import pandas as pd
import os

app = Flask(__name__)

DATA_FOLDER = "data"

YEAR_FILES = {
    "1st Year": "first_year.xlsx",
    "2nd Year": "second_year.xlsx",
    "3rd Year": "third_year.xlsx"
}

SUBJECTS = ["OSY", "STE", "ACN", "DAN"]


# =========================================================
# READ STUDENT FILE
# =========================================================

def load_students(year):

    filename = YEAR_FILES.get(year)

    if not filename:
        return pd.DataFrame()

    filepath = os.path.join(DATA_FOLDER, filename)

    if not os.path.exists(filepath):
        return pd.DataFrame()

    try:

        df = pd.read_excel(filepath)

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        # -------------------------------------------------
        # MAKE SURE BASIC COLUMNS EXIST
        # -------------------------------------------------

        for col in ["Student_ID", "Name", "Gender", "Class"]:
            if col not in df.columns:
                df[col] = ""

        # -------------------------------------------------
        # SUBJECT MARKS
        # -------------------------------------------------

        existing_subjects = []

        for subject in SUBJECTS:

            if subject in df.columns:

                df[subject] = pd.to_numeric(
                    df[subject],
                    errors="coerce"
                ).fillna(0)

                existing_subjects.append(subject)

        # -------------------------------------------------
        # TOTAL
        # -------------------------------------------------

        if existing_subjects:

            df["Total"] = df[
                existing_subjects
            ].sum(axis=1)

            max_marks = len(existing_subjects) * 100

            df["Percentage"] = (
                df["Total"] / max_marks
            ) * 100

        else:

            df["Total"] = 0
            df["Percentage"] = 0

        # -------------------------------------------------
        # ATTENDANCE
        # -------------------------------------------------

        if "Attendance" not in df.columns:
            df["Attendance"] = 0

        df["Attendance"] = pd.to_numeric(
            df["Attendance"],
            errors="coerce"
        ).fillna(0)

        df["Attendance Status"] = df[
            "Attendance"
        ].apply(
            lambda x: "Good" if x >= 75 else "Low"
        )

        # -------------------------------------------------
        # GRADE
        # -------------------------------------------------

        def calculate_grade(p):

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
        ].apply(calculate_grade)

        # -------------------------------------------------
        # ROUND VALUES
        # -------------------------------------------------

        df["Total"] = df["Total"].round(2)
        df["Percentage"] = df["Percentage"].round(2)
        df["Attendance"] = df["Attendance"].round(2)

        df = df.fillna("")

        return df

    except Exception as e:

        print("ERROR:", e)

        return pd.DataFrame()


# =========================================================
# SAVE DATA
# =========================================================

def save_students(year, df):

    filename = YEAR_FILES.get(year)

    if not filename:
        return False

    os.makedirs(DATA_FOLDER, exist_ok=True)

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


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():

    return render_template(
        "index.html"
    )


# =========================================================
# GET STUDENTS
# =========================================================

@app.route("/api/students")
def students():

    year = request.args.get(
        "year",
        "3rd Year"
    )

    df = load_students(year)

    return jsonify(
        df.to_dict(
            orient="records"
        )
    )


# =========================================================
# SEARCH
# =========================================================

@app.route("/api/search")
def search_student():

    year = request.args.get(
        "year",
        "3rd Year"
    )

    query = request.args.get(
        "q",
        ""
    ).strip().lower()

    df = load_students(year)

    if df.empty:
        return jsonify([])

    if not query:

        return jsonify(
            df.to_dict(
                orient="records"
            )
        )

    result = df[
        df.apply(
            lambda row:
            query in str(
                row.get(
                    "Name",
                    ""
                )
            ).lower()
            or
            query in str(
                row.get(
                    "Student_ID",
                    ""
                )
            ).lower(),
            axis=1
        )
    ]

    return jsonify(
        result.to_dict(
            orient="records"
        )
    )


# =========================================================
# DASHBOARD STATS
# =========================================================

@app.route("/api/stats")
def stats():

    year = request.args.get(
        "year",
        "3rd Year"
    )

    df = load_students(year)

    if df.empty:

        return jsonify({
            "total_students": 0,
            "average_percentage": 0,
            "top_performer": "-",
            "average_attendance": 0
        })

    total_students = len(df)

    # Average Percentage
    average_percentage = round(
        pd.to_numeric(
            df["Percentage"],
            errors="coerce"
        ).mean(),
        2
    )

    # Top Performer
    top_performer = "-"

    if "Name" in df.columns:

        index = df[
            "Percentage"
        ].idxmax()

        if index is not None:

            top_performer = str(
                df.loc[
                    index,
                    "Name"
                ]
            )

    # Average Attendance
    average_attendance = round(
        pd.to_numeric(
            df["Attendance"],
            errors="coerce"
        ).mean(),
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

        year = data.get(
            "year",
            "3rd Year"
        )

        df = load_students(year)

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

        # Check duplicate ID
        if not df.empty:

            if "Student_ID" in df.columns:

                duplicate = df[
                    df["Student_ID"]
                    .astype(str)
                    .str.lower()
                    == student_id.lower()
                ]

                if not duplicate.empty:

                    return jsonify({
                        "success": False,
                        "message":
                            "Student ID already exists."
                    })

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
                )
        }

        # Subjects
        for subject in SUBJECTS:

            new_student[subject] = float(
                data.get(
                    subject,
                    0
                ) or 0
            )

        new_student["Attendance"] = float(
            data.get(
                "Attendance",
                0
            ) or 0
        )

        new_df = pd.DataFrame(
            [new_student]
        )

        # Combine
        if df.empty:

            final_df = new_df

        else:

            final_df = pd.concat(
                [
                    df[
                        [
                            c for c in df.columns
                            if c not in [
                                "Total",
                                "Percentage",
                                "Attendance Status",
                                "Grade"
                            ]
                        ]
                    ],
                    new_df
                ],
                ignore_index=True
            )

        if save_students(
            year,
            final_df
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

    except Exception as e:

        print("ADD STUDENT ERROR:", e)

        return jsonify({
            "success": False,
            "message":
                "Error while adding student."
        })


# =========================================================
# UPLOAD CSV / EXCEL
# =========================================================

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
                    "Please select a file."
            })

        year = request.form.get(
            "year",
            "3rd Year"
        )

        filename = file.filename.lower()

        # -------------------------------------------------
        # READ CSV
        # -------------------------------------------------

        if filename.endswith(".csv"):

            uploaded_df = pd.read_csv(
                file
            )

        # -------------------------------------------------
        # READ EXCEL
        # -------------------------------------------------

        elif (
            filename.endswith(".xlsx")
            or filename.endswith(".xls")
        ):

            uploaded_df = pd.read_excel(
                file
            )

        else:

            return jsonify({
                "success": False,
                "message":
                    "Only CSV or Excel files are supported."
            })

        uploaded_df.columns = (
            uploaded_df.columns
            .astype(str)
            .str.strip()
        )

        if uploaded_df.empty:

            return jsonify({
                "success": False,
                "message":
                    "File is empty."
            })

        # -------------------------------------------------
        # EXISTING DATA
        # -------------------------------------------------

        existing_df = load_students(
            year
        )

        # Remove calculated columns
        calculated_columns = [
            "Total",
            "Percentage",
            "Attendance Status",
            "Grade"
        ]

        if not existing_df.empty:

            existing_df = existing_df[
                [
                    c for c in existing_df.columns
                    if c not in calculated_columns
                ]
            ]

        uploaded_df = uploaded_df[
            [
                c for c in uploaded_df.columns
                if c not in calculated_columns
            ]
        ]

        # -------------------------------------------------
        # COMBINE
        # -------------------------------------------------

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

        # Remove duplicate Student IDs
        if "Student_ID" in final_df.columns:

            final_df = final_df.drop_duplicates(
                subset=["Student_ID"],
                keep="last"
            )

        # Save
        if save_students(
            year,
            final_df
        ):

            return jsonify({
                "success": True,
                "message":
                    f"{len(uploaded_df)} student record(s) imported successfully."
            })

        return jsonify({
            "success": False,
            "message":
                "Unable to save uploaded data."
        })

    except Exception as e:

        print("UPLOAD ERROR:", e)

        return jsonify({
            "success": False,
            "message":
                "File upload failed. Check the file format."
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
