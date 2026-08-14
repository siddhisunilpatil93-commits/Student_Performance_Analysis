from flask import Flask, render_template, jsonify, request
import pandas as pd
import os
import shutil

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

os.makedirs(DATA_DIR, exist_ok=True)

# =========================================================
# YEAR FILES
# =========================================================

YEAR_FILES = {
    "1st Year": os.path.join(DATA_DIR, "first_year.csv"),
    "2nd Year": os.path.join(DATA_DIR, "second_year.csv"),
    "3rd Year": os.path.join(DATA_DIR, "third_year.csv")
}


# =========================================================
# OLD DATA -> 3RD YEAR
# =========================================================

old_files = [
    os.path.join(BASE_DIR, "student_result.csv"),
    os.path.join(BASE_DIR, "student_data.csv")
]

third_file = YEAR_FILES["3rd Year"]

if not os.path.exists(third_file):

    for old_file in old_files:

        if os.path.exists(old_file):

            try:
                shutil.copy2(old_file, third_file)
                print("Old student data copied to 3rd Year.")
                break

            except Exception as e:
                print("Migration error:", e)


# =========================================================
# HOME PAGE
# =========================================================

@app.route("/")
def home():
    return render_template("index.html")


# =========================================================
# PREPARE DATA
# =========================================================

def prepare_dataframe(df):

    df = df.copy()

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    required_columns = [
        "Student_ID",
        "Name",
        "Gender",
        "Class",
        "OSY",
        "STE",
        "ACN",
        "DAN",
        "Attendance"
    ]

    for col in required_columns:

        if col not in df.columns:
            df[col] = ""


    # -------------------------
    # Numeric Marks
    # -------------------------

    for col in ["OSY", "STE", "ACN", "DAN"]:

        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        ).fillna(0)


    # -------------------------
    # Attendance
    # -------------------------

    df["Attendance"] = (
        df["Attendance"]
        .astype(str)
        .str.replace("%", "", regex=False)
        .str.strip()
    )

    df["Attendance"] = pd.to_numeric(
        df["Attendance"],
        errors="coerce"
    ).fillna(0)


    # -------------------------
    # Total
    # -------------------------

    df["Total"] = (
        df["OSY"]
        + df["STE"]
        + df["ACN"]
        + df["DAN"]
    )


    # -------------------------
    # Percentage
    # -------------------------

    df["Percentage"] = (
        df["Total"] / 4
    ).round(2)


    # -------------------------
    # Grade
    # -------------------------

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

        else:
            return "F"


    df["Grade"] = (
        df["Percentage"]
        .apply(calculate_grade)
    )


    # -------------------------
    # Attendance Status
    # -------------------------

    df["Status"] = df["Attendance"].apply(
        lambda x:
        "Eligible"
        if x >= 75
        else "Shortage"
    )


    # -------------------------
    # Student ID
    # -------------------------

    df["Student_ID"] = (
        df["Student_ID"]
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )


    df = df.fillna("")

    return df


# =========================================================
# GET STUDENTS
# =========================================================

@app.route("/api/students")
def get_students():

    year = request.args.get(
        "year",
        "3rd Year"
    )

    if year not in YEAR_FILES:

        return jsonify({
            "success": False,
            "message": "Invalid year"
        }), 400


    csv_file = YEAR_FILES[year]


    try:

        if not os.path.exists(csv_file):
            return jsonify([])


        df = pd.read_csv(
            csv_file
        )

        df = prepare_dataframe(df)


        return jsonify(
            df.to_dict(
                orient="records"
            )
        )


    except Exception as e:

        print(
            "GET STUDENTS ERROR:",
            e
        )

        return jsonify({
            "success": False,
            "message": str(e)
        }), 500


# =========================================================
# ADD NEW STUDENT
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
                "message": "No student data received!"
            }), 400


        year = str(
            data.get(
                "Year",
                "3rd Year"
            )
        ).strip()


        if year not in YEAR_FILES:

            return jsonify({
                "success": False,
                "message": "Invalid year!"
            }), 400


        csv_file = YEAR_FILES[year]


        # -------------------------
        # Basic Details
        # -------------------------

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


        gender = str(
            data.get(
                "Gender",
                ""
            )
        ).strip()


        student_class = str(
            data.get(
                "Class",
                ""
            )
        ).strip()


        # -------------------------
        # Required Fields
        # -------------------------

        if student_id == "":

            return jsonify({
                "success": False,
                "message":
                "Student ID / Roll Number is required!"
            }), 400


        if name == "":

            return jsonify({
                "success": False,
                "message":
                "Student Name is required!"
            }), 400


        # =================================================
        # READ EXISTING DATA
        # =================================================

        if os.path.exists(csv_file):

            df = pd.read_csv(
                csv_file
            )

            df = prepare_dataframe(df)

        else:

            df = pd.DataFrame()


        # =================================================
        # DUPLICATE ID CHECK
        # =================================================

        if not df.empty:

            existing_ids = (
                df["Student_ID"]
                .astype(str)
                .str.strip()
            )

            if student_id in existing_ids.values:

                return jsonify({
                    "success": False,
                    "message":
                    f"Roll Number {student_id} already exists in {year}!"
                }), 400


        # =================================================
        # NUMBER FUNCTION
        # =================================================

        def number(value):

            try:

                return float(value)

            except:

                return 0


        osy = number(
            data.get("OSY", 0)
        )

        ste = number(
            data.get("STE", 0)
        )

        acn = number(
            data.get("ACN", 0)
        )

        dan = number(
            data.get("DAN", 0)
        )

        attendance = number(
            data.get("Attendance", 0)
        )


        # =================================================
        # VALIDATION
        # =================================================

        marks = [
            osy,
            ste,
            acn,
            dan
        ]


        for mark in marks:

            if mark < 0 or mark > 100:

                return jsonify({
                    "success": False,
                    "message":
                    "All marks must be between 0 and 100!"
                }), 400


        if attendance < 0 or attendance > 100:

            return jsonify({
                "success": False,
                "message":
                "Attendance must be between 0 and 100!"
            }), 400


        # =================================================
        # CALCULATE RESULT
        # =================================================

        total = (
            osy
            + ste
            + acn
            + dan
        )


        percentage = round(
            total / 4,
            2
        )


        if percentage >= 90:
            grade = "A+"

        elif percentage >= 80:
            grade = "A"

        elif percentage >= 70:
            grade = "B+"

        elif percentage >= 60:
            grade = "B"

        elif percentage >= 50:
            grade = "C"

        elif percentage >= 40:
            grade = "D"

        else:
            grade = "F"


        status = (
            "Eligible"
            if attendance >= 75
            else "Shortage"
        )


        # =================================================
        # NEW STUDENT
        # =================================================

        new_student = {

            "Student_ID": student_id,

            "Name": name,

            "Gender": gender,

            "Class": student_class,

            "OSY": osy,

            "STE": ste,

            "ACN": acn,

            "DAN": dan,

            "Total": total,

            "Percentage": percentage,

            "Attendance": attendance,

            "Status": status,

            "Grade": grade
        }


        new_df = pd.DataFrame(
            [new_student]
        )


        # =================================================
        # ADD DATA
        # =================================================

        if df.empty:

            final_df = new_df

        else:

            final_df = pd.concat(
                [
                    df,
                    new_df
                ],
                ignore_index=True
            )


        # =================================================
        # COLUMN ORDER
        # =================================================

        columns = [

            "Student_ID",
            "Name",
            "Gender",
            "Class",

            "OSY",
            "STE",
            "ACN",
            "DAN",

            "Total",
            "Percentage",

            "Attendance",
            "Status",
            "Grade"
        ]


        final_df = final_df[
            columns
        ]


        # =================================================
        # SAVE CSV
        # =================================================

        final_df.to_csv(
            csv_file,
            index=False
        )


        print(
            "STUDENT ADDED:",
            student_id,
            name,
            year
        )


        return jsonify({

            "success": True,

            "message":
            f"{name} added successfully to {year}!"

        })


    except Exception as e:

        print(
            "ADD STUDENT ERROR:",
            e
        )


        return jsonify({

            "success": False,

            "message":
            "Error while saving student: "
            + str(e)

        }), 500


# =========================================================
# UPLOAD CSV
# =========================================================

@app.route(
    "/api/upload_csv",
    methods=["POST"]
)
def upload_csv():

    try:

        year = request.form.get(
            "year",
            "3rd Year"
        )


        if year not in YEAR_FILES:

            return jsonify({
                "success": False,
                "message": "Invalid year!"
            }), 400


        if "file" not in request.files:

            return jsonify({
                "success": False,
                "message":
                "Please select a CSV file!"
            }), 400


        file = request.files["file"]


        if file.filename == "":

            return jsonify({
                "success": False,
                "message":
                "No file selected!"
            }), 400


        if not file.filename.lower().endswith(".csv"):

            return jsonify({
                "success": False,
                "message":
                "Only CSV files are allowed!"
            }), 400


        # =================================================
        # READ NEW CSV
        # =================================================

        df_new = pd.read_csv(
            file
        )

        df_new = prepare_dataframe(
            df_new
        )


        # =================================================
        # EXISTING CSV
        # =================================================

        csv_file = YEAR_FILES[year]


        if os.path.exists(csv_file):

            df_old = pd.read_csv(
                csv_file
            )

            df_old = prepare_dataframe(
                df_old
            )

        else:

            df_old = pd.DataFrame()


        # =================================================
        # DUPLICATE ID CHECK
        # =================================================

        old_ids = set()

        if not df_old.empty:

            old_ids = set(
                df_old["Student_ID"]
                .astype(str)
                .str.strip()
            )


        duplicate_ids = []


        for sid in df_new["Student_ID"]:

            sid = str(
                sid
            ).strip()


            if sid in old_ids:

                duplicate_ids.append(
                    sid
                )


        if duplicate_ids:

            return jsonify({

                "success": False,

                "message":
                "These Roll Numbers already exist: "
                + ", ".join(
                    duplicate_ids
                )

            }), 400


        # =================================================
        # CHECK DUPLICATES INSIDE UPLOADED FILE
        # =================================================

        duplicate_inside_file = (
            df_new["Student_ID"]
            .astype(str)
            .duplicated()
        )


        if duplicate_inside_file.any():

            duplicates = (
                df_new.loc[
                    duplicate_inside_file,
                    "Student_ID"
                ]
                .astype(str)
                .tolist()
            )


            return jsonify({

                "success": False,

                "message":
                "Duplicate Roll Numbers inside uploaded file: "
                + ", ".join(
                    duplicates
                )

            }), 400


        # =================================================
        # SAVE
        # =================================================

        if df_old.empty:

            final_df = df_new

        else:

            final_df = pd.concat(
                [
                    df_old,
                    df_new
                ],
                ignore_index=True
            )


        columns = [

            "Student_ID",
            "Name",
            "Gender",
            "Class",

            "OSY",
            "STE",
            "ACN",
            "DAN",

            "Total",
            "Percentage",

            "Attendance",
            "Status",
            "Grade"
        ]


        final_df = final_df[
            columns
        ]


        final_df.to_csv(
            csv_file,
            index=False
        )


        return jsonify({

            "success": True,

            "message":
            f"{len(df_new)} students uploaded successfully to {year}!"

        })


    except Exception as e:

        print(
            "UPLOAD ERROR:",
            e
        )


        return jsonify({

            "success": False,

            "message":
            str(e)

        }), 500


# =========================================================
# RUN SERVER
# =========================================================

if __name__ == "__main__":

    print()
    print("========================================")
    print(" STUDENT PERFORMANCE DASHBOARD")
    print("========================================")

    print(
        "1st Year :",
        YEAR_FILES["1st Year"]
    )

    print(
        "2nd Year :",
        YEAR_FILES["2nd Year"]
    )

    print(
        "3rd Year :",
        YEAR_FILES["3rd Year"]
    )

    print("========================================")
    print()

    # 0.0.0.0 allows other devices on same Wi-Fi
   # =========================================================
# RUN SERVER
# =========================================================

if __name__ == "__main__":

    print()
    print("========================================")
    print(" STUDENT PERFORMANCE DASHBOARD")
    print("========================================")

    print(
        "1st Year :",
        YEAR_FILES["1st Year"]
    )

    print(
        "2nd Year :",
        YEAR_FILES["2nd Year"]
    )

    print(
        "3rd Year :",
        YEAR_FILES["3rd Year"]
    )

    print("========================================")
    print()

    port = int(os.environ.get("PORT", 5000))

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
