from flask import Flask, render_template, jsonify, request
import pandas as pd
import os

app = Flask(__name__)

# --------------------------------------------------
# DATA FILES
# --------------------------------------------------

DATA_FOLDER = "data"

YEAR_FILES = {
    "1st Year": "first_year.xlsx",
    "2nd Year": "second_year.xlsx",
    "3rd Year": "third_year.xlsx"
}


# --------------------------------------------------
# LOAD STUDENT DATA
# --------------------------------------------------

def load_students(year):

    filename = YEAR_FILES.get(year)

    if not filename:
        return pd.DataFrame()

    filepath = os.path.join(DATA_FOLDER, filename)

    if not os.path.exists(filepath):
        return pd.DataFrame()

    try:
        df = pd.read_excel(filepath)

        df.columns = df.columns.astype(str).str.strip()

        # Calculate Total
        subjects = ["OSY", "STE", "ACN", "DAN"]

        existing_subjects = [
            col for col in subjects if col in df.columns
        ]

        if existing_subjects:
            df["Total"] = df[existing_subjects].sum(axis=1)

            df["Percentage"] = (
                df["Total"] /
                (len(existing_subjects) * 100)
            ) * 100

        # Attendance
        if "Attendance" in df.columns:

            df["Attendance"] = pd.to_numeric(
                df["Attendance"],
                errors="coerce"
            ).fillna(0)

            df["Attendance Status"] = df["Attendance"].apply(
                lambda x: "Good" if x >= 75 else "Low"
            )

        # Grade
        if "Percentage" in df.columns:

            def get_grade(p):

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

            df["Grade"] = df["Percentage"].apply(get_grade)

        df = df.fillna("")

        return df

    except Exception as e:

        print("ERROR:", e)

        return pd.DataFrame()


# --------------------------------------------------
# HOME
# --------------------------------------------------

@app.route("/")
def home():

    return render_template("index.html")


# --------------------------------------------------
# STUDENTS API
# --------------------------------------------------

@app.route("/api/students")
def students():

    year = request.args.get("year", "1st Year")

    df = load_students(year)

    return jsonify(df.to_dict(orient="records"))


# --------------------------------------------------
# STUDENT SEARCH
# --------------------------------------------------

@app.route("/api/search")
def search_student():

    year = request.args.get("year", "1st Year")
    query = request.args.get("q", "").strip().lower()

    df = load_students(year)

    if df.empty:
        return jsonify([])

    if query == "":
        return jsonify(df.to_dict(orient="records"))

    result = df[
        df.astype(str)
        .apply(
            lambda row:
            row.str.lower().str.contains(
                query,
                na=False
            ).any(),
            axis=1
        )
    ]

    return jsonify(result.to_dict(orient="records"))


# --------------------------------------------------
# DASHBOARD STATISTICS
# --------------------------------------------------

@app.route("/api/stats")
def statistics():

    year = request.args.get("year", "1st Year")

    df = load_students(year)

    if df.empty:

        return jsonify({
            "total_students": 0,
            "average_percentage": 0,
            "top_performer": "-",
            "average_attendance": 0
        })

    total_students = len(df)

    average_percentage = 0

    if "Percentage" in df.columns:
        average_percentage = round(
            pd.to_numeric(
                df["Percentage"],
                errors="coerce"
            ).mean(),
            2
        )

    top_performer = "-"

    if (
        "Percentage" in df.columns
        and "Name" in df.columns
        and len(df) > 0
    ):

        temp = df.copy()

        temp["Percentage"] = pd.to_numeric(
            temp["Percentage"],
            errors="coerce"
        )

        top = temp.loc[
            temp["Percentage"].idxmax()
        ]

        top_performer = str(top["Name"])

    average_attendance = 0

    if "Attendance" in df.columns:

        average_attendance = round(
            pd.to_numeric(
                df["Attendance"],
                errors="coerce"
            ).mean(),
            2
        )

    return jsonify({
        "total_students": total_students,
        "average_percentage": average_percentage,
        "top_performer": top_performer,
        "average_attendance": average_attendance
    })


# --------------------------------------------------
# RUN
# --------------------------------------------------

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
