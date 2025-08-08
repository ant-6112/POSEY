import json
import os
import time
import uuid
from datetime import datetime

import pandas as pd
import pyarrow as pa
from flask import jsonify, render_template, request
from werkzeug.utils import secure_filename

from app import app

CURRENT_USER = {"name": "John Doe", "avatar": "/static/img/avatar.png"}


@app.route("/")
def dashboard():
    return render_template("dashboard.html", user=CURRENT_USER)


@app.route("/datasets")
def datasets():
    return render_template("datasets.html", user=CURRENT_USER)


@app.route("/recent-runs")
def recent_runs():
    return render_template("recent_runs.html", user=CURRENT_USER)


@app.route("/results/<run_id>")
def results(run_id):
    return render_template("results.html", user=CURRENT_USER, run_id=run_id)


@app.route("/api/datasets", methods=["GET"])
def get_datasets():
    datasets = []
    if os.path.exists(app.config["UPLOAD_FOLDER"]):
        for filename in os.listdir(app.config["UPLOAD_FOLDER"]):
            if filename.endswith(".csv"):
                file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                df = pd.read_csv(file_path)
                table = pa.Table.from_pandas(df)
                data = table.to_pylist()
                stat = os.stat(file_path)

                # details = file_path.split("_")
                # id = file_path[0]
                # name = details[1]
                # uploaded_by = details[2]
                # upload_date = details[4]

                datasets.append(
                    {
                        "id": str(uuid.uuid4()),
                        "name": filename.replace(".csv", ""),
                        "uploaded_by": CURRENT_USER["name"],
                        "runs_count": 0,
                        "upload_date": datetime.fromtimestamp(
                            stat.st_mtime
                        ).isoformat(),
                        "data": data,
                    }
                )
            elif filename.endswith(".xlsx"):
                file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                df = pd.read_excel(file_path)
                table = pa.Table.from_pandas(df)
                data = table.to_pylist()

                stat = os.stat(file_path)
                datasets.append(
                    {
                        "id": str(uuid.uuid4()),
                        "name": filename.replace(".xlsx", ""),
                        "uploaded_by": CURRENT_USER["name"],
                        "runs_count": 0,
                        "upload_date": datetime.fromtimestamp(
                            stat.st_mtime
                        ).isoformat(),
                        "data": data,
                    }
                )
    return jsonify(datasets)


@app.route("/api/datasets/upload", methods=["POST"])
def upload_dataset():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if file and file.filename.endswith(".csv"):
        filename = (
            request.form.get("name") + ".csv"
            if request.form.get("name") != ""
            else secure_filename(file.filename)
        )

        id = str(uuid.uuid4())
        uploaded_by = CURRENT_USER["name"]
        runs_count = 0
        upload_date = datetime.now().isoformat()

        # filename = (
        #     filename
        #     + "_"
        #     + id
        #     + "_"
        #     + uploaded_by.replace(" ", "")
        #     + "_"
        #     + upload_date
        #     + ".csv"
        # )
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(file_path)

        return jsonify(
            {
                "id": id,
                "name": request.form.get("name")
                if request.form.get("name") != ""
                else filename.replace(".csv", ""),
                "uploaded_by": uploaded_by,
                "runs_count": 0,
                "upload_date": upload_date,
            }
        )
    elif file and file.filename.endswith(".xlsx"):
        filename = (
            request.form.get("name") + ".xlsx"
            if request.form.get("name") != ""
            else secure_filename(file.filename)
        )
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(file_path)

        return jsonify(
            {
                "id": str(uuid.uuid4()),
                "name": request.form.get("name")
                if request.form.get("name") != ""
                else filename.replace(".xlsx", ""),
                "uploaded_by": CURRENT_USER["name"],
                "runs_count": 0,
                "upload_date": datetime.now().isoformat(),
            }
        )

    return jsonify({"error": "Invalid file type"}), 400


@app.route("/api/runs", methods=["GET"])
def get_runs():
    runs = []
    if os.path.exists(app.config["RUNS_FOLDER"]):
        for filename in os.listdir(app.config["RUNS_FOLDER"]):
            if filename.endswith(".json"):
                with open(os.path.join(app.config["RUNS_FOLDER"], filename), "r") as f:
                    run_data = json.load(f)
                    runs.append(run_data)

    return jsonify(sorted(runs, key=lambda x: x["timestamp"], reverse=True))


def find_common_columns(dataset1, dataset2):
    columns1 = dataset1.columns
    columns2 = dataset2.columns

    common_columns = set(columns1).intersection(set(columns2))

    result = [[col, col, 1.0] for col in common_columns]

    return result


@app.route("/api/runs/createcde", methods=["POST"])
def createcde():
    try:
        data = request.get_json()

        dataset1 = pd.read_csv(
            os.path.join(app.config["UPLOAD_FOLDER"], data["dataset1"]) + ".csv"
        )
        dataset2 = pd.read_csv(
            os.path.join(app.config["UPLOAD_FOLDER"], data["dataset2"]) + ".csv"
        )

        common_columns = find_common_columns(dataset1, dataset2)

        run_data = {
            "dataset1": data.get("dataset1"),
            "dataset2": data.get("dataset2"),
            "matching_method": data.get("matching_method"),
            "num_cdes": data.get("num_cdes", 0),
            "timestamp": datetime.now().isoformat(),
            "common_columns": common_columns,
        }

        return jsonify(run_data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/runs/create", methods=["POST"])
def create_run():
    start = time.time()
    data = request.get_json()

    run_id = str(uuid.uuid4())

    dataset1 = pd.read_csv(
        os.path.join(app.config["UPLOAD_FOLDER"], data["dataset1"]) + ".csv"
    )
    dataset2 = pd.read_csv(
        os.path.join(app.config["UPLOAD_FOLDER"], data["dataset2"]) + ".csv"
    )

    table1 = pa.Table.from_pandas(dataset1)
    table2 = pa.Table.from_pandas(dataset2)
    data1 = table1.to_pylist()
    data2 = table2.to_pylist()
    merged_data = data1

    end = time.time()

    run_data = {
        "id": run_id,
        "dataset1": data.get("dataset1"),
        "dataset2": data.get("dataset2"),
        "matching_method": data.get("matching_method"),
        "num_cdes": data.get("num_cdes", 3),
        "timestamp": datetime.now().isoformat(),
        "data1": data1,
        "data2": data2,
        "merged": merged_data,
        "status": "Completed",
        "duration": f"{end - start:.2f} seconds",
        "selected": data.get("selected", []),
        "threshold": 0.1,
        "result_summary": {
            "total_records": 0,
            "total_anamolies": 0,
            "% of anaomlies": 0,
        },
    }

    # Save run data
    with open(os.path.join(app.config["RUNS_FOLDER"], f"{run_id}.json"), "w") as f:
        json.dump(run_data, f, indent=2)

    return jsonify(run_data)


@app.route("/api/runs/<run_id>", methods=["GET"])
def get_run(run_id):
    try:
        with open(os.path.join(app.config["RUNS_FOLDER"], f"{run_id}.json"), "r") as f:
            run_data = json.load(f)
        return jsonify(run_data)
    except FileNotFoundError:
        return jsonify({"error": "Run not found"}), 404


@app.route("/api/runs/<run_id>/results", methods=["GET"])
def get_run_results(run_id):
    results = {
        "summary": {
            "total_records": 10000,
            "matched_records": 9856,
            "unmatched_records": 144,
            "match_rate": 98.56,
            "execution_time": "2m 34s",
        },
        "metrics": [
            {"name": "Match Rate", "value": 98.56, "unit": "%"},
            {"name": "Total Records", "value": 10000, "unit": ""},
            {"name": "Processing Time", "value": 154, "unit": "seconds"},
            {"name": "Memory Usage", "value": 245, "unit": "MB"},
        ],
        "chart_data": {
            "match_distribution": [
                {"category": "Perfect Match", "count": 8500},
                {"category": "Fuzzy Match", "count": 1356},
                {"category": "No Match", "count": 144},
            ],
            "time_series": [
                {"timestamp": "2025-01-01", "matches": 856},
                {"timestamp": "2025-01-02", "matches": 923},
                {"timestamp": "2025-01-03", "matches": 1045},
                {"timestamp": "2025-01-04", "matches": 987},
                {"timestamp": "2025-01-05", "matches": 1123},
            ],
        },
    }
    return jsonify(results)
