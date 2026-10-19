
"""Local web dashboard for Smart CSV & Excel Cleaner."""

from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

from flask import (
    Flask,
    render_template_string,
    request,
    send_file,
)

from cleaner import clean_file


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, initial-scale=1">
    <title>Smart CSV & Excel Cleaner</title>
    <style>
        * { box-sizing: border-box; }

        body {
            font-family: Arial, sans-serif;
            background: #f2f5fa;
            margin: 0;
            color: #172b4d;
        }

        .container {
            max-width: 650px;
            margin: 75px auto;
            padding: 20px;
        }

        .card {
            background: white;
            padding: 40px;
            border-radius: 16px;
            box-shadow: 0 10px 35px #172b4d16;
        }

        h1 {
            font-size: 30px;
            margin-bottom: 10px;
        }

        .subtitle {
            color: #64748b;
            line-height: 1.7;
        }

        .upload {
            padding: 35px 20px;
            margin: 30px 0 20px;
            border: 2px dashed #90a4c4;
            border-radius: 12px;
            background: #f8faff;
            text-align: center;
        }

        input[type=file] {
            width: 100%;
            max-width: 360px;
        }

        button {
            width: 100%;
            background: #2563eb;
            color: white;
            border: 0;
            border-radius: 9px;
            padding: 15px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
        }

        button:hover {
            background: #1d4ed8;
        }

        .error {
            background: #fff0f0;
            color: #b42318;
            padding: 14px;
            border-radius: 8px;
            margin-bottom: 18px;
        }

        .note {
            margin-top: 22px;
            font-size: 13px;
            color: #64748b;
            line-height: 1.6;
        }

        .footer {
            text-align: center;
            font-size: 12px;
            color: #94a3b8;
            margin-top: 20px;
        }
    </style>
</head>

<body>
<div class="container">
    <div class="card">
        <h1>Smart CSV & Excel Cleaner</h1>

        <p class="subtitle">
            Upload a contact list and automatically
            clean duplicates, standardize formatting,
            and generate a data quality report.
        </p>

        {% if error %}
            <div class="error">{{ error }}</div>
        {% endif %}

        <form method="POST"
              action="/clean"
              enctype="multipart/form-data">

            <div class="upload">
                <p><strong>Select your file</strong></p>

                <input
                    type="file"
                    name="file"
                    accept=".csv,.xlsx"
                    required
                >

                <p class="subtitle">
                    CSV or XLSX | Max 10 MB
                </p>
            </div>

            <button type="submit">
                Clean Data & Download Excel
            </button>
        </form>

        <p class="note">
            Required columns: Name, Email, Phone, City.
            Email and phone checks validate format only;
            they do not verify that contacts are active.
        </p>
    </div>

    <p class="footer">
        Python + Pandas + Flask Automation Demo
    </p>
</div>
</body>
</html>
"""


@app.get("/")
def home():
    return render_template_string(HTML, error=None)


@app.post("/clean")
def clean_upload():
    uploaded = request.files.get("file")

    if not uploaded or not uploaded.filename:
        return render_template_string(
            HTML,
            error="Please select a file."
        ), 400

    extension = Path(uploaded.filename).suffix.lower()

    if extension not in (".csv", ".xlsx"):
        return render_template_string(
            HTML,
            error="Only CSV and XLSX files are supported."
        ), 400

    try:
        with TemporaryDirectory() as temp_folder:
            temp_path = Path(temp_folder)

            source = temp_path / ("uploaded" + extension)
            destination = temp_path / "cleaned_data.xlsx"

            uploaded.save(source)

            clean_file(source, destination)

            # Read workbook into memory before deleting
            # temporary files.
            workbook_bytes = destination.read_bytes()

        return send_file(
            BytesIO(workbook_bytes),
            as_attachment=True,
            download_name="cleaned_data.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    except (ValueError, PermissionError) as exc:
        return render_template_string(
            HTML,
            error=str(exc)
        ), 400

    except Exception:
        app.logger.exception("File processing failed")
        return render_template_string(
            HTML,
            error="Unable to process this file. "
                  "Check its format and contents."
        ), 400


@app.errorhandler(413)
def file_too_large(error):
    return render_template_string(
        HTML,
        error="File is too large. Maximum size is 10 MB."
    ), 413


if __name__ == "__main__":
    app.run(
        host="127.0.0.1",
        port=5001,
        debug=False,
    )
