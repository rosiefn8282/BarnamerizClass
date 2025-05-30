
from flask import Flask, render_template, request, send_file
import pandas as pd
import uuid
import io

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    table_html = None
    excel_data = None
    filename = None
    if request.method == "POST":
        file = request.files["file"]
        if file.filename.endswith((".xls", ".xlsx")):
            df = pd.read_excel(file)

            # بررسی وجود ستون استاد
            if "استاد" not in df.columns:
                df["استاد"] = "تعریف‌نشده"

            table_html = df.to_html(classes='data', header=True, index=False)

            # ذخیره در حافظه برای اکسل
            output = io.BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)
            excel_data = output.read()
            filename = f"{uuid.uuid4().hex}.xlsx"

            with open(filename, "wb") as f:
                f.write(excel_data)

    return render_template("index.html", table=table_html, filename=filename)

@app.route("/download-excel/<filename>")
def download_excel(filename):
    return send_file(filename, as_attachment=True)

@app.route("/download-pdf", methods=["POST"])
def download_pdf():
    from xhtml2pdf import pisa
    html = request.form["html"]
    pdf_file = "output.pdf"
    with open(pdf_file, "w+b") as f:
        pisa.CreatePDF(html, dest=f)
    return send_file(pdf_file, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
