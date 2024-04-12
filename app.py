from flask import Flask, render_template, request, Response
from datetime import datetime
import openpyxl, io, os
from courses2ics import get_cInfos, io_write

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")

@app.route("/schedulegenerator")
def schedulegenerator():
    return render_template("courses2schedule.html")

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return render_template("courses2schedule.html", error="No file selected")
    file = request.files["file"]
    if file.filename == "":
        return render_template("courses2schedule.html", error="No file selected")
    # check if the suffix is xlsx
    if file.filename[-5:] != ".xlsx":
        return render_template("courses2schedule.html", error="Invalid file format")
    # Go to @app.route convert
    return convert(file)

@app.route("/error")
def error(e):
    return render_template("error.html", error_message=e)

@app.route("/convert", methods=["POST"])
def convert(IOfile: io.BytesIO):
    
    try:
        # read the file
        sheet = openpyxl.load_workbook(io.BytesIO(IOfile.read()))["View My Courses"]
        # process data
        classes = get_cInfos(sheet)
        # generate the ics file in memory
        outputfile = io.StringIO()
        io_write(outputfile, classes)
    except Exception as e:
        import traceback
        traceback.print_exc()
        # return error(e)
        return traceback.format_exc()
    # send the ics file to the user
    response = Response(outputfile.getvalue(), mimetype="text/calendar")
    response.headers["Content-Disposition"] = "attachment; filename=courses.ics"
    return response

    """
    return download(response)

@app.route("/download", methods=["POST"])
def download(response):


    """
    


if __name__ == "__main__":
    app.run(host="::", port=15646, debug=True)
