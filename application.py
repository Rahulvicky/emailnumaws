import docx2txt
from flask import Flask, render_template, request, send_file
import re
import glob
import shutil
import xlwt
import os
from io import StringIO
from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer3.converter import TextConverter
from pdfminer3.layout import LAParams
from pdfminer3.pdfpage import PDFPage
from werkzeug.utils import secure_filename
from docx2pdf import convert
email_regex = re.compile(r"[\w\.-]+@[\w\.-]+")
phone_num = re.compile(r'[6-9]{1}[0-9]{9}')


application = app = Flask(__name__)

app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

ALLOWED_EXTENSIONS = set(['docx', 'pdf'])


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')


@app.route('/uploader', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        emailid = []
        mobile = []
        location = []
        di = {}
        pathnew = os.getcwd()
        # file delete
        UPLOAD_FOLDERNEW = os.path.join(pathnew, 'uploads')
        if os.path.isdir(UPLOAD_FOLDERNEW):
            print("hai dir")
            shutil.rmtree(UPLOAD_FOLDERNEW)

        UPLOAD_FILENEW = os.path.join(pathnew, 'document.xls')
        if os.path.isfile(UPLOAD_FILENEW):
            print("hai file")
            os.remove(UPLOAD_FILENEW)

        # Get current path
        path = os.getcwd()
        # file Upload
        UPLOAD_FOLDER = os.path.join(path, 'uploads')
        if not os.path.isdir(UPLOAD_FOLDER):
            print("naya dir bana")
            os.mkdir(UPLOAD_FOLDER)

        app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

        f = request.files.getlist('files[]')
        for file in f:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file.save(os.path.join(UPLOAD_FOLDER, filename))

    text1 = ""
    rawtext = ''
    pathtest2 = UPLOAD_FOLDER
    pathtest2 += "/*"

    # currentdir = os.getcwd()
    # currentdir += "/uploads/"
    # convert(currentdir)

    for file in glob.glob(pathtest2):
        if file.endswith(".pdf"):
            pagenums = set()
            output = StringIO()
            manager = PDFResourceManager()
            converter = TextConverter(manager, output, laparams=LAParams())
            interpreter = PDFPageInterpreter(manager, converter)
            infile = open(file, 'rb')
            for page in PDFPage.get_pages(infile, pagenums):
                interpreter.process_page(page)
            infile.close()
            converter.close()
            text = output.getvalue()
            output.close()
            text1 += text
            rawtext += text1
            location.append(os.path.abspath(file))
            emailid.append(email_regex.findall(text1))
            mobile.append(phone_num.findall(text1))
            key = str(email_regex.findall(text1))
            di[key] = phone_num.findall(text1)
            text1 = ""
        elif file.endswith(".docx"):
            text1 = docx2txt.process(file)
            rawtext += text1
            location.append(os.path.abspath(file))
            emailid.append(email_regex.findall(text1))
            mobile.append(phone_num.findall(text1))
            key = str(email_regex.findall(text1))
            di[key] = phone_num.findall(text1)
            text1 = ""




    resultsemail = email_regex.findall(rawtext)
    resultsmob = phone_num.findall(rawtext)
    num_of_results = len(emailid)
    # excel data
    workbook = xlwt.Workbook()
    sheet1 = workbook.add_sheet("My First Sheet")
    row = 1
    sheet1.write(0, 0, "EMAIL-ID")
    sheet1.write(0, 1, "Mobile Number")
    sheet1.write(0, 2, "Location of File")
    while row < len(emailid)+1:
        column = 0
        sheet1.write(row, column, emailid[row-1])
        column = column + 1
        sheet1.write(row, column, mobile[row-1])
        column = column + 1
        sheet1.write(row, column, location[row-1])
        row = row + 1
    workbook.save("document.xls")
    return render_template("index.html", resultsemail=di, resultsmob=resultsmob, num_of_results=num_of_results)


@app.route("/download", methods=['GET', 'POST'])
def download_file():
    file_name = 'document.xls'
    return send_file(file_name, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)