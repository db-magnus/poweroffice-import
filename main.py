import openpyxl
import pprint
import string
from flask import Flask, render_template, request
from flask import send_file, redirect, url_for, session
from werkzeug.utils import secure_filename
import os

allowed_extensions = ['xlsx', 'XLSX']
upload_folder = "uploads/"
download_folder = "downloads/"
salarydate = 27052022

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = upload_folder
app.config['DOWNLOAD_FOLDER'] = download_folder
app.config['MAX_CONTENT_LENGTH'] = 2 * 1024 * 1024
app.secret_key = "laskdfj alsdkfj lskdf"


if not os.path.exists(upload_folder):
    os.mkdir(upload_folder)
if not os.path.exists(download_folder):
    os.mkdir(download_folder)


# function to remove None from text arrays
def xstr(s):
    if s is None:
        return ''
    else:
        return s


# function to convert to floats.
def xflo(s):
    if s is None:
        return 0.0
    else:
        return float(s)


def check_file_extension(filename):
    return filename.split('.')[-1] in allowed_extensions


# Start point for application
@app.route('/')
def upload_file():
    return render_template('upload.html')


@app.route('/upload', methods=['GET', 'POST'])
def uploadfile():
    if request.method == 'POST':
        f = request.files['file']
        if check_file_extension(f.filename):
            f.save(os.path.join(app.config['UPLOAD_FOLDER'],
                   secure_filename(f.filename)))
            process_file(f.filename)
            session['my_filename'] = f.filename
            # return 'file uploaded successfully, now processing
            return redirect(url_for('download'))
        else:
            return 'The file extension is not allowed'


@app.route('/download')
def download():
    return send_file('downloads/go-'+session['my_filename'])


def process_file(file):
    inputfile = app.config['UPLOAD_FOLDER']+file
    wb = openpyxl.load_workbook(inputfile)
    sheet = wb.active

    input_list = []
    for i, row in enumerate(sheet.iter_rows(values_only=True)):
        if i == 0:
            continue
        else:
            if(row[0].startswith('Sum')):
                continue
            li = []
            li.append(xstr(row[1]))
            li.append(xstr(row[2]))
            li.append(xflo(row[4]))
            li.append(xflo(row[5]))
            input_list.append(li[:])

    # pprint.pprint(input_list)

    output_list = []
    for rows in input_list:
        debet = rows[2]
        credit = rows[3]
        o = []
        o.append(rows[0][:4])  # kontonr
        o.append(rows[1][:2])  # avdeling
        o.append(0)  # ukjent kolonne
        o.append(0)  # ukjent kolonne
        o.append('')  # tom
        o.append('')  # tom
        o.append(salarydate)  # dato
        o.append('')  # tom
        o.append(salarydate)  # dato
        o.append(debet)  # debet
        o.append(0)  # ukjent kolonne
        o.append(debet - credit)
        output_list.append(o)

    # print("Output list")
    # pprint.pprint(output_list)

    exp_workbook = openpyxl.Workbook()
    exp_sheet = exp_workbook.active
    exp_sheet.title = "Eksportfil fra Visma"

    columns = []
    for idx, let in enumerate(string.ascii_uppercase):
        if idx > 11:
            break
        columns.append(let)

    for idx, inner in enumerate(output_list):
        for idy, cell in enumerate(columns):
            exp_sheet[cell+str(idx+1)] = inner[idy]
            # print("cell and idx: ", cell+str(idx))
            # print("inner: ", inner[idy])

    exp_workbook.save('downloads/go-'+file)


if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
