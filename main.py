import openpyxl
from flask import Flask, render_template, request
from flask import send_file, redirect, url_for, session
import re
import os
import csv
import datetime

allowed_extensions = ['xlsx', 'XLSX']
upload_folder = "uploads/"
download_folder = "downloads/"
salarydate = 30042022

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


# regexp to clean unsafe strings
def clean_url(s):
    return(re.sub(r"[/\\?%*:|\"<>\x7F\x00-\x1F]", "-", s))


def check_file_extension(filename):
    return filename.split('.')[-1] in allowed_extensions


def check_positive(s):
    if s > 0:
        return("")
    else:
        return("-")


def validate_date(s):
    try:
        datetime.datetime.strptime(s, '%d%m%Y')
        return s
    except ValueError:
        raise ValueError("Feil datoformat, skal være DDMMYYYY, 28052022")


# Start point for application
@app.route('/')
def upload_file():
    return render_template('upload.html')


@app.route('/upload', methods=['GET', 'POST'])
def uploadfile():
    if request.method == 'POST':
        session['salarydate'] = validate_date(request.form['in_salarydate'])
        f = request.files['file']
        f.filename = clean_url(f.filename)
        if check_file_extension(f.filename):
            f.save(os.path.join(app.config['UPLOAD_FOLDER'],
                   f.filename))
            process_file(f.filename)
            session['my_filename'] = f.filename[:-4]+"HLT"
            return redirect(url_for('download'))
        else:
            return 'Sjekk at det ble lagt ved fil i xlsx format'


@app.route('/download')
def download():
    return send_file('downloads/go-'+session['my_filename'])


def process_file(file):
    inputfile = app.config['UPLOAD_FOLDER']+file
    wb = openpyxl.load_workbook(inputfile)
    sheet = wb.active

    # Get input spreadsheet into a list
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

    # Create the full column list including empty columns for export
    output_list = []
    for rows in input_list:
        debet = rows[2]
        credit = rows[3]
        o = []
        o.append(rows[0][:4].zfill(7))  # kontonr
        o.append((rows[1][:2] or "0").zfill(7))  # avdeling
        o.append("00000000")  # ukjent kolonne
        o.append("00000000")  # ukjent kolonne
        o.append("        ")  # tom
        o.append('')  # tom
        o.append(session['salarydate'])  # dato
        o.append("   ")  # filler
        o.append(session['salarydate'])  # dato
        o.append("0000000000")  # antall
        o.append("0000000000")  # sats
        o.append(str(round((debet - credit)*100)).zfill(10))  # beløp
        output_list.append(o)

    with open('downloads/go-'+file[:-4]
              + "HLT", 'w', newline="\r\n") as outfile:
        writer = csv.writer(outfile, delimiter=";")
        writer.writerows(output_list)


if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
