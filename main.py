from flask import Flask, request, send_from_directory, render_template

import pandas as pd
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
  os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(DOWNLOAD_FOLDER):
  os.makedirs(DOWNLOAD_FOLDER)


def allowed_file(filename):
  return '.' in filename and \
         filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def process_resume(file_path):
  df = pd.read_excel(file_path)
  processed_df = pd.DataFrame(
      columns=['Country', 'Language', 'Company', 'Date', 'Role'])
  months = [
      'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct',
      'Nov', 'Dec'
  ]

  for index, resume_text in df.iterrows():
    if pd.isna(resume_text['略歴']):
      break
    lines = resume_text['略歴'].split("\n")
    companies, dates, roles = [], [], []
    language = None
    country = None

    for line in lines:
      if line.startswith("Language:"):
        language = line.split("Language:")[1].strip()
        continue
      if line.startswith("@"):
        country = line[1:].strip()
        continue
      if (line and line[0].isdigit() and "　" in line) or (any(
          month.startswith(line[:3]) for month in months) and "　" in line):
        date, rest = line.split("　", 1)
        if " / " in rest:
          company, role = rest.split(" / ", 1)
          dates.append(date.strip())
          companies.append(company.strip())
          roles.append(role.strip())

    temp_df = pd.DataFrame({
        'Country': [country],
        'Language': [language],
        'Company': ["\n".join(companies)],
        'Date': ["\n".join(dates)],
        'Role': ["\n".join(roles)]
    })
    processed_df = pd.concat([processed_df, temp_df], ignore_index=True)

  output_path = os.path.join(app.config['DOWNLOAD_FOLDER'],
                             'processed_output.xlsx')
  processed_df.to_excel(output_path, index=False)
  return output_path


@app.route('/upload', methods=['POST'])
def upload_file():
  if 'file' not in request.files:
    return 'No file part'
  file = request.files['file']
  if file.filename == '':
    return 'No selected file'
  if file and allowed_file(file.filename):
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    processed_file_path = process_resume(file_path)
    return send_from_directory(os.path.dirname(processed_file_path),
                               os.path.basename(processed_file_path),
                               as_attachment=True)
  return 'Invalid file type'


@app.route('/')
def index():
  return render_template('upload_form.html')


if __name__ == '__main__':
  app.run(host="0.0.0.0", port=8080)
