from flask import Flask, render_template, request, send_file
from datetime import datetime
import pandas as pd
app = Flask(__name__)

@app.route('/')
def upload_file1():
   return render_template('upload.html')

def modify_data(before_f, after_f):
  before = None
  after = None

  template = pd.read_csv('template.csv', encoding='utf-8', header=None)

  try:
    before = pd.read_html(before_f, encoding='utf-8')[0]
  except:
    before = pd.read_excel(before_f)

  try:
    after = pd.read_html(after_f, encoding='utf-8')[0]
  except:
    after = pd.read_excel(after_f)

  cond = after['sox_no'].isin(before['sox_no'])
  after.drop(after[cond].index, inplace = True)
  data = after.groupby('sox_no').first()
  data_list = data.reset_index().values.tolist()
  data_list
  for line in data_list:
    detail = ""
    if line[2].strip() == "Thai Post (REG)":
      detail = "R"
    if line[2].strip() == "Thai Post (EMS)":
      detail = "E"
    if detail == "":
      continue
    name = f"{line[3]} {line[0]}"
    tel = f"0{line[5]}"
    address = line[4].strip()[:-6]
    postal = line[4].strip()[-5:]
    current_row = {4: detail, 9: name, 10: tel, 11: address, 12: postal}
    template = pd.concat([template, pd.DataFrame(current_row, index=[0])])
  currentDay = datetime.now().day
  currentMonth = datetime.now().month
  filename = f"thaipost-{currentDay}-{currentMonth}.xlsx"
  template.to_excel(filename, index=False, header=None, engine="xlsxwriter")
  return filename

	
@app.route('/uploader', methods = ['GET', 'POST'])
def upload_file():
   if request.method == 'POST':
      before = request.files['before']
      after = request.files['after']
      filename = modify_data(before, after)
      return send_file(filename, as_attachment=True)
		
if __name__ == '__main__':
   app.run(debug = True)