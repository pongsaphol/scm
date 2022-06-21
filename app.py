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
  R = 0
  E = 0
  for line in data_list:
    detail = ""
    if line[1] == None:
      continue
    if line[1].strip() == "Thai Post (REG)":
      detail = "R"
      R += 1
    if line[1].strip() == "Thai Post (EMS)":
      detail = "E"
      E += 1
    if detail == "":
      continue
    name = f"{line[2]} {line[0]}"
    tel = f"0{line[4]}"
    address = line[3].strip()[:-6]
    postal = line[3].strip()[-5:]
    current_row = {4: detail, 9: name, 10: tel, 11: address, 12: postal}
    template = pd.concat([template, pd.DataFrame(current_row, index=[0])])
  currentDay = datetime.now().day
  currentMonth = datetime.now().month
  filename = f"{currentDay}-{currentMonth}-R{R}-E{E}-T{R + E}.xlsx"
  template.to_excel(filename, index=False, header=None, engine="xlsxwriter")
  return filename

def modify_data2(before_f, after_f):
  before = None
  after = None

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
  after = after.sort_values(by=['sox_no']).reset_index()
  df = after.drop(['so_no'], axis=1)
  tmp = df.values.tolist()

  ans = []
  for i in range(len(tmp)):
    if i == 0:
      ans.append(1)
    elif tmp[i][1] == tmp[i-1][1]:
      ans.append(ans[i-1])
    else:
      ans.append(ans[i-1]+1)


  currentDay = datetime.now().day
  currentMonth = datetime.now().month
  filename = f"sox-{currentDay}-{currentMonth}-T{ans[-1]}.xlsx"
  writer = pd.ExcelWriter(filename, engine="xlsxwriter")

  df.to_excel(writer, sheet_name='Sheet1', index=False)

  workbook  = writer.book
  worksheet = writer.sheets['Sheet1']
  worksheet.set_column(1, 1, 12)
  worksheet.set_column(3, 3, 18.8)
  worksheet.set_column(4, 4, 60)

  center_format = workbook.add_format()

  center_format.set_align('center')


  green = workbook.add_format({'bg_color': '#C6EFCE'})

  for i in range(len(ans)):
    num = i + 1
    worksheet.write(num, 0, ans[i], center_format)
    if ans[i] % 2 == 1:
      file = f'=MOD($A${num + 1},2)=1'
      worksheet.conditional_format(num, 0, num, 5, {'type': 'formula', 'criteria': file, 'format': green})

  writer.save()
  return filename
  
	
@app.route('/uploader', methods = ['GET', 'POST'])
def upload_file():
   if request.method == 'POST':
      before = request.files['before']
      after = request.files['after']
      filename = modify_data(before, after)
      return send_file(filename, as_attachment=True)

@app.route('/uploader2', methods = ['GET', 'POST'])
def upload_file2():
   if request.method == 'POST':
      before = request.files['before']
      after = request.files['after']
      filename = modify_data2(before, after)
      return send_file(filename, as_attachment=True)
		
if __name__ == '__main__':
   app.run()