from flask import Flask, render_template, request
import openpyxl # type: ignore
import os

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/nota')
def nota():
    return render_template('nota.html')

@app.route('/peperiksaan')
def peperiksaan():
    return render_template('peperiksaan.html')

@app.route('/Feedback')
def hubungi():
    return render_template('hubungi.html')

@app.route('/hasil', methods=['POST'])
def hasil():
    nama = request.form['nama']
    kelas = request.form['kelas']
    mesej = request.form['mesej']

    filename = "data_pelawat.xlsx"

    if os.path.exists(filename):
        workbook = openpyxl.load_workbook(filename)
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Nama", "Kelas", "Mesej"])  # tajuk column

    sheet.append([nama, kelas, mesej])
    workbook.save(filename)

    return render_template('hasil.html', nama=nama, kelas=kelas, mesej=mesej)

if __name__ == '__main__':
    app.run(debug=True)
