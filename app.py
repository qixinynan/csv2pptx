import os
import random

from flask import Flask, request, send_file
from flask import render_template
import uuid
import csv2pptx

app = Flask(__name__)
app.config['DEBUG'] = True
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), './uploads')


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/download')
def download():
    filename = request.values.get('file')
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename))


@app.post('/generate')
def generate():
    csv_file = request.files.get('csv')
    count = int(request.values.get('count'))
    start_row = int(request.values.get('start_row'))
    start_column = int(request.values.get('start_column'))

    if csv_file is None:
        return "缺少需要上传的文件"
    save_csv_file_name = str(uuid.uuid4()) + '.csv'
    csv_path = os.path.join(app.config['UPLOAD_FOLDER'], save_csv_file_name)
    csv_file.save(csv_path)

    data = csv2pptx.read_csv(csv_path, start_row, start_column)
    samples = random.sample(data, count)
    presentation = csv2pptx.generate_presentation_by_array(samples)

    pptx_filename = str(uuid.uuid4()) + '.pptx'
    pptx_path = os.path.join(app.config['UPLOAD_FOLDER'], pptx_filename)
    presentation.save(pptx_path)
    print(samples)
    download_link = "/download?file=" + pptx_filename
    return render_template('generate.html', samples=samples, download=download_link)


if __name__ == '__main__':
    app.run()
