import os
import traceback
import pandas as pd
import requests
from flask import Flask, request, render_template, send_file, jsonify
from PIL import Image
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage

app = Flask(__name__, template_folder='templates')
logs = []

def log(msg):
    logs.append(msg)
    print(msg)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_excel():
    try:
        file = request.files['file']
        log(f"[上传文件] {file.filename}")

        df = pd.read_excel(file)
        output_path = 'output.xlsx'
        df.to_excel(output_path, index=False)

        wb = load_workbook(output_path)
        ws = wb.active

        for col in range(2, df.shape[1] + 1):
            for row in range(2, df.shape[0] + 2):
                cell = ws.cell(row=row, column=col)
                url = cell.value
                if isinstance(url, str) and url.startswith("http"):
                    try:
                        img_bytes = requests.get(url, timeout=10).content
                        image = Image.open(BytesIO(img_bytes))
                        image = image.resize((245, 342))
                        img_io = BytesIO()
                        image.save(img_io, format='PNG')
                        img_io.seek(0)

                        excel_img = ExcelImage(img_io)
                        cell_coordinate = cell.coordinate
                        excel_img.anchor = cell_coordinate
                        ws.add_image(excel_img)
                        log(f"[图片成功] {url} -> {cell_coordinate}")
                    except Exception as e:
                        log(f"[图片失败] {url}，错误: {e}")

        wb.save(output_path)
        return send_file(output_path, as_attachment=True)

    except Exception as err:
        log(f"[服务处理失败] {str(err)}")
        log(traceback.format_exc())
        return "服务异常：" + str(err), 500

@app.route('/logs', methods=['GET'])
def get_logs():
    return jsonify(logs)