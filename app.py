import os
import uuid
import datetime
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

from openpyxl import Workbook
from openpyxl.drawing.image import Image

import barcode
from barcode.writer import ImageWriter

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "API SERVER RUNNING"

@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/create_excel", methods=["POST"])
def create_excel():
    try:
        data = request.json
        mode = data.get("mode", "normal")

        name = data.get("name", "")
        exp = data.get("exp", "")
        mfg = data.get("mfg", "")
        lot = data.get("lot", "")
        qty = data.get("qty", "")
        count = int(data.get("barcode_qty") or 1)

        wb = Workbook()
        ws = wb.active

        row = 1
        temp_files = []

        for i in range(count):
            code = datetime.datetime.now().strftime("%Y%m%d") + f"{i:04d}"

            if mode == "lot":
                ws[f"A{row}"] = "품명"
                ws[f"B{row}"] = name

                ws[f"A{row+1}"] = f"소비기한\n{exp}"
                ws[f"B{row+1}"] = f"제조일자\n{mfg}"

                ws[f"A{row+2}"] = "수량"
                ws[f"B{row+2}"] = qty

                ws[f"A{row+3}"] = "바코드"
                ws[f"B{row+3}"] = lot
            else:
                ws[f"A{row}"] = "품명"
                ws[f"B{row}"] = name

                ws[f"A{row+1}"] = "소비기한"
                ws[f"B{row+1}"] = exp

                ws[f"A{row+2}"] = "수량"
                ws[f"B{row+2}"] = qty

                ws[f"A{row+3}"] = "바코드"

            # 바코드 생성
            filename = f"{uuid.uuid4().hex}"
            barcode_class = barcode.get_barcode_class("code128")
            barcode_obj = barcode_class(code, writer=ImageWriter())
            barcode_obj.save(filename)

            img = Image(f"{filename}.png")
            ws.add_image(img, f"B{row+3}")

            temp_files.append(f"{filename}.png")

            row += 4

        file = f"{uuid.uuid4().hex}.xlsx"
        wb.save(file)

        for f in temp_files:
            try:
                os.remove(f)
            except:
                pass

        return send_file(file, as_attachment=True, download_name="barcode.xlsx")

    except Exception as e:
        return jsonify({"error": str(e)}), 500
