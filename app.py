import os
import uuid
import datetime
from flask import Flask, request, send_file, jsonify, make_response, render_template
from flask_cors import CORS

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image

import barcode
from barcode.writer import ImageWriter

app = Flask(__name__, template_folder="templates")
CORS(app)

@app.after_request
def after_request(response):
    response.headers.add("Access-Control-Allow-Origin", "*")
    return response


# ✅ UI 페이지
@app.route("/")
def index():
    return render_template("index.html")


# ✅ API
@app.route("/create_excel", methods=["POST"])
def create_excel():
    try:
        data = request.json
        mode = data.get("mode", "normal")

        if mode == "lot":
            return create_lot_excel(data)
        else:
            return create_normal_excel(data)

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# =========================
# 일반 모드
# =========================
def create_normal_excel(data):
    name = data.get("name", "")
    exp = data.get("exp", "")
    qty = data.get("qty", "")
    count = int(data.get("barcode_qty", 1))

    wb = Workbook()
    ws = wb.active

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 120

    row = 1
    temp_files = []

    for i in range(count):
        code = datetime.datetime.now().strftime("%Y%m%d") + f"{i:04d}"

        ws[f"A{row}"] = "품명"
        ws[f"B{row}"] = name

        ws[f"A{row+1}"] = "소비기한"
        ws[f"B{row+1}"] = exp

        ws[f"A{row+2}"] = "수량"
        ws[f"B{row+2}"] = qty

        ws[f"A{row+3}"] = "바코드"

        # 바코드 생성
        fname = f"{uuid.uuid4().hex}.png"
        temp_files.append(fname)

        barcode.get("code128", code, writer=ImageWriter()).write(open(fname, "wb"))
        img = Image(fname)
        ws.add_image(img, f"B{row+3}")

        row += 4

    file = f"{uuid.uuid4().hex}.xlsx"
    wb.save(file)

    for f in temp_files:
        os.remove(f)

    return send_file(file, as_attachment=True, download_name="barcode.xlsx")


# =========================
# 로트 모드
# =========================
def create_lot_excel(data):
    name = data.get("name", "")
    exp = data.get("exp", "")
    mfg = data.get("mfg", "")
    lot = data.get("lot", "")
    qty = data.get("qty", "")
    count = int(data.get("barcode_qty", 1))

    wb = Workbook()
    ws = wb.active

    ws.column_dimensions["A"].width = 80
    ws.column_dimensions["B"].width = 80

    row = 1
    temp_files = []

    for i in range(count):
        code = datetime.datetime.now().strftime("%Y%m%d") + f"{i:04d}"

        ws[f"A{row}"] = "품명"
        ws[f"B{row}"] = name

        ws[f"A{row+1}"] = f"소비기한\n{exp}"
        ws[f"B{row+1}"] = f"제조일자\n{mfg}"

        ws[f"A{row+2}"] = "수량"
        ws[f"B{row+2}"] = qty

        ws[f"A{row+3}"] = "바코드"
        ws[f"B{row+3}"] = lot

        # 바코드 생성
        fname = f"{uuid.uuid4().hex}.png"
        temp_files.append(fname)

        barcode.get("code128", code, writer=ImageWriter()).write(open(fname, "wb"))
        img = Image(fname)
        ws.add_image(img, f"A{row+3}")

        row += 4

    file = f"{uuid.uuid4().hex}.xlsx"
    wb.save(file)

    for f in temp_files:
        os.remove(f)

    return send_file(file, as_attachment=True, download_name="lot_barcode.xlsx")


if __name__ == "__main__":
    app.run()
