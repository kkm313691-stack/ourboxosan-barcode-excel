import os
import datetime
from flask import Flask, request, send_file, jsonify, make_response
from flask_cors import CORS

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image

import barcode
from barcode.writer import ImageWriter

app = Flask(__name__)

# CORS 완전 허용
CORS(app, resources={r"/*": {"origins": "*"}})

@app.after_request
def after_request(response):
    response.headers.add("Access-Control-Allow-Origin", "*")
    response.headers.add("Access-Control-Allow-Headers", "Content-Type")
    response.headers.add("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    return response


@app.route("/")
def health():
    return jsonify({"status": "ok"})


@app.route("/create_excel", methods=["POST", "OPTIONS"])
def create_excel():
    if request.method == "OPTIONS":
        return make_response("", 200)

    try:
        data = request.json

        name = data.get("name", "")
        exp = data.get("exp", "")
        qty_info = data.get("qty", "")
        qty_generate = int(data.get("barcode_qty", 1))

        today_prefix = datetime.datetime.now().strftime("%Y%m%d")

        wb = Workbook()
        ws = wb.active
        ws.title = "바코드 라벨"

        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 140

        a_font = Font(size=40, bold=True)
        a_align = Alignment(horizontal="center", vertical="center")
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        row = 1

        for i in range(1, qty_generate + 1):
            barcode_number = f"{today_prefix}{i:04d}"

            for r in range(row, row + 4):
                ws.row_dimensions[r].height = 180

            labels = ["품명", "소비기한", "수량", "바코드"]
            values = [name, exp, qty_info, barcode_number]

            for idx in range(4):
                ws[f"A{row+idx}"] = labels[idx]
                ws[f"B{row+idx}"] = values[idx]

                ws[f"A{row+idx}"].font = a_font
                ws[f"A{row+idx}"].alignment = a_align

                ws[f"A{row+idx}"].border = border
                ws[f"B{row+idx}"].border = border

            # 바코드 이미지 생성
            barcode_class = barcode.get_barcode_class("code128")
            barcode_obj = barcode_class(barcode_number, writer=ImageWriter())
            barcode_obj.save(f"barcode_{i}")

            img = Image(f"barcode_{i}.png")
            img.width = 600
            img.height = 150
            ws.add_image(img, f"B{row+3}")

            row += 4

        file_path = "바코드_라벨.xlsx"
        wb.save(file_path)

        for i in range(1, qty_generate + 1):
            os.remove(f"barcode_{i}.png")

        return send_file(
            file_path,
            as_attachment=True,
            download_name="바코드_라벨.xlsx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500
