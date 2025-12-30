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

# ✅ CORS 완전 허용 (모든 origin, 모든 method)
CORS(
    app,
    resources={r"/*": {"origins": "*"}},
    supports_credentials=True
)

# --------------------------
# Preflight(OPTIONS) 대응
# --------------------------
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

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 140

        a_font = Font(size=40, bold=True)
        a_align = Alignment(horizontal="center", vertical="center")

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        current_row = 1

        for i in range(1, qty_generate + 1):
            barcode_number = f"{today_prefix}{i:04d}"

            for r in range(current_row, current_row + 4):
                ws.row_dimensions[r].height = 180

            ws[f"A{current_row}"] = "품명"
            ws[f"A{current_row+1}"] = "소비기한"
            ws[f"A{current_row+2}"] = "수량"
            ws[f"A{current_row+3}"] = "바코드"

            for r in range(current_row, current_row + 4):
                ws[f"A{r}"].font = a_font
                ws[f"A{r}"].alignment = a_align
                ws[f"A{r}"].border = thin_border

            ws[f"B{current_row}"] = name
            ws[f"B{current_row+1}"] = exp
            ws[f"B{current_row+2}"] = qty_info
            ws[f"B{current_row+3}"] = barcode_number

            for r in range(current_row, current_row + 4):
                ws[f"B{r}"].border = thin_border

            barcode_path = f"barcode_{i}.png"
            barcode_class = barcode.get_barcode_class("code128")
            barcode_obj = barcode_class(barcode_number, writer=ImageWriter())
            barcode_obj.save(f"barcode_{i}")

            img = Image(barcode_path)
            img.width = 600
            img.height = 150
            ws.add_image(img, f"B{current_row+3}")

            current_row += 4

        file_path = "바코드_라벨.xlsx"
        wb.save(file_path)

        for i in range(1, qty_generate + 1):
            try:
                os.remove(f"barcode_{i}.png")
            except:
                pass

        return send_file(
            file_path,
            as_attachment=True,
            download_name="바코드_라벨.xlsx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500
