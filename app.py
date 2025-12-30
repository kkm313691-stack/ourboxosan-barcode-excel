from flask import Flask, request, send_file, jsonify
import datetime
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image
import barcode
from barcode.writer import ImageWriter

app = Flask(__name__)

@app.route("/create_excel", methods=["POST"])
def create_excel():
    data = request.json

    name = data.get("name", "")
    exp = data.get("exp", "")
    qty_info = data.get("qty", "")
    qty_generate = int(data.get("barcode_qty", 1))

    today_prefix = datetime.datetime.now().strftime("%Y%m%d")
    filename = "barcode_label.xlsx"

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

        labels = ["품명", "소비기한", "수량", "바코드"]
        values = [name, exp, qty_info, barcode_number]

        for idx in range(4):
            ws[f"A{current_row+idx}"] = labels[idx]
            ws[f"B{current_row+idx}"] = values[idx]

            ws[f"A{current_row+idx}"].font = a_font
            ws[f"A{current_row+idx}"].alignment = a_align
            ws[f"A{current_row+idx}"].border = thin_border
            ws[f"B{current_row+idx}"].border = thin_border

        barcode_class = barcode.get_barcode_class("code128")
        barcode_obj = barcode_class(barcode_number, writer=ImageWriter())
        barcode_path = f"barcode_{i}"
        barcode_obj.save(barcode_path)

        img = Image(f"{barcode_path}.png")
        img.width = 600
        img.height = 150
        ws.add_image(img, f"B{current_row+3}")

        current_row += 4

    wb.save(filename)

    return send_file(
        filename,
        as_attachment=True,
        download_name="바코드_라벨.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/")
def health():
    return jsonify({"status": "ok"})
