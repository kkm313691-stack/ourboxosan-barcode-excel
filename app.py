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
CORS(app)

@app.route("/create_excel", methods=["POST","OPTIONS"])
def create_excel():
    if request.method == "OPTIONS":
        return make_response("",200)

    try:
        data = request.json

        mode = data.get("mode","normal")
        name = data.get("name","")
        exp = data.get("exp","")
        mfg = data.get("mfg","")
        lot = data.get("lot","")
        qty = data.get("qty","")
        count = int(data.get("barcode_qty") or 1)

        wb = Workbook()
        ws = wb.active

        # ✅ 열 너비 (먼저 설정)
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 140

        # ✅ 스타일 정의
        label_font = Font(size=40, bold=True)
        value_font = Font(size=100, bold=True)

        center = Alignment(horizontal="center", vertical="center")

        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        row = 1

        for i in range(count):

            code = datetime.datetime.now().strftime("%Y%m%d") + f"{i:04d}"

            # ✅ 행 높이 먼저
            for r in range(row, row+4):
                ws.row_dimensions[r].height = 200

            # ======================
            # 값 먼저 세팅
            # ======================
            if mode == "lot":
                data_map = [
                    ("품명", name),
                    (exp, mfg),      # 소비기한 / 제조일자
                    ("수량", qty),
                    ("바코드", lot)
                ]
            else:
                data_map = [
                    ("품명", name),
                    ("소비기한", exp),
                    ("수량", qty),
                    ("바코드", "")
                ]

            # ======================
            # 값 + 스타일 한번에 적용
            # ======================
            for idx, (a_val, b_val) in enumerate(data_map):
                a_cell = ws[f"A{row+idx}"]
                b_cell = ws[f"B{row+idx}"]

                # 값
                a_cell.value = a_val
                b_cell.value = b_val

                # 스타일 (무조건 적용)
                a_cell.font = label_font
                a_cell.alignment = center
                a_cell.border = border

                b_cell.font = value_font
                b_cell.alignment = center
                b_cell.border = border

            # ======================
            # 바코드
            # ======================
            barcode_class = barcode.get_barcode_class("code128")
            barcode_obj = barcode_class(code, writer=ImageWriter())
            barcode_obj.save(f"barcode_{i}")

            img = Image(f"barcode_{i}.png")
            img.width = 600
            img.height = 150

            if mode == "lot":
                ws.add_image(img, f"A{row+3}")
            else:
                ws.add_image(img, f"B{row+3}")

            row += 4

        file = "barcode.xlsx"
        wb.save(file)

        for i in range(count):
            try:
                os.remove(f"barcode_{i}.png")
            except:
                pass

        return send_file(file, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}),500
