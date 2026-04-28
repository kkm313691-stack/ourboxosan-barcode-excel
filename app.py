import os
import datetime
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image

import barcode
from barcode.writer import ImageWriter

app = Flask(__name__)
CORS(app)


def base_style(ws):
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 140

    label_font = Font(size=40, bold=True)
    value_font = Font(size=100, bold=True)

    align = Alignment(horizontal="center", vertical="center")

    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    return label_font, value_font, align, border


# =========================
# ✅ 일반 바코드 (절대 수정 금지)
# =========================
@app.route("/create_excel_normal", methods=["POST"])
def normal():
    try:
        data = request.json

        name = data.get("name","")
        exp = data.get("exp","")
        qty = data.get("qty","")
        count = int(data.get("barcode_qty") or 1)

        wb = Workbook()
        ws = wb.active

        label_font, value_font, align, border = base_style(ws)

        row = 1

        for i in range(count):
            code = datetime.datetime.now().strftime("%Y%m%d") + f"{i:04d}"

            for r in range(row, row+4):
                ws.row_dimensions[r].height = 200

            labels = ["품명","소비기한","수량","바코드"]
            values = [name,exp,qty]

            for idx in range(4):
                a = ws[f"A{row+idx}"]
                b = ws[f"B{row+idx}"]

                a.value = labels[idx]
                a.font = label_font
                a.alignment = align
                a.border = border

                if labels[idx] != "바코드":
                    b.value = values[idx]
                    b.font = value_font
                    b.alignment = align
                else:
                    b.value = ""

                b.border = border

            barcode_class = barcode.get_barcode_class("code128")
            barcode_obj = barcode_class(code, writer=ImageWriter())
            barcode_obj.save(f"barcode_{i}")

            img = Image(f"barcode_{i}.png")
            img.width = 600
            img.height = 150
            ws.add_image(img, f"B{row+3}")

            row += 4

        file = "barcode.xlsx"
        wb.save(file)

        for i in range(count):
            try: os.remove(f"barcode_{i}.png")
            except: pass

        return send_file(file, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}),500


# =========================
# ✅ 로트 바코드 (완전 고정)
# =========================
@app.route("/create_excel_lot", methods=["POST"])
def lot():
    try:
        data = request.json

        name = data.get("name","")
        mfg = data.get("mfg","")   # 제조일자
        qty = data.get("qty","")
        lot = data.get("lot","")
        count = int(data.get("barcode_qty") or 1)

        wb = Workbook()
        ws = wb.active

        label_font, value_font, align, border = base_style(ws)

        row = 1

        for i in range(count):

            code = datetime.datetime.now().strftime("%Y%m%d") + f"{i:04d}"

            # 행 높이 고정
            ws.row_dimensions[row].height = 200
            ws.row_dimensions[row+1].height = 200
            ws.row_dimensions[row+2].height = 200
            ws.row_dimensions[row+3].height = 200

            # 🔥 완전 고정 좌표
            ws[f"A{row}"].value = "품명"
            ws[f"B{row}"].value = name

            ws[f"A{row+1}"].value = "소비기한"
            ws[f"B{row+1}"].value = mfg

            ws[f"A{row+2}"].value = "수량"
            ws[f"B{row+2}"].value = qty

            ws[f"A{row+3}"].value = ""  # 이미지용
            ws[f"B{row+3}"].value = lot

            # 스타일 적용
            for r in range(row, row+4):
                a = ws[f"A{r}"]
                b = ws[f"B{r}"]

                a.font = label_font
                a.alignment = align
                a.border = border

                b.font = value_font
                b.alignment = align
                b.border = border

            # 바코드 생성
            barcode_class = barcode.get_barcode_class("code128")
            barcode_obj = barcode_class(code, writer=ImageWriter())
            barcode_obj.save(f"barcode_{i}")

            img = Image(f"barcode_{i}.png")
            img.width = 600
            img.height = 150

            # 🔥 A4에만 넣는다
            ws.add_image(img, f"A{row+3}")

            row += 4

        file = "barcode.xlsx"
        wb.save(file)

        for i in range(count):
            try: os.remove(f"barcode_{i}.png")
            except: pass

        return send_file(file, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}),500


if __name__ == "__main__":
    app.run()
