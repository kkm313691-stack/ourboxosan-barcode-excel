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


def base(ws):
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
# 일반 모드 (변경 없음)
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

        label_font, value_font, align, border = base(ws)

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
# 로트 모드 (요청 구조 유지)
# =========================
@app.route("/create_excel_lot", methods=["POST"])
def lot():
    try:
        data = request.json

        name = data.get("name","")
        mfg = data.get("mfg","")
        qty = data.get("qty","")
        lot = data.get("lot","")
        count = int(data.get("barcode_qty") or 1)

        wb = Workbook()
        ws = wb.active

        ws.column_dimensions["A"].width = 25
        ws.column_dimensions["B"].width = 85

        label_font = Font(size=25, bold=True)
        value_font = Font(size=80, bold=True)

        align = Alignment(horizontal="center", vertical="center")

        row = 1

        for i in range(count):

            code = datetime.datetime.now().strftime("%Y%m%d") + f"{i:04d}"

            for r in range(row, row+6):
                ws.row_dimensions[r].height = 160

            items = [
                ("품명", name),
                ("소비기한", mfg),
                ("제조일자", mfg),
                ("수량", qty),
                ("로트번호", lot),
                ("바코드", code)
            ]

            for idx, (label, value) in enumerate(items):

                r = row + idx

                ws[f"A{r}"].value = label
                ws[f"A{r}"].font = label_font
                ws[f"A{r}"].alignment = align

                if label != "바코드":
                    ws[f"B{r}"].value = value
                    ws[f"B{r}"].font = value_font
                    ws[f"B{r}"].alignment = align
                else:
                    ws[f"B{r}"].value = ""

            barcode_class = barcode.get_barcode_class("code128")
            barcode_obj = barcode_class(code, writer=ImageWriter())
            barcode_obj.save(f"barcode_{i}")

            img = Image(f"barcode_{i}.png")
            img.width = 615
            img.height = 142

            ws.add_image(img, f"B{row+5}")

            row += 7

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


if __name__ == "__main__":
    app.run()
