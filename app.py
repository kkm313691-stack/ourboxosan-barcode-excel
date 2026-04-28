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

        # 🔥 mode 안정 처리
        mode = str(data.get("mode","normal")).lower()

        name = data.get("name","")
        exp = data.get("exp","")
        mfg = data.get("mfg","")
        lot = data.get("lot","")
        qty = data.get("qty","")
        count = int(data.get("barcode_qty") or 1)

        wb = Workbook()
        ws = wb.active

        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 140

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

            for r in range(row, row+4):
                ws.row_dimensions[r].height = 200

            # =========================
            # ✅ 일반 모드 (절대 수정 안함)
            # =========================
            if mode != "lot":
                labels = ["품명", "소비기한", "수량", "바코드"]
                values = [name, exp, qty]

                for idx in range(4):
                    a_cell = ws[f"A{row+idx}"]
                    b_cell = ws[f"B{row+idx}"]

                    a_cell.value = labels[idx]
                    a_cell.font = label_font
                    a_cell.alignment = center
                    a_cell.border = border

                    if labels[idx] != "바코드":
                        b_cell.value = values[idx]
                        b_cell.font = value_font
                        b_cell.alignment = center
                    else:
                        b_cell.value = ""

                    b_cell.border = border

                barcode_class = barcode.get_barcode_class("code128")
                barcode_obj = barcode_class(code, writer=ImageWriter())
                barcode_obj.save(f"barcode_{i}")

                img = Image(f"barcode_{i}.png")
                img.width = 600
                img.height = 150
                ws.add_image(img, f"B{row+3}")

            # =========================
            # ✅ 로트 모드 (정확 구현)
            # =========================
            else:
                ws[f"A{row}"].value = "품명"
                ws[f"B{row}"].value = name

                ws[f"A{row+1}"].value = exp
                ws[f"B{row+1}"].value = mfg

                ws[f"A{row+2}"].value = "수량"
                ws[f"B{row+2}"].value = qty

                ws[f"A{row+3}"].value = "바코드"
                ws[f"B{row+3}"].value = lot

                for idx in range(4):
                    a_cell = ws[f"A{row+idx}"]
                    b_cell = ws[f"B{row+idx}"]

                    a_cell.font = label_font
                    a_cell.alignment = center
                    a_cell.border = border

                    b_cell.font = value_font
                    b_cell.alignment = center
                    b_cell.border = border

                barcode_class = barcode.get_barcode_class("code128")
                barcode_obj = barcode_class(code, writer=ImageWriter())
                barcode_obj.save(f"barcode_{i}")

                img = Image(f"barcode_{i}.png")
                img.width = 600
                img.height = 150
                ws.add_image(img, f"A{row+3}")

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


if __name__ == "__main__":
    app.run()
