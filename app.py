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

        mode = str(data.get("mode", "normal")).lower()

        name = data.get("name", "")
        exp = data.get("exp", "")
        qty_info = data.get("qty", "")

        # 🔥 일반모드에서는 강제 제거
        if mode != "lot":
            mfg = ""
            lot = ""
        else:
            mfg = data.get("mfg", "")
            lot = data.get("lot", "")

        qty_generate = int(data.get("barcode_qty") or 1)

        today_prefix = datetime.datetime.now().strftime("%Y%m%d")

        wb = Workbook()
        ws = wb.active
        ws.title = "바코드 라벨"

        # 기존 스타일 유지
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 140

        label_font = Font(size=40, bold=True)
        label_align = Alignment(horizontal="center", vertical="center")

        value_font = Font(size=100, bold=True)
        value_align = Alignment(horizontal="center", vertical="center")

        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        row = 1

        for i in range(1, qty_generate + 1):
            barcode_number = f"{today_prefix}{i:04d}"

            # 행 높이 유지
            for r in range(row, row + 4):
                ws.row_dimensions[r].height = 200

            # =========================
            # ✅ 일반 모드 (절대 수정 금지)
            # =========================
            if mode != "lot":
                labels = ["품명", "소비기한", "수량", "바코드"]
                values = [name, exp, qty_info]

                for idx in range(4):
                    a_cell = ws[f"A{row+idx}"]
                    b_cell = ws[f"B{row+idx}"]

                    a_cell.value = labels[idx]
                    a_cell.font = label_font
                    a_cell.alignment = label_align
                    a_cell.border = border

                    if labels[idx] != "바코드":
                        b_cell.value = values[idx]
                        b_cell.font = value_font
                        b_cell.alignment = value_align
                    else:
                        b_cell.value = ""

                    b_cell.border = border

                # 바코드 이미지 (기존 그대로 B열)
                barcode_class = barcode.get_barcode_class("code128")
                barcode_obj = barcode_class(barcode_number, writer=ImageWriter())
                barcode_obj.save(f"barcode_{i}")

                img = Image(f"barcode_{i}.png")
                img.width = 600
                img.height = 150
                ws.add_image(img, f"B{row+3}")

            # =========================
            # ✅ 로트 모드 (완전 고정)
            # =========================
            else:
                # 1행
                ws[f"A{row}"].value = "품명"
                ws[f"B{row}"].value = name

                # 2행 (핵심)
                ws[f"A{row+1}"].value = "소비기한"
                ws[f"B{row+1}"].value = mfg

                # 3행
                ws[f"A{row+2}"].value = "수량"
                ws[f"B{row+2}"].value = qty_info

                # 4행
                ws[f"A{row+3}"].value = ""   # 텍스트 제거 (이미지용)
                ws[f"B{row+3}"].value = lot

                # 스타일 적용
                for idx in range(4):
                    a_cell = ws[f"A{row+idx}"]
                    b_cell = ws[f"B{row+idx}"]

                    a_cell.font = label_font
                    a_cell.alignment = label_align
                    a_cell.border = border

                    b_cell.font = value_font
                    b_cell.alignment = value_align
                    b_cell.border = border

                # 바코드 이미지 (A열)
                barcode_class = barcode.get_barcode_class("code128")
                barcode_obj = barcode_class(barcode_number, writer=ImageWriter())
                barcode_obj.save(f"barcode_{i}")

                img = Image(f"barcode_{i}.png")
                img.width = 600
                img.height = 150
                ws.add_image(img, f"A{row+3}")

            row += 4

        file_path = "바코드_라벨.xlsx"
        wb.save(file_path)

        # 임시 파일 삭제
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


if __name__ == "__main__":
    app.run()
