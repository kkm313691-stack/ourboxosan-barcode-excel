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

        mode = data.get("mode", "normal")  # ✅ 모드 추가

        name = data.get("name", "")
        exp = data.get("exp", "")
        mfg = data.get("mfg", "")
        lot = data.get("lot", "")
        qty_info = data.get("qty", "")
        qty_generate = int(data.get("barcode_qty") or 1)

        today_prefix = datetime.datetime.now().strftime("%Y%m%d")

        wb = Workbook()
        ws = wb.active
        ws.title = "바코드 라벨"

        # ✅ 기존 스타일 유지
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
            # ✅ 모드별 레이아웃
            # =========================
            if mode == "lot":
                labels_A = ["품명", "소비기한", "수량", "바코드"]
                values_B = [
                    name,
                    mfg,  # 제조일자
                    qty_info,
                    lot
                ]
                extra_B_row2 = exp  # 소비기한은 A2에 표시됨

            else:
                labels_A = ["품명", "소비기한", "수량", "바코드"]
                values_B = [
                    name,
                    exp,
                    qty_info,
                    ""
                ]

            for idx in range(4):
                a_cell = ws[f"A{row+idx}"]
                b_cell = ws[f"B{row+idx}"]

                # A열
                a_cell.value = labels_A[idx]
                a_cell.font = label_font
                a_cell.alignment = label_align
                a_cell.border = border

                # B열
                if labels_A[idx] != "바코드":
                    if mode == "lot" and idx == 1:
                        # 소비기한 / 제조일자 분리
                        b_cell.value = mfg
                        a_cell.value = exp
                    else:
                        b_cell.value = values_B[idx]

                    b_cell.font = value_font
                    b_cell.alignment = value_align
                else:
                    b_cell.value = values_B[idx]

                b_cell.border = border

            # =========================
            # 바코드 생성
            # =========================
            barcode_class = barcode.get_barcode_class("code128")
            barcode_obj = barcode_class(barcode_number, writer=ImageWriter())
            barcode_obj.save(f"barcode_{i}")

            img = Image(f"barcode_{i}.png")
            img.width = 600
            img.height = 150

            # ✅ 위치 (요구사항)
            if mode == "lot":
                ws.add_image(img, f"A{row+3}")
            else:
                ws.add_image(img, f"B{row+3}")

            row += 4

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


if __name__ == "__main__":
    app.run()
