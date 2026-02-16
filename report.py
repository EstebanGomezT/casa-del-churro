from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XlImage
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent

IMG_HEIGHT_PX = 120
ROW_HEIGHT_PT = 95

def create_month_report_xlsx(rows: list[dict], year: int, month: int, date_to_inclusive: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Ventas"

    headers = ["Fecha", "Punto Venta", "Total", "Débito", "Crédito", "Efectivo",
               "Bol. Débito", "Bol. Crédito", "Bol. Efectivo", "Folio", "Boleta"]
    ws.append(headers)

    total_sum = debit_sum = credit_sum = cash_sum = 0
    bd_sum = bc_sum = bca_sum = 0

    for idx, r in enumerate(rows):
        ws.append([
            r["date"], r.get("punto_venta", ""),
            r["total"], r["debit"], r["credit"], r["cash"],
            r.get("boletas_debit", 0), r.get("boletas_credit", 0), r.get("boletas_cash", 0),
            r["folio"], ""
        ])
        total_sum += int(r["total"])
        debit_sum += int(r["debit"])
        credit_sum += int(r["credit"])
        cash_sum += int(r["cash"])
        bd_sum += int(r.get("boletas_debit", 0))
        bc_sum += int(r.get("boletas_credit", 0))
        bca_sum += int(r.get("boletas_cash", 0))

        # Insert receipt image
        receipt = r.get("receipt_path", "")
        receipt_file = (BASE_DIR / receipt) if receipt else None
        if receipt_file and receipt_file.exists():
            try:
                img = XlImage(str(receipt_file))
                img.height = IMG_HEIGHT_PX
                img.width = int(img.width * (IMG_HEIGHT_PX / img.height)) if img.height else 100
                cell = f"K{idx + 2}"  # row idx+2 (1-based + header)
                ws.add_image(img, cell)
                ws.row_dimensions[idx + 2].height = ROW_HEIGHT_PT
            except Exception:
                ws.cell(row=idx + 2, column=11, value="(imagen no disponible)")

    ws.append([])
    ws.append(["TOTAL", "", total_sum, debit_sum, credit_sum, cash_sum,
               bd_sum, bc_sum, bca_sum, "", ""])

    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16
    ws.column_dimensions["K"].width = 22  # wider for images

    out_dir = BASE_DIR / "storage/reports"
    out_dir.mkdir(parents=True, exist_ok=True)

    filename = f"reporte_{year:04d}-{month:02d}_hasta_{date_to_inclusive}.xlsx"
    path = out_dir / filename
    wb.save(path)
    return path
