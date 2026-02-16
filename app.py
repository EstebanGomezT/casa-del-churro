import os
import re
from datetime import datetime, date
from calendar import monthrange
from pathlib import Path
from flask import Flask, request, send_from_directory, abort, jsonify, render_template

from db import init_db, insert_sale, fetch_sales_between, update_sale, delete_sale
from report import create_month_report_xlsx

app = Flask(__name__)
init_db()

STORAGE_RECEIPTS = Path("storage/receipts")
STORAGE_RECEIPTS.mkdir(parents=True, exist_ok=True)

# -----------------------
# Helpers
# -----------------------
def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")

def parse_int_or_none(text: str):
    text = (text or "").strip()
    text = re.sub(r"[^\d]", "", text)
    if not text:
        return None
    return int(text)

def parse_date_yyyy_mm_dd(text: str):
    try:
        return datetime.strptime(text.strip(), "%Y-%m-%d").date()
    except Exception:
        return None

PUNTOS_VENTA_VALIDOS = {"Carro Plaza", "Carro Amarillo", "Carro Chico", "Carro Tren", "Modulo"}

# -----------------------
# WEB UI
# -----------------------
@app.get("/")
def web_index():
    return render_template("index.html")

# Servir reportes / archivos
@app.get("/files/<path:filename>")
def serve_report(filename: str):
    reports_dir = Path("storage/reports")
    file_path = reports_dir / filename
    if not file_path.exists() or not file_path.is_file():
        abort(404)
    return send_from_directory(reports_dir, filename, as_attachment=True)

# -----------------------
# API
# -----------------------
@app.post("/api/sales")
def api_create_sale():
    try:
        sale_date = request.form.get("date", "").strip()
        if not parse_date_yyyy_mm_dd(sale_date):
            return jsonify({"error": "Fecha inválida (YYYY-MM-DD)"}), 400

        total = parse_int_or_none(request.form.get("total", ""))
        debit = parse_int_or_none(request.form.get("debit", ""))
        credit = parse_int_or_none(request.form.get("credit", ""))
        cash = parse_int_or_none(request.form.get("cash", ""))
        boletas_debit = parse_int_or_none(request.form.get("boletas_debit", ""))
        boletas_credit = parse_int_or_none(request.form.get("boletas_credit", ""))
        boletas_cash = parse_int_or_none(request.form.get("boletas_cash", ""))
        folio = request.form.get("folio", "").strip()

        if any(v is None for v in [total, debit, credit, cash,
                                    boletas_debit, boletas_credit, boletas_cash]):
            return jsonify({"error": "Todos los campos numéricos son obligatorios"}), 400
        if not folio:
            return jsonify({"error": "El folio es obligatorio"}), 400

        receipt_path = ""
        if "receipt" in request.files:
            f = request.files["receipt"]
            if f.filename:
                ext = Path(f.filename).suffix or ".jpg"
                filename = f"web_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
                path = STORAGE_RECEIPTS / filename
                f.save(path)
                receipt_path = str(path)

        if not receipt_path:
            return jsonify({"error": "La boleta (imagen) es obligatoria"}), 400

        punto_venta = request.form.get("punto_venta", "").strip()
        if punto_venta not in PUNTOS_VENTA_VALIDOS:
            return jsonify({"error": "Punto de venta inválido"}), 400

        sale = {
            "phone": "web",
            "date": sale_date,
            "total": total,
            "debit": debit,
            "credit": credit,
            "cash": cash,
            "boletas_debit": boletas_debit,
            "boletas_credit": boletas_credit,
            "boletas_cash": boletas_cash,
            "punto_venta": punto_venta,
            "folio": folio,
            "receipt_path": receipt_path,
            "created_at": now_iso(),
        }
        sale_id = insert_sale(sale)
        return jsonify({"ok": True, "id": sale_id})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.get("/api/sales")
def api_list_sales():
    month_str = request.args.get("month", "")
    m = re.match(r"^(\d{4})-(\d{2})$", month_str)
    if not m:
        return jsonify({"error": "Formato: YYYY-MM"}), 400
    year, month = int(m.group(1)), int(m.group(2))
    date_from = date(year, month, 1)
    last_day = monthrange(year, month)[1]
    date_to = date(year, month, last_day)
    today = date.today()
    if year == today.year and month == today.month:
        date_to = today
    rows = fetch_sales_between(date_from.isoformat(), date_to.isoformat())
    return jsonify(rows)

@app.get("/api/report")
def api_download_report():
    month_str = request.args.get("month", "")
    m = re.match(r"^(\d{4})-(\d{2})$", month_str)
    if not m:
        return jsonify({"error": "Formato: YYYY-MM"}), 400
    year, month = int(m.group(1)), int(m.group(2))
    date_from = date(year, month, 1)
    last_day = monthrange(year, month)[1]
    date_to = date(year, month, last_day)
    today = date.today()
    if year == today.year and month == today.month:
        date_to = today
    rows = fetch_sales_between(date_from.isoformat(), date_to.isoformat())
    xlsx_path = create_month_report_xlsx(rows, year, month, date_to.isoformat())
    return send_from_directory(xlsx_path.parent, xlsx_path.name, as_attachment=True)

@app.put("/api/sales/<int:sale_id>")
def api_update_sale(sale_id: int):
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No hay datos"}), 400

        fields = {}
        if "date" in data:
            if not parse_date_yyyy_mm_dd(data["date"]):
                return jsonify({"error": "Fecha inválida"}), 400
            fields["date"] = data["date"]
        for f in ("total", "debit", "credit", "cash",
                   "boletas_debit", "boletas_credit", "boletas_cash"):
            if f in data:
                v = parse_int_or_none(str(data[f]))
                if v is None:
                    return jsonify({"error": f"{f} debe ser un número"}), 400
                fields[f] = v
        if "folio" in data:
            if not data["folio"].strip():
                return jsonify({"error": "Folio vacío"}), 400
            fields["folio"] = data["folio"].strip()
        if "punto_venta" in data:
            if data["punto_venta"].strip() not in PUNTOS_VENTA_VALIDOS:
                return jsonify({"error": "Punto de venta inválido"}), 400
            fields["punto_venta"] = data["punto_venta"].strip()

        if not fields:
            return jsonify({"error": "No hay campos para actualizar"}), 400

        update_sale(sale_id, fields)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.delete("/api/sales/<int:sale_id>")
def api_delete_sale(sale_id: int):
    try:
        deleted = delete_sale(sale_id)
        if not deleted:
            return jsonify({"error": "No encontrado"}), 404
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port)
