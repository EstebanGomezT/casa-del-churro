import sqlite3
from pathlib import Path
from typing import Any, Dict

DB_PATH = Path("sales.db")

def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db() -> None:
    with get_conn() as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            phone TEXT NOT NULL,
            date TEXT NOT NULL,
            total INTEGER NOT NULL,
            debit INTEGER NOT NULL,
            credit INTEGER NOT NULL,
            cash INTEGER NOT NULL,
            boletas_debit INTEGER NOT NULL DEFAULT 0,
            boletas_credit INTEGER NOT NULL DEFAULT 0,
            boletas_cash INTEGER NOT NULL DEFAULT 0,
            punto_venta TEXT NOT NULL DEFAULT '',
            folio TEXT NOT NULL,
            receipt_path TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_sales_date ON sales(date)")
        conn.commit()

def insert_sale(sale: Dict[str, Any]) -> int:
    with get_conn() as conn:
        cur = conn.execute("""
        INSERT INTO sales(phone, date, total, debit, credit, cash,
                          boletas_debit, boletas_credit, boletas_cash,
                          punto_venta, folio, receipt_path, created_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            sale["phone"],
            sale["date"],
            sale["total"],
            sale["debit"],
            sale["credit"],
            sale["cash"],
            sale["boletas_debit"],
            sale["boletas_credit"],
            sale["boletas_cash"],
            sale["punto_venta"],
            sale["folio"],
            sale["receipt_path"],
            sale["created_at"],
        ))
        conn.commit()
        return int(cur.lastrowid)

def update_sale(sale_id: int, fields: Dict[str, Any]) -> bool:
    allowed = {"date", "total", "debit", "credit", "cash",
               "boletas_debit", "boletas_credit", "boletas_cash",
               "punto_venta", "folio"}
    updates = {k: v for k, v in fields.items() if k in allowed}
    if not updates:
        return False
    set_clause = ", ".join(f"{k}=?" for k in updates)
    values = list(updates.values()) + [sale_id]
    with get_conn() as conn:
        conn.execute(f"UPDATE sales SET {set_clause} WHERE id=?", values)
        conn.commit()
    return True

def delete_sale(sale_id: int) -> bool:
    with get_conn() as conn:
        cur = conn.execute("DELETE FROM sales WHERE id=?", (sale_id,))
        conn.commit()
        return cur.rowcount > 0

def fetch_sales_between(date_from: str, date_to: str) -> list[Dict[str, Any]]:
    with get_conn() as conn:
        cur = conn.execute("""
        SELECT id, date, total, debit, credit, cash,
               boletas_debit, boletas_credit, boletas_cash,
               punto_venta, folio, receipt_path, phone, created_at
        FROM sales
        WHERE date >= ? AND date <= ?
        ORDER BY date ASC, id ASC
        """, (date_from, date_to))
        rows = cur.fetchall()
        return [dict(r) for r in rows]
