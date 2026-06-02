"""SQLite 주문 DB 스키마 및 CRUD."""

from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

from src.config_loader import load_config, resolve_path
from src.settings import get_database_path


def get_db_path() -> Path:
    """환경변수 DATABASE_URL 또는 config.yaml 기준 DB 경로."""
    return get_database_path(load_config())


def connect() -> sqlite3.Connection:
    db_path = get_db_path()
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db(conn: sqlite3.Connection | None = None) -> None:
    """orders, order_items, import_logs, order_images 테이블 생성."""
    own_conn = conn is None
    if own_conn:
        conn = connect()

    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS import_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_file TEXT NOT NULL,
            sheet_name TEXT,
            imported_at TEXT NOT NULL,
            order_count INTEGER DEFAULT 0,
            item_count INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            import_id INTEGER,
            source_file TEXT NOT NULL,
            sheet_name TEXT NOT NULL,
            sheet_date TEXT,
            order_no TEXT NOT NULL,
            order_date TEXT,
            platform TEXT,
            file_ref TEXT,
            customer TEXT,
            phone TEXT,
            address TEXT,
            sales REAL,
            deduct REAL,
            ship REAL,
            total REAL,
            pay_card TEXT,
            pay_transfer TEXT,
            pay_bank TEXT,
            remark TEXT,
            has_image INTEGER DEFAULT 0,
            order_qty REAL,
            expected_ship_type TEXT,
            expected_freight TEXT,
            expected_ship_qty REAL,
            payment_status TEXT,
            deposit_date TEXT,
            UNIQUE(source_file, sheet_name, order_no),
            FOREIGN KEY (import_id) REFERENCES import_logs(id)
        );

        CREATE TABLE IF NOT EXISTS order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER NOT NULL,
            line_no INTEGER NOT NULL,
            frame TEXT,
            size TEXT,
            width REAL,
            height REAL,
            color TEXT,
            plate TEXT,
            acrylic TEXT,
            hook TEXT,
            item_note TEXT,
            qty REAL,
            unit_price REAL,
            FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS order_images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER,
            source_file TEXT NOT NULL,
            sheet_name TEXT NOT NULL,
            excel_row INTEGER,
            image_file TEXT NOT NULL,
            mapped INTEGER DEFAULT 0,
            FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE SET NULL
        );

        CREATE INDEX IF NOT EXISTS idx_orders_platform ON orders(platform);
        CREATE INDEX IF NOT EXISTS idx_orders_sheet_date ON orders(sheet_date);
        CREATE INDEX IF NOT EXISTS idx_order_items_order ON order_items(order_id);
        """
    )
    _migrate_columns(conn)
    conn.commit()
    if own_conn:
        conn.close()


def _migrate_columns(conn: sqlite3.Connection) -> None:
    """기존 DB에 엑셀 양식 확장 컬럼 추가."""
    existing = {row[1] for row in conn.execute("PRAGMA table_info(orders)")}
    for name, typ in (
        ("order_qty", "REAL"),
        ("expected_ship_type", "TEXT"),
        ("expected_freight", "TEXT"),
        ("expected_ship_qty", "REAL"),
        ("payment_status", "TEXT"),
        ("deposit_date", "TEXT"),
    ):
        if name not in existing:
            conn.execute(f"ALTER TABLE orders ADD COLUMN {name} {typ}")
    # 마이그레이션 후 인덱스 (컬럼 존재 보장)
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_orders_deposit_date ON orders(deposit_date)"
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_orders_payment_status ON orders(payment_status)"
    )


def insert_import_log(
    conn: sqlite3.Connection,
    source_file: str,
    sheet_name: str,
    order_count: int,
    item_count: int,
) -> int:
    cur = conn.execute(
        """
        INSERT INTO import_logs (source_file, sheet_name, imported_at, order_count, item_count)
        VALUES (?, ?, ?, ?, ?)
        """,
        (source_file, sheet_name, datetime.now().isoformat(), order_count, item_count),
    )
    conn.commit()
    return int(cur.lastrowid)


def upsert_order(conn: sqlite3.Connection, order: dict[str, Any], import_id: int) -> int:
    """주문 upsert 후 order_id 반환."""
    conn.execute(
        """
        INSERT INTO orders (
            import_id, source_file, sheet_name, sheet_date, order_no, order_date,
            platform, file_ref, customer, phone, address,
            sales, deduct, ship, total,
            pay_card, pay_transfer, pay_bank, remark, has_image,
            order_qty, expected_ship_type, expected_freight, expected_ship_qty
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(source_file, sheet_name, order_no) DO UPDATE SET
            import_id=excluded.import_id,
            sheet_date=excluded.sheet_date,
            order_date=excluded.order_date,
            platform=excluded.platform,
            file_ref=excluded.file_ref,
            customer=excluded.customer,
            phone=excluded.phone,
            address=excluded.address,
            sales=excluded.sales,
            deduct=excluded.deduct,
            ship=excluded.ship,
            total=excluded.total,
            pay_card=excluded.pay_card,
            pay_transfer=excluded.pay_transfer,
            pay_bank=excluded.pay_bank,
            remark=excluded.remark,
            has_image=excluded.has_image,
            order_qty=excluded.order_qty,
            expected_ship_type=excluded.expected_ship_type,
            expected_freight=excluded.expected_freight,
            expected_ship_qty=excluded.expected_ship_qty
        """,
        (
            import_id,
            order["source_file"],
            order["sheet_name"],
            order.get("sheet_date"),
            order["order_no"],
            order.get("order_date"),
            order.get("platform"),
            order.get("file_ref"),
            order.get("customer"),
            order.get("phone"),
            order.get("address"),
            order.get("sales"),
            order.get("deduct"),
            order.get("ship"),
            order.get("total"),
            order.get("pay_card"),
            order.get("pay_transfer"),
            order.get("pay_bank"),
            order.get("remark"),
            order.get("has_image", 0),
            order.get("order_qty"),
            order.get("expected_ship_type"),
            order.get("expected_freight"),
            order.get("expected_ship_qty"),
        ),
    )
    row = conn.execute(
        """
        SELECT id FROM orders
        WHERE source_file=? AND sheet_name=? AND order_no=?
        """,
        (order["source_file"], order["sheet_name"], order["order_no"]),
    ).fetchone()
    order_id = int(row["id"])

    # 품목은 재적재 (간단 upsert)
    conn.execute("DELETE FROM order_items WHERE order_id=?", (order_id,))
    for idx, item in enumerate(order.get("items", []), start=1):
        conn.execute(
            """
            INSERT INTO order_items (
                order_id, line_no, frame, size, width, height, color,
                plate, acrylic, hook, item_note, qty, unit_price
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                order_id,
                idx,
                item.get("frame"),
                item.get("size"),
                item.get("width"),
                item.get("height"),
                item.get("color"),
                item.get("plate"),
                item.get("acrylic"),
                item.get("hook"),
                item.get("item_note"),
                item.get("qty"),
                item.get("unit_price"),
            ),
        )
    conn.commit()
    return order_id


def mark_order_has_image(conn: sqlite3.Connection, order_id: int) -> None:
    conn.execute("UPDATE orders SET has_image=1 WHERE id=?", (order_id,))
    conn.commit()


def insert_order_image(
    conn: sqlite3.Connection,
    *,
    order_id: int | None,
    source_file: str,
    sheet_name: str,
    excel_row: int,
    image_file: str,
    mapped: bool,
) -> None:
    conn.execute(
        """
        INSERT INTO order_images (order_id, source_file, sheet_name, excel_row, image_file, mapped)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (order_id, source_file, sheet_name, excel_row, image_file, 1 if mapped else 0),
    )
    conn.commit()


def fetch_orders_for_month(conn: sqlite3.Connection, year_month: str) -> list[sqlite3.Row]:
    """YYYY-MM 형식 월 필터."""
    return conn.execute(
        """
        SELECT o.*, GROUP_CONCAT(oi.frame || ' ' || COALESCE(oi.size,''), ', ') AS item_summary
        FROM orders o
        LEFT JOIN order_items oi ON oi.order_id = o.id
        WHERE o.sheet_date LIKE ? OR o.order_date LIKE ?
        GROUP BY o.id
        ORDER BY o.sheet_date, o.order_no
        """,
        (f"{year_month}%", f"%{year_month[-2:]}%"),
    ).fetchall()


def fetch_all_orders(conn: sqlite3.Connection, source_file: str | None = None) -> list[sqlite3.Row]:
    sql = "SELECT * FROM orders"
    params: tuple = ()
    if source_file:
        sql += " WHERE source_file=?"
        params = (source_file,)
    sql += " ORDER BY sheet_date, sheet_name, CAST(order_no AS INTEGER)"
    return conn.execute(sql, params).fetchall()


def fetch_order_items(conn: sqlite3.Connection, order_id: int) -> list[sqlite3.Row]:
    return conn.execute(
        "SELECT * FROM order_items WHERE order_id=? ORDER BY line_no",
        (order_id,),
    ).fetchall()
