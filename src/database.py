"""주문 DB 스키마 및 CRUD — SQLite / Supabase(Postgres) 공통."""

from __future__ import annotations

from datetime import datetime
from typing import Any

from src.db_backend import (
    ConnectionWrapper,
    connect,
    init_postgres_schema,
    insert_returning_id,
    is_postgres,
    order_no_order,
)
from src.db_backend import Row  # noqa: F401 — 하위 호환


def get_db_path():
    """로컬 SQLite 경로 (Postgres 사용 시 None에 가까움)."""
    from src.settings import get_database_path
    from src.config_loader import load_config

    if is_postgres():
        return None
    return get_database_path(load_config())


def init_db(conn: ConnectionWrapper | None = None) -> None:
    """orders, order_items, import_logs, order_images 테이블 생성."""
    own_conn = conn is None
    if own_conn:
        conn = connect()

    if is_postgres():
        init_postgres_schema(conn)
    else:
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


def _migrate_columns(conn: ConnectionWrapper) -> None:
    """SQLite — 기존 DB에 엑셀 양식 확장 컬럼 추가."""
    existing = {row[1] for row in conn.execute("PRAGMA table_info(orders)").fetchall()}
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
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_orders_deposit_date ON orders(deposit_date)"
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_orders_payment_status ON orders(payment_status)"
    )


def insert_import_log(
    conn: ConnectionWrapper,
    source_file: str,
    sheet_name: str,
    order_count: int,
    item_count: int,
) -> int:
    return insert_returning_id(
        conn,
        """
        INSERT INTO import_logs (source_file, sheet_name, imported_at, order_count, item_count)
        VALUES (?, ?, ?, ?, ?)
        """,
        (source_file, sheet_name, datetime.now().isoformat(), order_count, item_count),
    )


def upsert_order(conn: ConnectionWrapper, order: dict[str, Any], import_id: int) -> int:
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


def mark_order_has_image(conn: ConnectionWrapper, order_id: int) -> None:
    val = True if is_postgres() else 1
    conn.execute("UPDATE orders SET has_image=? WHERE id=?", (val, order_id))
    conn.commit()


def insert_order_image(
    conn: ConnectionWrapper,
    *,
    order_id: int | None,
    source_file: str,
    sheet_name: str,
    excel_row: int,
    image_file: str,
    mapped: bool,
) -> None:
    mapped_val = mapped if is_postgres() else (1 if mapped else 0)
    conn.execute(
        """
        INSERT INTO order_images (order_id, source_file, sheet_name, excel_row, image_file, mapped)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (order_id, source_file, sheet_name, excel_row, image_file, mapped_val),
    )
    conn.commit()


def fetch_orders_for_month(conn: ConnectionWrapper, year_month: str) -> list[Row]:
    """YYYY-MM 형식 월 필터."""
    agg = "STRING_AGG(oi.frame || ' ' || COALESCE(oi.size,''), ', ')" if is_postgres() else "GROUP_CONCAT(oi.frame || ' ' || COALESCE(oi.size,''), ', ')"
    like_op = f"o.sheet_date::text LIKE %s" if is_postgres() else "o.sheet_date LIKE ?"
    return conn.execute(
        f"""
        SELECT o.*, {agg} AS item_summary
        FROM orders o
        LEFT JOIN order_items oi ON oi.order_id = o.id
        WHERE {like_op} OR o.order_date LIKE ?
        GROUP BY o.id
        ORDER BY o.sheet_date, o.order_no
        """,
        (f"{year_month}%", f"%{year_month[-2:]}%"),
    ).fetchall()


def fetch_all_orders(conn: ConnectionWrapper, source_file: str | None = None) -> list[Row]:
    sql = f"SELECT * FROM orders"
    params: tuple = ()
    if source_file:
        sql += " WHERE source_file=?"
        params = (source_file,)
    sql += f" ORDER BY sheet_date, sheet_name, {order_no_order()}"
    return conn.execute(sql, params).fetchall()


def fetch_order_items(conn: ConnectionWrapper, order_id: int) -> list[Row]:
    return conn.execute(
        "SELECT * FROM order_items WHERE order_id=? ORDER BY line_no",
        (order_id,),
    ).fetchall()
