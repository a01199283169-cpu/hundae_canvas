"""웹앱용 주문·결산·통계 서비스."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any

from src.config_loader import ROOT_DIR, load_config, resolve_path
from src.database import connect, fetch_order_items, init_db
from src.db_backend import has_image_clause, insert_returning_id, is_postgres


def image_to_url(image_file: str | None) -> str | None:
    """DB에 저장된 이미지 경로 → 웹에서 접근 가능한 URL."""
    if not image_file:
        return None
    path = Path(str(image_file))
    normalized = str(image_file).replace("\\", "/")

    if "uploads" in normalized:
        return f"/uploads/{path.name}"

    images_root = resolve_path("output/images")
    try:
        abs_path = path if path.is_absolute() else (ROOT_DIR / path)
        if abs_path.is_file():
            rel = abs_path.relative_to(images_root)
            return f"/order-images/{rel.as_posix()}"
    except ValueError:
        pass
    return None


def format_spec_summary(item: dict | None, item_count: int = 1) -> str:
    """엑셀 '상세내역' 형식으로 품목 스펙 요약."""
    if not item:
        return "-"
    parts: list[str] = []
    for key in ("frame", "size", "color", "plate", "acrylic", "hook"):
        val = item.get(key)
        if val:
            parts.append(str(val))
    w, h = item.get("width"), item.get("height")
    if w or h:
        parts.append(f"{w or ''}×{h or ''}")
    note = item.get("item_note")
    if note:
        parts.append(str(note))
    summary = " / ".join(parts) if parts else "-"
    if item_count and item_count > 1:
        summary += f" (+{item_count - 1}품목)"
    return summary


def format_payment(order: dict) -> str:
    """결제수단 표시."""
    if order.get("pay_card"):
        return "카드"
    if order.get("pay_transfer"):
        return "계좌이체"
    if order.get("pay_bank"):
        return "무통장"
    return "-"


def resolve_payment_status(order: dict) -> str:
    """
    결제여부 판정.
    - 명시 payment_status 우선
    - 입금일 있으면 완료
    - 카드 결제는 기본 완료
    - 그 외 미결
    """
    explicit = order.get("payment_status")
    if explicit in ("completed", "pending"):
        return explicit
    if str(order.get("deposit_date") or "").strip():
        return "completed"
    if order.get("pay_card"):
        return "completed"
    return "pending"


def payment_status_label(status: str) -> str:
    return {"completed": "완료", "pending": "미결"}.get(status, "-")


def _order_date_sql() -> str:
    return "COALESCE(o.sheet_date, o.order_date)"


def _build_settlement_where(
    *,
    period: str = "month",
    day: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    month: str | None = None,
    platform: str | None = None,
    pay_method: str | None = None,
) -> tuple[list[str], list[Any]]:
    """결산 조회 SQL WHERE 조건."""
    where: list[str] = []
    params: list[Any] = []
    date_expr = _order_date_sql()

    if period == "day" and day:
        where.append(f"{date_expr} = ?")
        params.append(day)
    elif period == "range" and date_from and date_to:
        where.append(f"{date_expr} BETWEEN ? AND ?")
        params.extend([date_from, date_to])
    elif month:
        if is_postgres():
            where.append(f"{date_expr}::text LIKE %s")
        else:
            where.append(f"{date_expr} LIKE ?")
        params.append(f"{month}%")

    if platform:
        where.append("o.platform = ?")
        params.append(platform)

    if pay_method == "card":
        where.append("o.pay_card IS NOT NULL AND o.pay_card != ''")
    elif pay_method == "transfer":
        where.append("o.pay_transfer IS NOT NULL AND o.pay_transfer != ''")
    elif pay_method == "bank":
        where.append("o.pay_bank IS NOT NULL AND o.pay_bank != ''")

    return where, params


def _order_display_date(order: dict) -> str:
    """주문 표시용 일자."""
    return str(order.get("sheet_date") or order.get("order_date") or "")


def _order_month(date_str: str) -> str:
    """YYYY-MM 추출."""
    return date_str[:7] if len(date_str) >= 7 else ""


def _sum_orders(orders: list[dict]) -> dict[str, Any]:
    """주문 목록 합계."""
    return {
        "order_count": len(orders),
        "sales": sum(float(o.get("sales") or 0) for o in orders),
        "deduct": sum(float(o.get("deduct") or 0) for o in orders),
        "ship": sum(float(o.get("ship") or 0) for o in orders),
        "total": sum(float(o.get("total") or 0) for o in orders),
        "completed_count": sum(
            1 for o in orders if o.get("payment_status") == "completed"
        ),
        "pending_count": sum(
            1 for o in orders if o.get("payment_status") != "completed"
        ),
    }


def _aggregate_by_date(orders: list[dict]) -> list[dict[str, Any]]:
    """일자별 집계 (플랫폼 무시)."""
    buckets: dict[str, list[dict]] = {}
    for o in orders:
        d = _order_display_date(o) or "(날짜없음)"
        buckets.setdefault(d, []).append(o)

    result: list[dict[str, Any]] = []
    for d in sorted(buckets.keys()):
        stats = _sum_orders(buckets[d])
        stats["date"] = d
        stats["month"] = _order_month(d) if d != "(날짜없음)" else ""
        result.append(stats)
    return result


def _month_label(month: str) -> str:
    """YYYY-MM → 표시용 라벨."""
    if not month or len(month) < 7:
        return "날짜없음"
    y, m = month.split("-")
    return f"{y}년 {int(m)}월"


def _sum_daily_stats(stats_list: list[dict[str, Any]]) -> dict[str, Any]:
    """일자별 집계 행들 → 합계."""
    return {
        "order_count": sum(int(s.get("order_count") or 0) for s in stats_list),
        "sales": sum(float(s.get("sales") or 0) for s in stats_list),
        "deduct": sum(float(s.get("deduct") or 0) for s in stats_list),
        "ship": sum(float(s.get("ship") or 0) for s in stats_list),
        "total": sum(float(s.get("total") or 0) for s in stats_list),
        "completed_count": sum(int(s.get("completed_count") or 0) for s in stats_list),
        "pending_count": sum(int(s.get("pending_count") or 0) for s in stats_list),
    }


def _build_daily_display_rows(daily_stats: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """일자별 합계 표 — 월 구분선 · 일자 행 · 월합계."""
    if not daily_stats:
        return []

    rows: list[dict[str, Any]] = []
    prev_month: str | None = None
    month_days: list[dict[str, Any]] = []

    def flush_month() -> None:
        nonlocal month_days, prev_month
        if not month_days or not prev_month:
            month_days = []
            return
        stats = _sum_daily_stats(month_days)
        stats["type"] = "month_subtotal"
        stats["month"] = prev_month
        stats["label"] = _month_label(prev_month)
        rows.append(stats)
        month_days = []

    for d in daily_stats:
        m = d.get("month") or ""
        if m and m != (prev_month or ""):
            flush_month()
            rows.append({
                "type": "month_header",
                "label": _month_label(m),
                "month": m,
            })
            prev_month = m

        rows.append({"type": "day", **d})
        month_days.append(d)
        if m:
            prev_month = m

    flush_month()
    return rows


def _build_sales_rows(orders: list[dict]) -> list[dict[str, Any]]:
    """
    매출 상세 표시용 행 목록.
    - month_header: 월 시작 구분
    - day_divider: 일자 변경 구분
    - order: 주문 1건
    - day_subtotal: 일 합계
    - month_subtotal: 월 합계
    """
    if not orders:
        return []

    sorted_orders = sorted(
        orders,
        key=lambda o: (
            _order_display_date(o) or "9999-99-99",
            int(o.get("order_no") or 0),
        ),
    )

    rows: list[dict[str, Any]] = []
    prev_month: str | None = None
    prev_date: str | None = None
    day_orders: list[dict] = []
    month_orders: list[dict] = []

    def flush_day() -> None:
        nonlocal day_orders, prev_date
        if not day_orders or prev_date is None:
            day_orders = []
            return
        stats = _sum_orders(day_orders)
        stats["type"] = "day_subtotal"
        stats["date"] = prev_date
        rows.append(stats)
        day_orders = []

    def append_month_subtotal(month_key: str, batch: list[dict]) -> None:
        if not batch or not month_key:
            return
        stats = _sum_orders(batch)
        stats["type"] = "month_subtotal"
        stats["month"] = month_key
        stats["label"] = _month_label(month_key)
        rows.append(stats)

    for o in sorted_orders:
        d = _order_display_date(o) or "(날짜없음)"
        month = _order_month(d) if d != "(날짜없음)" else ""

        # 월이 바뀌면: 일소계 → 월합계 → 새 월 구분선
        if month and month != (prev_month or ""):
            flush_day()
            if prev_month and month_orders:
                append_month_subtotal(prev_month, month_orders)
                month_orders = []
            rows.append({
                "type": "month_header",
                "label": _month_label(month),
                "month": month,
            })
            prev_date = None

        # 일자가 바뀌면: 일소계 → 일 구분선
        if prev_date is None or d != prev_date:
            if prev_date is not None:
                flush_day()
            rows.append({"type": "day_divider", "date": d, "label": d})

        rows.append({"type": "order", "order": o})
        day_orders.append(o)
        month_orders.append(o)
        prev_date = d
        if month:
            prev_month = month

    flush_day()
    if prev_month and month_orders:
        append_month_subtotal(prev_month, month_orders)

    return rows


def build_sales_rows(orders: list[dict]) -> list[dict[str, Any]]:
    """매출현황 리스트 행 생성 (템플릿·API용 공개 함수)."""
    return _build_sales_rows(orders)


def _aggregate_orders(orders: list[dict]) -> tuple[list[dict], dict[str, Any]]:
    """주문 목록 → 플랫폼별 집계 + 전체 합계."""
    platform_map: dict[str, dict[str, Any]] = {}
    grand = {
        "order_count": 0,
        "sales": 0.0,
        "deduct": 0.0,
        "ship": 0.0,
        "total": 0.0,
        "completed_count": 0,
        "pending_count": 0,
    }

    for o in orders:
        plat = o.get("platform") or "(미분류)"
        if plat not in platform_map:
            platform_map[plat] = {
                "platform": plat,
                "order_count": 0,
                "sales": 0.0,
                "deduct": 0.0,
                "ship": 0.0,
                "total": 0.0,
                "completed_count": 0,
                "pending_count": 0,
            }
        s = platform_map[plat]
        s["order_count"] += 1
        s["sales"] += float(o.get("sales") or 0)
        s["deduct"] += float(o.get("deduct") or 0)
        s["ship"] += float(o.get("ship") or 0)
        s["total"] += float(o.get("total") or 0)
        if o.get("payment_status") == "completed":
            s["completed_count"] += 1
        else:
            s["pending_count"] += 1

        grand["order_count"] += 1
        grand["sales"] += float(o.get("sales") or 0)
        grand["deduct"] += float(o.get("deduct") or 0)
        grand["ship"] += float(o.get("ship") or 0)
        grand["total"] += float(o.get("total") or 0)
        if o.get("payment_status") == "completed":
            grand["completed_count"] += 1
        else:
            grand["pending_count"] += 1

    platforms = sorted(platform_map.values(), key=lambda x: x["total"], reverse=True)
    return platforms, grand


def get_settlement_data(
    *,
    period: str = "month",
    day: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    month: str | None = None,
    platform: str | None = None,
    pay_method: str | None = None,
    payment_status: str | None = None,
) -> dict[str, Any]:
    """매출현황 조회 — 일자별 집계 + 주문 상세 + 합계."""
    conn = connect()
    where, params = _build_settlement_where(
        period=period,
        day=day,
        date_from=date_from,
        date_to=date_to,
        month=month,
        platform=platform,
        pay_method=pay_method,
    )
    where_sql = " AND ".join(where) if where else "1=1"
    date_expr = _order_date_sql()

    rows = conn.execute(
        f"""
        SELECT o.*,
               (SELECT COUNT(*) FROM order_items oi WHERE oi.order_id = o.id) AS item_count
        FROM orders o
        WHERE {where_sql}
        ORDER BY {date_expr}, CAST(o.order_no AS INTEGER)
        """,
        params,
    ).fetchall()

    orders = [dict(r) for r in rows]
    _attach_first_items(conn, orders)
    conn.close()

    for o in orders:
        o["payment_label"] = format_payment(o)
        o["payment_status"] = resolve_payment_status(o)
        o["payment_status_label"] = payment_status_label(o["payment_status"])
        o["display_date"] = _order_display_date(o) or "-"

    if payment_status in ("completed", "pending"):
        orders = [o for o in orders if o["payment_status"] == payment_status]

    daily_stats = _aggregate_by_date(orders)
    daily_rows = _build_daily_display_rows(daily_stats)
    grand = _sum_orders(orders)
    sales_rows = _build_sales_rows(orders)

    return {
        "period": period,
        "day": day,
        "date_from": date_from,
        "date_to": date_to,
        "month": month,
        "platform_filter": platform,
        "pay_method": pay_method,
        "payment_status": payment_status,
        "daily_stats": daily_stats,
        "daily_rows": daily_rows,
        "sales_rows": sales_rows,
        "grand": grand,
        "orders": orders,
    }


def format_expected_ship(order: dict) -> str:
    """예상출고(타입·운임비·수량) 표시."""
    parts: list[str] = []
    if order.get("expected_ship_type"):
        parts.append(str(order["expected_ship_type"]))
    if order.get("expected_freight"):
        parts.append(str(order["expected_freight"]))
    if order.get("expected_ship_qty"):
        parts.append(f"수량{order['expected_ship_qty']}")
    return " ".join(parts) or "-"


def _attach_first_items(conn, orders: list[dict]) -> None:
    """목록용 첫 품목·요약 필드 부착."""
    if not orders:
        return
    ids = [o["id"] for o in orders]
    placeholders = ",".join("?" * len(ids))
    rows = conn.execute(
        f"""
        SELECT oi.* FROM order_items oi
        INNER JOIN (
            SELECT order_id, MIN(line_no) AS min_ln
            FROM order_items WHERE order_id IN ({placeholders})
            GROUP BY order_id
        ) t ON oi.order_id = t.order_id AND oi.line_no = t.min_ln
        """,
        ids,
    ).fetchall()
    by_order = {r["order_id"]: dict(r) for r in rows}
    for o in orders:
        first = by_order.get(o["id"])
        cnt = o.get("item_count") or 0
        o["spec_summary"] = format_spec_summary(first, cnt)
        o["unit_price"] = first.get("unit_price") if first else None
        o["item_qty"] = first.get("qty") if first else None
        # 엑셀 상세내역 개별 열
        if first:
            o["item_frame"] = first.get("frame")
            o["item_size"] = first.get("size")
            o["item_width"] = first.get("width")
            o["item_height"] = first.get("height")
            o["item_color"] = first.get("color")
            o["item_plate"] = first.get("plate")
            o["item_acrylic"] = first.get("acrylic")
            o["item_hook"] = first.get("hook")
            o["item_note"] = first.get("item_note")
        else:
            for k in (
                "item_frame", "item_size", "item_width", "item_height",
                "item_color", "item_plate", "item_acrylic", "item_hook", "item_note",
            ):
                o[k] = None
        if cnt and cnt > 1:
            o["item_extra"] = f"+{cnt - 1}품목"
        else:
            o["item_extra"] = ""
        o["payment_label"] = format_payment(o)
        o["expected_ship_label"] = format_expected_ship(o)


def _attach_thumbnails(conn, orders: list[dict]) -> None:
    """주문 목록용 첫 이미지 썸네일 URL."""
    if not orders:
        return
    ids = [o["id"] for o in orders]
    placeholders = ",".join("?" * len(ids))
    rows = conn.execute(
        f"""
        SELECT order_id, image_file FROM order_images
        WHERE order_id IN ({placeholders})
        ORDER BY order_id, id
        """,
        ids,
    ).fetchall()
    first_path: dict[int, str] = {}
    for r in rows:
        oid = r["order_id"]
        if oid not in first_path:
            first_path[oid] = r["image_file"]
    for o in orders:
        path = first_path.get(o["id"])
        o["thumb_url"] = image_to_url(path)
        if path:
            o["has_image"] = 1


def get_dashboard_stats() -> dict[str, Any]:
    """대시보드 요약 통계 + 차트용 데이터."""
    conn = connect()
    date_expr = "COALESCE(NULLIF(sheet_date, ''), order_date)"

    total_orders = conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
    total_items = conn.execute("SELECT COUNT(*) FROM order_items").fetchone()[0]
    total_sales = conn.execute("SELECT COALESCE(SUM(sales),0) FROM orders").fetchone()[0]
    total_amount = conn.execute("SELECT COALESCE(SUM(total),0) FROM orders").fetchone()[0]
    with_image = conn.execute(
        f"SELECT COUNT(*) FROM orders WHERE {has_image_clause()}"
    ).fetchone()[0]
    platforms = conn.execute(
        "SELECT COUNT(DISTINCT platform) FROM orders WHERE platform IS NOT NULL"
    ).fetchone()[0]
    today = date.today().isoformat()
    today_orders = conn.execute(
        f"SELECT COUNT(*) FROM orders WHERE {date_expr} = ?",
        (today,),
    ).fetchone()[0]

    # 일별 추이 (최근 14일, ISO 날짜만)
    daily_rows = conn.execute(
        f"""
        SELECT {date_expr} AS d, COUNT(*) AS cnt, COALESCE(SUM(total), 0) AS amt
        FROM orders
        WHERE {date_expr} GLOB '????-??-??'
        GROUP BY d
        ORDER BY d DESC
        LIMIT 14
        """
    ).fetchall()
    daily_rows = list(reversed(daily_rows))

    # 구분(플랫폼)별
    platform_rows = conn.execute(
        """
        SELECT COALESCE(platform, '(미분류)') AS p,
               COUNT(*) AS cnt,
               COALESCE(SUM(total), 0) AS amt
        FROM orders
        GROUP BY p
        ORDER BY cnt DESC
        LIMIT 8
        """
    ).fetchall()

    # 월별 매출
    month_rows = conn.execute(
        f"""
        SELECT substr({date_expr}, 1, 7) AS m,
               COUNT(*) AS cnt,
               COALESCE(SUM(total), 0) AS amt
        FROM orders
        WHERE {date_expr} GLOB '????-??-??'
        GROUP BY m
        ORDER BY m DESC
        LIMIT 6
        """
    ).fetchall()
    month_rows = list(reversed(month_rows))

    # 결제 완료 / 미결
    pay_rows = conn.execute(
        """
        SELECT payment_status, deposit_date, pay_card FROM orders
        """
    ).fetchall()
    pay_completed = 0
    pay_pending = 0
    for r in pay_rows:
        row = dict(r)
        if resolve_payment_status(row) == "completed":
            pay_completed += 1
        else:
            pay_pending += 1

    conn.close()

    def _fmt_day(d: str) -> str:
        return d[5:] if len(d) >= 10 else d

    def _fmt_month(m: str) -> str:
        if len(m) >= 7:
            return f"{int(m[5:7])}월"
        return m

    image_pct = round(with_image / total_orders * 100) if total_orders else 0

    return {
        "total_orders": total_orders,
        "total_items": total_items,
        "total_sales": total_sales,
        "total_amount": total_amount,
        "with_image": with_image,
        "platform_count": platforms,
        "today_orders": today_orders,
        "image_pct": image_pct,
        "charts": {
            "daily_labels": [_fmt_day(r["d"]) for r in daily_rows],
            "daily_orders": [int(r["cnt"]) for r in daily_rows],
            "daily_sales": [float(r["amt"]) for r in daily_rows],
            "platform_labels": [str(r["p"]) for r in platform_rows],
            "platform_counts": [int(r["cnt"]) for r in platform_rows],
            "platform_sales": [float(r["amt"]) for r in platform_rows],
            "month_labels": [_fmt_month(r["m"]) for r in month_rows],
            "month_orders": [int(r["cnt"]) for r in month_rows],
            "month_sales": [float(r["amt"]) for r in month_rows],
            "payment_labels": ["결제완료", "미결"],
            "payment_counts": [pay_completed, pay_pending],
        },
    }


def list_orders(
    *,
    period: str = "month",
    month: str | None = None,
    day: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    search: str | None = None,
    incomplete_only: bool = False,
    limit: int = 500,
    offset: int = 0,
    expand_items: bool = True,
) -> tuple[list[dict], int]:
    """주문 목록 — expand_items=True 이면 엑셀처럼 품목별 1행."""
    conn = connect()
    date_where, params = _build_settlement_where(
        period=period,
        day=day,
        date_from=date_from,
        date_to=date_to,
        month=month,
    )
    where = date_where if date_where else ["1=1"]

    if search:
        where.append("(o.customer LIKE ? OR o.phone LIKE ? OR o.order_no LIKE ?)")
        params.extend([f"%{search}%"] * 3)

    where_sql = " AND ".join(where)

    if expand_items:
        rows = conn.execute(
            f"""
            SELECT o.*,
                   oi.line_no,
                   oi.frame AS item_frame,
                   oi.size AS item_size,
                   oi.width AS item_width,
                   oi.height AS item_height,
                   oi.color AS item_color,
                   oi.plate AS item_plate,
                   oi.acrylic AS item_acrylic,
                   oi.hook AS item_hook,
                   oi.item_note,
                   oi.qty AS item_qty,
                   oi.unit_price,
                   (SELECT COUNT(*) FROM order_items x WHERE x.order_id = o.id) AS item_count
            FROM orders o
            INNER JOIN order_items oi ON oi.order_id = o.id
            WHERE {where_sql}
            ORDER BY o.sheet_date DESC, CAST(o.order_no AS INTEGER) DESC, oi.line_no
            """,
            params,
        ).fetchall()
        orders = [dict(r) for r in rows]
        for o in orders:
            o["is_first_line"] = int(o.get("line_no") or 1) == 1
            o["line_sales"] = (float(o.get("unit_price") or 0) * float(o.get("item_qty") or 0)) or None
            o["item_extra"] = ""
            o["payment_label"] = format_payment(o)
            o["expected_ship_label"] = format_expected_ship(o)
    else:
        rows = conn.execute(
            f"""
            SELECT o.*,
                   (SELECT COUNT(*) FROM order_items oi WHERE oi.order_id=o.id) AS item_count
            FROM orders o
            WHERE {where_sql}
            ORDER BY o.sheet_date DESC, CAST(o.order_no AS INTEGER) DESC
            """,
            params,
        ).fetchall()
        orders = [dict(r) for r in rows]
        _attach_first_items(conn, orders)

    _attach_thumbnails(conn, orders)

    # 주문 단위 완료 여부
    seen: dict[int, bool] = {}
    for o in orders:
        oid = o["id"]
        if oid not in seen:
            missing = order_missing_fields(o)
            seen[oid] = len(missing) == 0
            o["missing"] = missing
        o["is_complete"] = seen[oid]

    if incomplete_only:
        incomplete_ids = {oid for oid, ok in seen.items() if not ok}
        orders = [o for o in orders if o["id"] in incomplete_ids]

    order_count = len({o["id"] for o in orders})
    orders = orders[offset : offset + limit]
    conn.close()
    return orders, order_count


def _calc_total(sales: float, ship: float, deduct: float) -> float:
    """합계 = 판매가 + 택배비 + 공제액(음수 가능)."""
    return float(sales or 0) + float(ship or 0) + float(deduct or 0)


def get_order(order_id: int) -> dict | None:
    """주문 상세 + 품목."""
    conn = connect()
    row = conn.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
    if not row:
        conn.close()
        return None
    order = dict(row)
    order["line_items"] = [dict(i) for i in fetch_order_items(conn, order_id)]
    order["images"] = [
        dict(i)
        for i in conn.execute(
            "SELECT * FROM order_images WHERE order_id=?", (order_id,)
        ).fetchall()
    ]
    for img in order["images"]:
        img["url"] = image_to_url(img.get("image_file"))
    order["thumb_url"] = image_to_url(order["images"][0]["image_file"]) if order["images"] else None
    order["missing"] = order_missing_fields(order)
    order["is_complete"] = len(order["missing"]) == 0
    order["resolved_payment_status"] = resolve_payment_status(order)
    order["payment_status_label"] = payment_status_label(order["resolved_payment_status"])
    order["payment_label"] = format_payment(order)
    conn.close()
    return order


def _next_order_no(conn, sheet_date: str) -> str:
    """당일 주문번호 자동 부여."""
    row = conn.execute(
        """
        SELECT MAX(CAST(order_no AS INTEGER)) FROM orders
        WHERE source_file='web' AND sheet_date=?
        """,
        (sheet_date,),
    ).fetchone()
    n = (row[0] or 0) + 1
    return str(n)


def order_missing_fields(order: dict) -> list[str]:
    """주문 필수 정보 누락 항목 반환."""
    missing = []
    if not str(order.get("customer") or "").strip():
        missing.append("주문자")
    if not str(order.get("phone") or "").strip():
        missing.append("연락처")
    if not str(order.get("address") or "").strip():
        missing.append("주소")
    if not str(order.get("platform") or "").strip():
        missing.append("플랫폼")
    return missing


def is_order_complete(order: dict) -> bool:
    return len(order_missing_fields(order)) == 0


def validate_order_info(data: dict) -> list[str]:
    """주문자·배송 정보만 검증."""
    errors = order_missing_fields(data)
    if not str(data.get("pay_method") or "").strip() and not any(
        data.get(k) for k in ("pay_card", "pay_transfer", "pay_bank")
    ):
        errors.append("결제수단")
    return errors


def validate_new_order(data: dict, items: list[dict]) -> list[str]:
    """신규 주문 저장 전 검증. 오류 메시지 목록 반환."""
    errors = order_missing_fields(data)
    if not str(data.get("pay_method") or "").strip() and not any(
        data.get(k) for k in ("pay_card", "pay_transfer", "pay_bank")
    ):
        errors.append("결제수단")
    for it in items:
        if not str(it.get("frame") or "").strip():
            errors.append("프레임 종류")
            break
        if not it.get("qty") or float(it.get("qty") or 0) <= 0:
            errors.append("수량")
            break
        if not it.get("unit_price") and it.get("unit_price") != 0:
            errors.append("단가")
            break
    return errors


def update_order_info(order_id: int, data: dict) -> None:
    """주문 정보·금액·예상출고 수정."""
    conn = connect()
    sales = data.get("sales")
    if sales is None:
        sales = conn.execute("SELECT sales FROM orders WHERE id=?", (order_id,)).fetchone()[0]
    ship = float(data.get("ship") or 0)
    deduct = float(data.get("deduct") or 0)
    total = _calc_total(float(sales or 0), ship, deduct)

    conn.execute(
        """
        UPDATE orders SET
            platform=?, customer=?, phone=?, address=?,
            sales=?, ship=?, deduct=?, total=?,
            order_qty=?, remark=?,
            pay_card=?, pay_transfer=?, pay_bank=?,
            payment_status=?, deposit_date=?,
            expected_ship_type=?, expected_freight=?, expected_ship_qty=?
        WHERE id=?
        """,
        (
            data.get("platform"),
            data.get("customer"),
            data.get("phone"),
            data.get("address"),
            sales,
            ship,
            deduct,
            total,
            data.get("order_qty"),
            data.get("remark"),
            data.get("pay_card"),
            data.get("pay_transfer"),
            data.get("pay_bank"),
            data.get("payment_status"),
            data.get("deposit_date"),
            data.get("expected_ship_type"),
            data.get("expected_freight"),
            data.get("expected_ship_qty"),
            order_id,
        ),
    )
    conn.commit()
    conn.close()


def count_incomplete_orders() -> int:
    conn = connect()
    rows = conn.execute("SELECT id, customer, phone, address, platform FROM orders").fetchall()
    conn.close()
    return sum(1 for r in rows if not is_order_complete(dict(r)))


def create_order_web(data: dict, items: list[dict]) -> int:
    """웹에서 신규 주문 등록."""
    conn = connect()
    sheet_date = data.get("sheet_date") or date.today().isoformat()
    order_no = _next_order_no(conn, sheet_date)

    sales = data.get("sales")
    if sales is None:
        sales = sum((float(it.get("qty") or 0) * float(it.get("unit_price") or 0)) for it in items)
    order_qty = data.get("order_qty")
    if order_qty is None:
        order_qty = sum(float(it.get("qty") or 0) for it in items)
    total = _calc_total(float(sales or 0), float(data.get("ship") or 0), float(data.get("deduct") or 0))

    order_id = insert_returning_id(
        conn,
        """
        INSERT INTO orders (
            import_id, source_file, sheet_name, sheet_date, order_no, order_date,
            platform, file_ref, customer, phone, address,
            sales, deduct, ship, total,
            pay_card, pay_transfer, pay_bank, remark, has_image,
            order_qty, expected_ship_type, expected_freight, expected_ship_qty,
            payment_status, deposit_date
        ) VALUES (NULL, 'web', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?, ?, ?, ?, ?)
        """,
        (
            sheet_date,
            sheet_date,
            order_no,
            data.get("order_date") or sheet_date,
            data.get("platform"),
            data.get("file_ref"),
            data.get("customer"),
            data.get("phone"),
            data.get("address"),
            sales,
            data.get("deduct") or 0,
            data.get("ship") or 0,
            total,
            data.get("pay_card"),
            data.get("pay_transfer"),
            data.get("pay_bank"),
            data.get("remark"),
            order_qty,
            data.get("expected_ship_type"),
            data.get("expected_freight"),
            data.get("expected_ship_qty"),
            data.get("payment_status"),
            data.get("deposit_date"),
        ),
    )

    for idx, it in enumerate(items, start=1):
        conn.execute(
            """
            INSERT INTO order_items (
                order_id, line_no, frame, size, width, height, color,
                plate, acrylic, hook, item_note, qty, unit_price
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                order_id, idx,
                it.get("frame"), it.get("size"),
                it.get("width"), it.get("height"), it.get("color"),
                it.get("plate"), it.get("acrylic"), it.get("hook"),
                it.get("item_note"), it.get("qty"), it.get("unit_price"),
            ),
        )
    conn.commit()
    conn.close()
    return order_id


def delete_order(order_id: int) -> bool:
    conn = connect()
    cur = conn.execute("DELETE FROM orders WHERE id=?", (order_id,))
    conn.commit()
    conn.close()
    return cur.rowcount > 0


def get_platform_list() -> list[str]:
    conn = connect()
    rows = conn.execute(
        "SELECT DISTINCT platform FROM orders WHERE platform IS NOT NULL ORDER BY platform"
    ).fetchall()
    conn.close()
    config = load_config()
    aliases = set(config.get("platform_aliases", {}).values())
    db_platforms = {r[0] for r in rows}
    return sorted(aliases | db_platforms)


def get_settlement_summary(month: str | None = None) -> dict[str, Any]:
    """일자별 매출 집계 (하위 호환)."""
    data = get_settlement_data(period="month", month=month)
    return {"month": month, "daily_stats": data["daily_stats"], "grand": data["grand"]}


def get_production_list(sheet_date: str | None = None) -> list[dict]:
    """생산지시용 품목 목록."""
    conn = connect()
    where = "1=1"
    params: list = []
    if sheet_date:
        where = "o.sheet_date = ?"
        params = [sheet_date]

    rows = conn.execute(
        f"""
        SELECT o.id AS order_id, o.sheet_date, o.order_no, o.platform,
               o.customer, o.phone, o.address, o.has_image, o.remark,
               oi.frame, oi.size, oi.width, oi.height, oi.color,
               oi.plate, oi.acrylic, oi.hook, oi.qty, oi.unit_price, oi.item_note
        FROM orders o
        JOIN order_items oi ON oi.order_id = o.id
        WHERE {where}
        ORDER BY o.sheet_date, CAST(o.order_no AS INTEGER), oi.line_no
        """,
        params,
    ).fetchall()
    conn.close()
    items = []
    for r in rows:
        d = dict(r)
        d["missing"] = order_missing_fields(d)
        d["is_complete"] = len(d["missing"]) == 0
        items.append(d)
    return items


def save_uploaded_image(
    order_id: int,
    file_path: Path,
    original_name: str | None = None,
) -> None:
    """주문에 이미지 파일 연결."""
    conn = connect()
    mapped_val = True if is_postgres() else 1
    hi_val = True if is_postgres() else 1
    conn.execute(
        """
        INSERT INTO order_images (order_id, source_file, sheet_name, excel_row, image_file, mapped)
        VALUES (?, 'web', 'upload', 0, ?, ?)
        """,
        (order_id, str(file_path), mapped_val),
    )
    conn.execute(
        "UPDATE orders SET has_image=?, file_ref=? WHERE id=?",
        (hi_val, original_name or file_path.name, order_id),
    )
    conn.commit()
    conn.close()


def get_upload_dir() -> Path:
    from src.settings import get_upload_dir_override

    override = get_upload_dir_override()
    if override:
        override.mkdir(parents=True, exist_ok=True)
        return override
    p = resolve_path("output/images/uploads")
    p.mkdir(parents=True, exist_ok=True)
    return p
