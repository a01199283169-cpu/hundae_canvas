"""모닝프레임 웹앱 - FastAPI."""

from __future__ import annotations

import shutil
import uuid
from contextlib import asynccontextmanager
from datetime import date, datetime
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from jinja2 import Environment, FileSystemLoader, select_autoescape

from src.settings import is_production, load_dotenv
from src.config_loader import ROOT_DIR, ensure_dirs, load_config, resolve_path
from src.database import init_db
from src.main import find_latest_xlsx
from src.parser import import_to_db
from src.price_validator import create_price_catalog_template, validate_prices
from src.production_export import export_production_sheet
from src.settlement_export import export_monthly_settlement
from src.web_service import (
    build_sales_rows,
    count_incomplete_orders,
    create_order_web,
    delete_order,
    get_dashboard_stats,
    get_order,
    get_platform_list,
    get_production_list,
    get_settlement_data,
    get_upload_dir,
    list_orders,
    save_uploaded_image,
    update_order_info,
    validate_new_order,
    validate_order_info,
)

WEB_DIR = ROOT_DIR / "web"

_db_ready = False
_db_error: str | None = None


def _bootstrap_db() -> None:
    """DB·폴더 초기화 — import 시점이 아니라 첫 요청 전 (Render 기동 대기 방지)."""
    global _db_ready, _db_error
    if _db_ready:
        return
    if _db_error:
        raise RuntimeError(_db_error)
    try:
        init_db()
        ensure_dirs()
        catalog = resolve_path(load_config()["paths"]["price_catalog"])
        if not catalog.exists():
            create_price_catalog_template()
        _db_ready = True
    except Exception as exc:
        _db_error = str(exc)
        raise


@asynccontextmanager
async def _lifespan(_app: FastAPI):
    load_dotenv()
    yield


app = FastAPI(title="모닝프레임 주문관리", version="1.0.0", lifespan=_lifespan)

load_dotenv()

_jinja = Environment(
    loader=FileSystemLoader(str(WEB_DIR / "templates")),
    autoescape=select_autoescape(["html"]),
    cache_size=400 if is_production() else 0,
)


def render(request: Request, template: str, context: dict) -> HTMLResponse:
    # Jinja2에서 dict 키가 메서드명(items 등)과 충돌하지 않도록 기본값 보강
    ctx = {"incomplete_count": 0, **context, "request": request}
    html = _jinja.get_template(template).render(**ctx)
    return HTMLResponse(html)


def render_print(template: str, context: dict) -> HTMLResponse:
    """인쇄 전용 레이아웃 — 사이드바 없음."""
    ctx = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        **context,
    }
    html = _jinja.get_template(template).render(**ctx)
    return HTMLResponse(html)


def _period_subtitle(
    *,
    period: str = "month",
    month: str | None = None,
    day: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
) -> str:
    """인쇄 문서 상단 조회 조건 요약."""
    if period == "day" and day:
        return f"일자 {day}"
    if period == "range" and date_from and date_to:
        return f"기간 {date_from} ~ {date_to}"
    if month:
        return f"{month}"
    return "전체"

app.mount("/static", StaticFiles(directory=str(WEB_DIR / "static")), name="static")
_uploads = get_upload_dir()
app.mount("/uploads", StaticFiles(directory=str(_uploads)), name="uploads")
_order_images = ROOT_DIR / "output" / "images"
_order_images.mkdir(parents=True, exist_ok=True)
app.mount("/order-images", StaticFiles(directory=str(_order_images)), name="order-images")


def _persist_upload(order_id: int, file: UploadFile) -> None:
    """업로드 파일 저장 후 주문에 연결."""
    ext = Path(file.filename or "img.png").suffix or ".png"
    dest = get_upload_dir() / f"order_{order_id}_{uuid.uuid4().hex[:8]}{ext}"
    with dest.open("wb") as f:
        shutil.copyfileobj(file.file, f)
    save_uploaded_image(order_id, dest, file.filename)


@app.get("/healthz")
async def healthz():
    """Render 헬스체크 — DB 없이 즉시 응답."""
    return {"status": "ok"}


@app.middleware("http")
async def bootstrap_middleware(request: Request, call_next):
    if request.url.path not in ("/healthz", "/health"):
        _bootstrap_db()
    return await call_next(request)


@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    stats = get_dashboard_stats()
    incomplete = count_incomplete_orders()
    return render(
        request, "dashboard.html",
        {"stats": stats, "page": "dashboard", "incomplete_count": incomplete},
    )


@app.get("/orders", response_class=HTMLResponse)
async def orders_page(
    request: Request,
    period: str = "month",
    month: str = "",
    day: str = "",
    date_from: str = "",
    date_to: str = "",
    search: str = "",
    incomplete: str = "",
):
    if not month and period == "month":
        month = date.today().strftime("%Y-%m")
    orders, total = list_orders(
        period=period,
        month=month or None,
        day=day or None,
        date_from=date_from or None,
        date_to=date_to or None,
        search=search or None,
        incomplete_only=incomplete == "1",
    )
    return render(
        request,
        "orders.html",
        {
            "page": "orders",
            "orders": orders,
            "total": total,
            "row_count": len(orders),
            "period": period,
            "month": month,
            "day": day,
            "date_from": date_from,
            "date_to": date_to,
            "search": search,
            "incomplete_only": incomplete == "1",
            "incomplete_count": count_incomplete_orders(),
            "query_string": request.url.query,
        },
    )


@app.get("/orders/new", response_class=HTMLResponse)
async def order_new_page(request: Request, error: str = ""):
    return render(
        request,
        "order_form.html",
        {
            "page": "orders",
            "platforms": get_platform_list(),
            "today": date.today().isoformat(),
            "error": error,
            "form": {},
        },
    )


@app.post("/orders/new")
async def order_create(
    request: Request,
    sheet_date: str = Form(...),
    platform: str = Form(...),
    customer: str = Form(""),
    phone: str = Form(""),
    address: str = Form(""),
    sales: float = Form(0),
    order_qty: float = Form(0),
    ship: float = Form(0),
    deduct: float = Form(0),
    remark: str = Form(""),
    pay_method: str = Form(""),
    expected_ship_type: str = Form(""),
    expected_freight: str = Form(""),
    expected_ship_qty: float = Form(0),
    frame: str = Form(...),
    size: str = Form(""),
    width: float = Form(0),
    height: float = Form(0),
    color: str = Form(""),
    plate: str = Form(""),
    acrylic: str = Form(""),
    hook: str = Form(""),
    item_note: str = Form(""),
    qty: float = Form(1),
    unit_price: float = Form(0),
    image_file: UploadFile | None = File(None),
):
    pay = {"pay_card": None, "pay_transfer": None, "pay_bank": None}
    if pay_method == "card":
        pay["pay_card"] = "O"
    elif pay_method == "transfer":
        pay["pay_transfer"] = "O"
    elif pay_method == "bank":
        pay["pay_bank"] = "O"

    form_data = {
        "sheet_date": sheet_date, "platform": platform, "customer": customer,
        "phone": phone, "address": address, "sales": sales, "order_qty": order_qty,
        "ship": ship, "deduct": deduct, "remark": remark, "pay_method": pay_method,
        "expected_ship_type": expected_ship_type, "expected_freight": expected_freight,
        "expected_ship_qty": expected_ship_qty or None,
        "frame": frame, "size": size, "width": width, "height": height,
        "color": color, "plate": plate, "acrylic": acrylic, "hook": hook,
        "item_note": item_note, "qty": qty, "unit_price": unit_price, **pay,
    }
    items = [{
        "frame": frame, "size": size, "width": width or None,
        "height": height or None, "color": color, "plate": plate,
        "acrylic": acrylic, "hook": hook, "item_note": item_note or None,
        "qty": qty, "unit_price": unit_price,
    }]

    errors = validate_new_order(form_data, items)
    if errors:
        msg = "다음 필수 항목을 입력해 주세요: " + ", ".join(dict.fromkeys(errors))
        return render(request, "order_form.html", {
            "page": "orders", "platforms": get_platform_list(),
            "today": date.today().isoformat(), "error": msg, "form": form_data,
        })

    order_id = create_order_web(
        {
            "sheet_date": sheet_date,
            "platform": platform,
            "customer": customer.strip(),
            "phone": phone.strip(),
            "address": address.strip(),
            "sales": sales or None,
            "order_qty": order_qty or None,
            "ship": ship,
            "deduct": deduct,
            "remark": remark,
            "expected_ship_type": expected_ship_type or None,
            "expected_freight": expected_freight or None,
            "expected_ship_qty": expected_ship_qty or None,
            **pay,
        },
        items,
    )
    if image_file and image_file.filename:
        _persist_upload(order_id, image_file)
    return RedirectResponse(f"/orders/{order_id}", status_code=303)


@app.get("/orders/{order_id}", response_class=HTMLResponse)
async def order_detail(request: Request, order_id: int, saved: str = ""):
    order = get_order(order_id)
    if not order:
        return RedirectResponse("/orders", status_code=303)
    return render(request, "order_detail.html", {
        "page": "orders",
        "order": order,
        "platforms": get_platform_list(),
        "saved": saved == "1",
    })


@app.post("/orders/{order_id}/edit")
async def order_edit(
    order_id: int,
    platform: str = Form(...),
    customer: str = Form(...),
    phone: str = Form(...),
    address: str = Form(...),
    sales: float = Form(0),
    order_qty: float = Form(0),
    ship: float = Form(0),
    deduct: float = Form(0),
    remark: str = Form(""),
    pay_method: str = Form(...),
    payment_status: str = Form(""),
    deposit_date: str = Form(""),
    expected_ship_type: str = Form(""),
    expected_freight: str = Form(""),
    expected_ship_qty: float = Form(0),
):
    pay = {"pay_card": None, "pay_transfer": None, "pay_bank": None}
    if pay_method == "card":
        pay["pay_card"] = "O"
    elif pay_method == "transfer":
        pay["pay_transfer"] = "O"
    elif pay_method == "bank":
        pay["pay_bank"] = "O"

    data = {
        "platform": platform.strip(),
        "customer": customer.strip(),
        "phone": phone.strip(),
        "address": address.strip(),
        "sales": sales,
        "order_qty": order_qty or None,
        "ship": ship,
        "deduct": deduct,
        "remark": remark,
        "expected_ship_type": expected_ship_type or None,
        "expected_freight": expected_freight or None,
        "expected_ship_qty": expected_ship_qty or None,
        "payment_status": payment_status or None,
        "deposit_date": deposit_date or None,
        **pay,
    }
    missing = validate_order_info({**data, "pay_method": pay_method})
    if missing:
        order = get_order(order_id)
        order["missing"] = missing
        return render(request, "order_detail.html", {
            "page": "orders", "order": order,
            "platforms": get_platform_list(), "saved": False,
        })

    update_order_info(order_id, data)
    return RedirectResponse(f"/orders/{order_id}?saved=1", status_code=303)


@app.post("/orders/{order_id}/delete")
async def order_delete(order_id: int):
    delete_order(order_id)
    return RedirectResponse("/orders", status_code=303)


@app.post("/orders/{order_id}/upload-image")
async def upload_image(order_id: int, file: UploadFile = File(...)):
    _persist_upload(order_id, file)
    return RedirectResponse(f"/orders/{order_id}", status_code=303)


@app.get("/production", response_class=HTMLResponse)
async def production_page(request: Request, date_filter: str = ""):
    items = get_production_list(date_filter or None)
    dates = sorted({i["sheet_date"] for i in items if i.get("sheet_date")}, reverse=True)
    return render(
        request,
        "production.html",
        {
            "page": "production",
            "items": items,
            "dates": dates,
            "date_filter": date_filter,
            "incomplete_count": count_incomplete_orders(),
        },
    )


@app.get("/production/download")
async def production_download():
    path = export_production_sheet()
    return FileResponse(path, filename=path.name)


@app.get("/output", response_class=HTMLResponse)
async def output_page(request: Request):
    """출력·다운로드 허브."""
    items = get_production_list(None)
    prod_dates = sorted(
        {i["sheet_date"] for i in items if i.get("sheet_date")},
        reverse=True,
    )
    return render(
        request,
        "output.html",
        {
            "page": "output",
            "prod_dates": prod_dates,
            "default_month": date.today().strftime("%Y-%m"),
        },
    )


@app.get("/output/print/production", response_class=HTMLResponse)
async def print_production(date_filter: str = ""):
    items = get_production_list(date_filter or None)
    return render_print(
        "print_production.html",
        {"items": items, "date_filter": date_filter},
    )


@app.get("/output/print/orders", response_class=HTMLResponse)
async def print_orders(
    period: str = "month",
    month: str = "",
    day: str = "",
    date_from: str = "",
    date_to: str = "",
    search: str = "",
    incomplete: str = "",
):
    if not month and period == "month":
        month = date.today().strftime("%Y-%m")
    orders, total = list_orders(
        period=period,
        month=month or None,
        day=day or None,
        date_from=date_from or None,
        date_to=date_to or None,
        search=search or None,
        incomplete_only=incomplete == "1",
    )
    subtitle = _period_subtitle(
        period=period,
        month=month or None,
        day=day or None,
        date_from=date_from or None,
        date_to=date_to or None,
    )
    if search:
        subtitle += f" · 검색「{search}」"
    if incomplete == "1":
        subtitle += " · 누락만"
    subtitle += f" · 주문 {total}건 / {len(orders)}행"
    return render_print(
        "print_orders.html",
        {"orders": orders, "subtitle": subtitle},
    )


@app.get("/output/print/settlement", response_class=HTMLResponse)
async def print_settlement(
    period: str = "month",
    month: str = "",
    day: str = "",
    date_from: str = "",
    date_to: str = "",
    pay_method: str = "",
    payment_status: str = "",
):
    if not month and period == "month":
        month = date.today().strftime("%Y-%m")
    data = get_settlement_data(
        period=period,
        month=month or None,
        day=day or None,
        date_from=date_from or None,
        date_to=date_to or None,
        pay_method=pay_method or None,
        payment_status=payment_status or None,
    )
    if not data.get("sales_rows") and data.get("orders"):
        data["sales_rows"] = build_sales_rows(data["orders"])
    elif "sales_rows" not in data:
        data["sales_rows"] = []
    subtitle = _period_subtitle(
        period=period,
        month=month or None,
        day=day or None,
        date_from=date_from or None,
        date_to=date_to or None,
    )
    subtitle += f" · {data['grand']['order_count']}건"
    return render_print(
        "print_settlement.html",
        {"data": data, "subtitle": subtitle},
    )


@app.get("/settlement", response_class=HTMLResponse)
async def settlement_page(
    request: Request,
    period: str = "month",
    month: str = "",
    day: str = "",
    date_from: str = "",
    date_to: str = "",
    pay_method: str = "",
    payment_status: str = "",
):
    if not month and period == "month":
        month = date.today().strftime("%Y-%m")
    data = get_settlement_data(
        period=period,
        month=month or None,
        day=day or None,
        date_from=date_from or None,
        date_to=date_to or None,
        pay_method=pay_method or None,
        payment_status=payment_status or None,
    )
    # sales_rows 누락 시(구버전 프로세스) 주문 목록으로 재구성
    if not data.get("sales_rows") and data.get("orders"):
        data["sales_rows"] = build_sales_rows(data["orders"])
    elif "sales_rows" not in data:
        data["sales_rows"] = []
    # Excel 다운로드 링크용 쿼리스트링
    qs = request.url.query
    return render(
        request,
        "settlement.html",
        {
            "page": "settlement",
            "data": data,
            "month": month,
            "query_string": qs,
        },
    )


@app.get("/settlement/download")
async def settlement_download(
    period: str = "month",
    month: str = "",
    day: str = "",
    date_from: str = "",
    date_to: str = "",
    pay_method: str = "",
    payment_status: str = "",
):
    if not month and period == "month":
        month = date.today().strftime("%Y-%m")
    path = export_monthly_settlement(
        month=month or None,
        period=period,
        day=day or None,
        date_from=date_from or None,
        date_to=date_to or None,
        pay_method=pay_method or None,
        payment_status=payment_status or None,
    )
    return FileResponse(path, filename=path.name)


@app.get("/prices", response_class=HTMLResponse)
async def prices_page(request: Request):
    from src.config_loader import resolve_path, load_config
    catalog = resolve_path(load_config()["paths"]["price_catalog"])
    return render(request, "prices.html", {"page": "prices", "catalog_path": str(catalog)})


@app.get("/prices/download")
async def prices_download():
    from src.config_loader import resolve_path, load_config
    path = resolve_path(load_config()["paths"]["price_catalog"])
    create_price_catalog_template(path)
    return FileResponse(path, filename="price_catalog.xlsx")


@app.get("/prices/validate")
async def prices_validate():
    path = validate_prices()
    return FileResponse(path, filename=path.name)


@app.get("/import", response_class=HTMLResponse)
async def import_page(request: Request, err: str = ""):
    return render(
        request,
        "import.html",
        {"page": "import", "error": err or None},
    )


@app.post("/import/excel")
async def import_excel(request: Request, file: UploadFile = File(...)):
    """참고용 엑셀 파일 일회 업로드 → DB 적재."""
    try:
        upload_dir = ROOT_DIR / "data" / "imports"
        upload_dir.mkdir(parents=True, exist_ok=True)
        dest = upload_dir / (file.filename or "upload.xlsx")
        with dest.open("wb") as f:
            shutil.copyfileobj(file.file, f)
        summary = import_to_db(dest)
        return RedirectResponse(
            f"/orders?msg=imported_{summary['total_orders']}",
            status_code=303,
        )
    except Exception as exc:
        return render(
            request,
            "import.html",
            {"page": "import", "error": str(exc)},
            status_code=500,
        )


@app.post("/import/local")
async def import_local():
    """프로젝트 폴더의 참고 엑셀 자동 import."""
    xlsx = find_latest_xlsx()
    if not xlsx:
        return RedirectResponse("/import?err=no_file", status_code=303)
    summary = import_to_db(xlsx)
    return RedirectResponse(
        f"/orders?msg=imported_{summary['total_orders']}",
        status_code=303,
    )
