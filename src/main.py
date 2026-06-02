"""
모닝프레임 주문·생산·월결산 자동화 CLI 진입점.

사용법:
  python -m src.main import [엑셀파일]
  python -m src.main all [엑셀파일]
  python -m src.main --help
"""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from src.config_loader import ROOT_DIR, ensure_dirs, load_config, resolve_path
from src.database import init_db
from src.image_extractor import build_row_order_map, extract_and_save_images
from src.parser import import_to_db
from src.price_validator import create_price_catalog_template, validate_prices
from src.production_export import export_production_sheet
from src.seed_price_catalog import seed_price_catalog_from_orders
from src.settlement_export import export_monthly_settlement


def setup_logging() -> Path:
    """로그 파일 설정."""
    paths = ensure_dirs()
    log_file = paths["logs"] / f"{datetime.now().strftime('%Y-%m-%d')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return log_file


def find_latest_xlsx(search_dir: Path | None = None) -> Path | None:
    """프로젝트 루트 또는 imports 폴더에서 최신 xlsx 탐색."""
    config = load_config()
    dirs = [ROOT_DIR, resolve_path(config["paths"]["imports_dir"])]
    if search_dir:
        dirs.insert(0, search_dir)

    candidates: list[Path] = []
    for d in dirs:
        if not d.exists():
            continue
        for f in list(d.glob("*.xlsx")) + list(d.glob("*.xlsm")):
            if not f.name.startswith("~$") and "price_catalog" not in f.name:
                candidates.append(f)

    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def archive_import(xlsx_path: Path) -> Path:
    """원본 엑셀을 data/imports에 복사 보관."""
    paths = ensure_dirs()
    dest = paths["imports"] / xlsx_path.name
    if xlsx_path.resolve() != dest.resolve():
        shutil.copy2(xlsx_path, dest)
    return dest


def cmd_import(xlsx_path: Path) -> dict:
    """엑셀 → DB 적재."""
    logging.info("주문 엑셀 import 시작: %s", xlsx_path)
    archived = archive_import(xlsx_path)
    summary = import_to_db(archived)
    logging.info(
        "Import 완료: 주문 %d건, 품목 %d개",
        summary["total_orders"],
        summary["total_items"],
    )
    return summary


def cmd_images(xlsx_path: Path) -> dict:
    """이미지 추출 및 주문 매핑."""
    logging.info("이미지 추출 시작: %s", xlsx_path)
    row_map = build_row_order_map(xlsx_path)
    stats = extract_and_save_images(xlsx_path, row_map)
    logging.info(
        "이미지 추출: %d개 (매핑 %d, 미매핑 %d)",
        stats["extracted"],
        stats["mapped"],
        stats["unmapped"],
    )
    return stats


def cmd_validate(source_file: str | None) -> Path:
    """단가 검증 리포트."""
    logging.info("단가 검증 시작")
    out = validate_prices(source_file)
    logging.info("단가 검증 리포트: %s", out)
    return out


def cmd_production(source_file: str | None) -> Path:
    """생산지시서 생성."""
    logging.info("생산지시서 생성")
    out = export_production_sheet(source_file)
    logging.info("생산지시서: %s", out)
    return out


def cmd_settlement(source_file: str | None, year_month: str | None) -> Path:
    """월 결산표 생성."""
    logging.info("월 결산 생성 (month=%s)", year_month)
    out = export_monthly_settlement(month=year_month, source_file=source_file)
    logging.info("월 결산표: %s", out)
    return out


def open_file(path: Path) -> None:
    """Windows에서 생성된 엑셀 파일을 자동으로 연다."""
    if path and path.exists():
        os.startfile(str(path))


def open_output_folder() -> None:
    """output 폴더를 탐색기로 연다."""
    paths = ensure_dirs()
    subprocess.Popen(["explorer", str(paths["output"])])


def save_result_summary(results: dict) -> Path:
    """처리 결과 경로를 텍스트 파일로 저장 (VBA/사용자 확인용)."""
    paths = ensure_dirs()
    summary_path = paths["output"] / "최신_처리결과.txt"
    lines = [
        f"처리일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"주문 건수: {results.get('order_count', 0)}건",
        "",
        "[생성된 파일]",
        f"생산지시서: {results.get('production', '')}",
        f"월결산표:   {results.get('settlement', '')}",
        f"단가검증:   {results.get('validation', '')}",
        "",
        "※ 위 파일이 자동으로 열리지 않으면 경로를 복사해 직접 여세요.",
    ]
    summary_path.write_text("\n".join(lines), encoding="utf-8")
    return summary_path


def cmd_all(xlsx_path: Path, year_month: str | None = None) -> dict:
    """전체 파이프라인 실행."""
    init_db()
    create_price_catalog_template()

    summary = cmd_import(xlsx_path)
    source_file = summary["source_file"]

    # 주문 데이터 기반 단가표 보강 (플랫폼 공개단가 확인용 초기값)
    seed_price_catalog_from_orders()

    cmd_images(xlsx_path)
    validation_path = cmd_validate(source_file)
    production_path = cmd_production(source_file)
    settlement_path = cmd_settlement(source_file, year_month)

    results = {
        "order_count": summary["total_orders"],
        "production": production_path,
        "settlement": settlement_path,
        "validation": validation_path,
    }
    save_result_summary(results)

    logging.info("=== 전체 처리 완료 ===")
    logging.info("주문 %d건 처리됨", summary["total_orders"])
    logging.info("생산지시서: %s", production_path)
    logging.info("월결산표: %s", settlement_path)

    # ★ 결과 엑셀 자동 열기 (사용자가 '반응 없음'으로 느끼는 문제 해결)
    print("\n" + "=" * 50)
    print("  처리 완료! 결과 파일을 엽니다...")
    print("=" * 50)
    print(f"  [생산지시서] {production_path.name}")
    print(f"  [월결산표]   {settlement_path.name}")
    print("=" * 50 + "\n")

    open_file(production_path)
    open_file(settlement_path)
    open_output_folder()

    return results


def main(argv: list[str] | None = None) -> int:
    setup_logging()
    init_db()

    parser = argparse.ArgumentParser(
        description="모닝프레임 주문·생산·월결산 자동화",
    )
    sub = parser.add_subparsers(dest="command")

    p_all = sub.add_parser("all", help="전체 처리 (import+이미지+단가+생산+결산)")
    p_all.add_argument("file", nargs="?", help="엑셀 파일 경로")
    p_all.add_argument("--month", help="결산 월 (YYYY-MM)")

    p_import = sub.add_parser("import", help="엑셀 DB 적재")
    p_import.add_argument("file", nargs="?", help="엑셀 파일 경로")

    p_img = sub.add_parser("images", help="이미지 추출")
    p_img.add_argument("file", nargs="?", help="엑셀 파일 경로")

    p_val = sub.add_parser("validate", help="단가 검증")
    p_val.add_argument("--file", help="source_file명")

    p_prod = sub.add_parser("production", help="생산지시서")
    p_prod.add_argument("--file", help="source_file명")

    p_set = sub.add_parser("settlement", help="월 결산")
    p_set.add_argument("--file", help="source_file명")
    p_set.add_argument("--month", help="YYYY-MM")

    p_init = sub.add_parser("init", help="DB·단가표 초기화")

    args = parser.parse_args(argv)

    if args.command == "init":
        create_price_catalog_template()
        logging.info("초기화 완료 (DB + 단가표 템플릿)")
        return 0

    if args.command is None:
        # 인자 없으면 all + 자동 파일 탐색
        xlsx = find_latest_xlsx()
        if not xlsx:
            parser.print_help()
            logging.error("엑셀 파일을 찾을 수 없습니다.")
            return 1
        cmd_all(xlsx)
        return 0

    # 파일 경로 결정
    file_arg = getattr(args, "file", None)
    xlsx_path = Path(file_arg) if file_arg else find_latest_xlsx()
    if args.command in ("all", "import", "images") and (not xlsx_path or not xlsx_path.exists()):
        logging.error("엑셀 파일을 찾을 수 없습니다: %s", file_arg)
        return 1

    if args.command == "all":
        cmd_all(xlsx_path, getattr(args, "month", None))
    elif args.command == "import":
        cmd_import(xlsx_path)
    elif args.command == "images":
        cmd_images(xlsx_path)
    elif args.command == "validate":
        cmd_validate(getattr(args, "file", None))
    elif args.command == "production":
        cmd_production(getattr(args, "file", None))
    elif args.command == "settlement":
        cmd_settlement(getattr(args, "file", None), getattr(args, "month", None))

    return 0


if __name__ == "__main__":
    sys.exit(main())
