"""설정 파일 로드 및 경로 유틸리티."""

from __future__ import annotations

import yaml
from pathlib import Path

# 프로젝트 루트 (src 상위)
ROOT_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT_DIR / "config.yaml"


def load_config() -> dict:
    """config.yaml을 읽어 dict로 반환."""
    with CONFIG_PATH.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


def resolve_path(relative: str) -> Path:
    """설정 기준 상대 경로를 절대 경로로 변환."""
    return ROOT_DIR / relative


def ensure_dirs(config: dict | None = None) -> dict[str, Path]:
    """필요한 data/output 하위 폴더 생성."""
    if config is None:
        config = load_config()

    paths = {
        "data": resolve_path(config["paths"]["data_dir"]),
        "output": resolve_path(config["paths"]["output_dir"]),
        "imports": resolve_path(config["paths"]["imports_dir"]),
        "production": resolve_path(config["paths"]["output_dir"]) / "production",
        "settlement": resolve_path(config["paths"]["output_dir"]) / "settlement",
        "images": resolve_path(config["paths"]["output_dir"]) / "images",
        "logs": resolve_path(config["paths"]["output_dir"]) / "logs",
        "validation": resolve_path(config["paths"]["output_dir"]) / "validation",
    }
    for p in paths.values():
        p.mkdir(parents=True, exist_ok=True)
    return paths
