"""배포 환경 설정 — 로컬·클라우드 공통."""

from __future__ import annotations

import os
from pathlib import Path

from src.config_loader import ROOT_DIR, resolve_path


def load_dotenv() -> None:
    """프로젝트 루트 .env 파일 로드 (있을 때만)."""
    env_path = ROOT_DIR / ".env"
    if not env_path.exists():
        return
    try:
        from dotenv import load_dotenv as _load

        _load(env_path)
    except ImportError:
        # python-dotenv 미설치 시 수동 파싱
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))


def get_database_path(config: dict) -> Path:
    """
    DATABASE_URL 우선, 없으면 config.yaml db_file.
    sqlite:///data/orders.db 형식 지원 (Supabase Postgres는 추후 확장).
    """
    url = os.getenv("DATABASE_URL", "").strip()
    if url.startswith("sqlite:///"):
        rel = url.removeprefix("sqlite:///")
        p = Path(rel)
        return p if p.is_absolute() else ROOT_DIR / rel
    if url.startswith("postgres://") or url.startswith("postgresql://"):
        # 배포 시 Postgres 어댑터 연결 전까지 config 경로 fallback
        pass
    return resolve_path(config["paths"]["db_file"])


def get_upload_dir_override() -> Path | None:
    """MONING_UPLOAD_DIR 환경변수 (클라우드 스토리지 연동 전 로컬 경로)."""
    raw = os.getenv("MONING_UPLOAD_DIR", "").strip()
    if not raw:
        return None
    p = Path(raw)
    return p if p.is_absolute() else ROOT_DIR / raw


def get_api_base_url() -> str:
    """Netlify 등 프론트 분리 배포 시 API 베이스 URL."""
    return os.getenv("MONING_API_URL", "").rstrip("/")


def is_production() -> bool:
    return os.getenv("MONING_ENV", "development").lower() in ("production", "prod")
