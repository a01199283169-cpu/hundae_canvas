# 모닝프레임 배포 가이드 (GitHub → DB → Netlify)

## 권장 아키텍처

```
GitHub (소스) → Supabase (PostgreSQL + Storage) → Render (FastAPI API)
                                              ↘ Netlify (선택: 정적/프록시)
```

| 구성요소 | 역할 | 비고 |
|---------|------|------|
| **GitHub** | 코드 저장·CI | push 시 자동 배포 연동 |
| **Supabase** | PostgreSQL DB, 이미지 Storage | SQLite 대체, 영구 저장 |
| **Render / Railway** | FastAPI(uvicorn) 호스팅 | Netlify 단독으로는 API+DB 불가 |
| **Netlify** | (선택) 도메인·CDN·리다이렉트 | API URL로 프록시 |

> **중요:** Netlify는 서버리스·정적 호스팅 중심입니다. 현재 앱(FastAPI + Jinja2 + SQLite 파일)은 **Netlify 단독 배포가 불가**하며, API 서버와 클라우드 DB가 필요합니다.

## 로컬 개발

```bat
copy .env.example .env
pip install -r requirements.txt
run_web.bat
```

## Render 배포 (API)

1. GitHub 저장소 연결
2. `render.yaml` 사용 또는 수동:
   - Build: `pip install -r requirements.txt`
   - Start: `uvicorn web.app:app --host 0.0.0.0 --port $PORT`
3. 환경변수:
   - `DATABASE_URL=sqlite:////data/orders.db` (Render Disk 마운트 `/data`)
   - 또는 Supabase Postgres URL

## Supabase 전환 (추후)

1. `supabase/migrations/001_initial.sql` 실행
2. `DATABASE_URL=postgresql://...` 설정
3. `src/database.py`에 Postgres 드라이버(psycopg2/asyncpg) 어댑터 추가 필요

## Netlify 연동

- API를 Render에 배포한 뒤 `netlify.toml`의 redirect로 `/api/*` 프록시
- 또는 Netlify에 커스텀 도메인만 연결하고 API는 `api.도메인.com` 서브도메인

## 환경변수

| 변수 | 설명 |
|------|------|
| `DATABASE_URL` | DB 연결 (sqlite:/// 또는 postgresql://) |
| `MONING_UPLOAD_DIR` | 이미지 업로드 경로 |
| `MONING_ENV` | production 시 디버그 비활성화 |
| `MONING_API_URL` | 프론트 분리 시 API 베이스 URL |
| `PORT` | uvicorn 포트 |

## 배포 전 체크리스트

- [ ] `.env`는 Git에 커밋하지 않음 (`.env.example`만)
- [ ] `data/orders.db`는 Git 제외 (`.gitignore`)
- [ ] Supabase RLS·API 키 보안 설정
- [ ] 업로드 이미지 → Supabase Storage 이전
- [ ] 엑셀 import는 관리자 전용으로 제한
