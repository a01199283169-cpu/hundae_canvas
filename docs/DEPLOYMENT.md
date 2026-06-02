# Render + Supabase 배포 가이드

## 아키텍처

```
GitHub (main)
    ↓
Render Web Service     ← FastAPI (화면 + API)
    + Disk /data       ← 업로드 이미지
    ↓
Supabase PostgreSQL    ← 주문·매출 DB (영구)
```

| 구성 | 역할 |
|------|------|
| **Supabase** | PostgreSQL DB — 주문, 품목, 매출 데이터 |
| **Render** | Python 서버 (uvicorn) — 웹 화면 전체 |
| **Render Disk** | 주문 이미지 파일 (`/data/uploads`) |

> Netlify는 **필요 없습니다.** Render URL 하나로 접속합니다.

---

## 1단계: Supabase 프로젝트

1. [supabase.com](https://supabase.com) → **New project**
2. Region: **Northeast Asia (Seoul)** 권장
3. DB 비밀번호 저장 (분실 시 재설정 필요)

### 테이블 생성

**SQL Editor** → New query → 아래 파일 내용 붙여넣기 후 **Run**:

`supabase/migrations/001_initial.sql`

또는 로컬에서 `DATABASE_URL`을 Supabase URI로 설정 후 앱 기동 시 `init_db()`가 자동 생성합니다.

### Connection string 복사

**Project Settings → Database → Connection string → URI**

- **Transaction pooler** (포트 `6543`) — Render 권장
- 끝에 `?sslmode=require` 포함 확인

예:

```text
postgresql://postgres.xxxxx:비밀번호@aws-0-ap-northeast-2.pooler.supabase.com:6543/postgres?sslmode=require
```

---

## 2단계: Render Web Service

1. [render.com](https://render.com) → **New → Blueprint** 또는 **Web Service**
2. GitHub `a01199283169-cpu/hundae_canvas` 연결
3. Branch: **main**
4. **Blueprint** 사용 시: 저장소의 `render.yaml` 자동 적용

### 수동 설정 시

| 항목 | 값 |
|------|-----|
| Runtime | Python 3 |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `uvicorn web.app:app --host 0.0.0.0 --port $PORT` |

### Persistent Disk

- **Add Disk** → Name: `hundae-data`, Mount: `/data`, Size: 1GB

### Environment Variables

| Key | Value |
|-----|--------|
| `MONING_ENV` | `production` |
| `DATABASE_URL` | Supabase URI (위에서 복사) |
| `MONING_UPLOAD_DIR` | `/data/uploads` |

5. **Create Web Service** → Deploy

배포 URL 예: `https://hundae-canvas.onrender.com`

---

## 3단계: 데이터 넣기

Supabase DB는 처음 **비어 있습니다.**

1. Render URL 접속
2. **엑셀 참조 Import** → 주문내역서 xlsx 업로드  
   또는 로컬에서 import 후 Supabase만 쓰는 경우는 Render에서 Import

---

## 4단계: 동작 확인

- [ ] 대시보드 차트 표시
- [ ] 주문 관리 목록
- [ ] 매출현황 일자별·월합계
- [ ] 주문 등록 + 이미지 업로드
- [ ] Supabase **Table Editor**에서 `orders` 데이터 확인

---

## 로컬에서 Supabase 연결 테스트

```bat
copy .env.example .env
```

`.env`에 Supabase `DATABASE_URL` 설정 후:

```bat
pip install -r requirements.txt
run_web.bat
```

---

## 환경변수 요약

| 변수 | 로컬 | Render |
|------|------|--------|
| `DATABASE_URL` | `sqlite:///data/orders.db` | Supabase Postgres URI |
| `MONING_UPLOAD_DIR` | `output/images/uploads` | `/data/uploads` |
| `MONING_ENV` | `development` | `production` |

---

## 트러블슈팅

| 증상 | 해결 |
|------|------|
| `could not connect to server` | URI에 `?sslmode=require` 추가, 비밀번호 특수문자 URL 인코딩 |
| 테이블 없음 | Supabase SQL Editor에서 `001_initial.sql` 실행 |
| 이미지 사라짐 | `MONING_UPLOAD_DIR=/data/uploads` + Render Disk 마운트 확인 |
| 슬립(첫 접속 느림) | Render 무료 플랜 — Starter($7/월) 업그레이드 |

---

## (선택) Supabase Storage

이미지를 Disk 대신 Supabase Storage에 두려면 추후 `src/storage.py` 연동이 필요합니다.  
현재는 **Render Disk + Postgres** 조합으로 충분합니다.
