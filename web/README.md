# SoulStrike 쿠폰 등록 — 웹 버전 (Vercel)

PC를 켜 둘 필요 없는 웹페이지 버전. 쿠폰 코드를 넣으면 등록된 모든 CS 코드에 한 번에 등록한다.
루트의 텔레그램 봇 / Selenium 버전은 그대로 두었다.

- 주소: https://soulstrike-coupon.vercel.app (비밀번호: Vercel 환경변수 `PAGE_PASSWORD`)
- 크롬 없음: 쿠폰 페이지 HTML에서 `Page-Key`, `Coupon-Page-Token` 을 읽어 `POST /tp/coupon/use` 를 직접 호출
- 저장소: Upstash Redis (Vercel Marketplace) — 키 `coupon:ids`, `coupon:server`

## 구성

| 파일 | 역할 |
|---|---|
| `index.html` | 화면 (비밀번호, 쿠폰 등록, ID 추가/삭제/엑셀, 서버 변경) |
| `api/index.py` | API (`/api/state`, `/api/servers`, `/api/ids`, `/api/server`, `/api/coupon`) |
| `vercel.json` | `/api/*` → `api/index.py` |

응답 코드: `101` = 사용 완료(성공), `302` = 없는 쿠폰, `304` = 이미 사용한 쿠폰.

## 배포

```bash
cd web
vercel link --project soulstrike-coupon
vercel env add PAGE_PASSWORD production
vercel deploy --prod
```
