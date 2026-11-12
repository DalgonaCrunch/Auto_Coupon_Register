# SoulStrike 쿠폰 등록 — 웹 버전 (Vercel)

PC를 켜 둘 필요 없는 웹페이지 버전. 쿠폰 코드를 넣으면 등록된 모든 CS 코드에 한 번에 등록한다.
루트의 텔레그램 봇 / Selenium 버전은 그대로 두었다.

- 주소: https://soulstrike-coupon.vercel.app (비밀번호: Vercel 환경변수 `PAGE_PASSWORD`)
- 관리자 비밀번호: Vercel 환경변수 `ADMIN_PASSWORD`. 삭제, 이미 붙은 아이디 바꾸기, 서버 변경에 필요하다 (화면 맨 아래 "관리자 모드"). 설정이 없으면 이 기능들은 잠긴다. 일반 비밀번호로는 쿠폰 등록과 새 CS 코드 추가, 아이디가 없는 CS 코드에 아이디 붙이기까지 된다.
- 크롬 없음: 쿠폰 페이지 HTML에서 `Page-Key`, `Coupon-Page-Token` 을 읽어 `POST /tp/coupon/use` 를 직접 호출
- 저장소: Upstash Redis (Vercel Marketplace) — 키 `coupon:ids`(CS 코드 목록), `coupon:names`(CS 코드 → 게임 아이디), `coupon:server`
- ID 목록·결과에 아이디와 CS 코드를 같이 보여 준다. 직접 추가는 한 줄에 `아이디 CS코드`, 엑셀은 2행부터 A·B열 (숫자만 있는 칸 = CS 코드, 다른 칸 = 아이디). 이미 있는 CS 코드에 아이디를 주면 아이디만 바뀐다.

## 구성

| 파일 | 역할 |
|---|---|
| `index.html` | 화면 (비밀번호, 쿠폰 등록, 아이디+CS 코드 추가/삭제/엑셀, 서버 변경) |
| `api/index.py` | API (`/api/state`, `/api/servers`, `/api/ids`, `/api/server`, `/api/coupon`) |
| `vercel.json` | `/api/*` → `api/index.py` |

응답 코드: `101` = 사용 완료(성공), `302` = 없는 쿠폰, `304` = 이미 사용한 쿠폰.

## 배포

GitHub `main` 에 push 하면 Vercel 이 자동 배포한다 (프로젝트 `soulstrike-coupon`, Root Directory = `web`).
`web/` 밖만 바뀐 커밋은 배포를 건너뛴다 (Ignored Build Step: `git diff --quiet HEAD^ HEAD -- .`).

처음 설정할 때만:

```bash
vercel env add PAGE_PASSWORD production   # 페이지 비밀번호
vercel env add ADMIN_PASSWORD production  # 관리자 비밀번호 (삭제·수정·서버 변경)
vercel integration add upstash/upstash-kv # KV_REST_API_URL / KV_REST_API_TOKEN 자동 등록
```
