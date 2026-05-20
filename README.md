# SoulStrike 쿠폰 텔레그램 봇

기존 Tkinter + Excel 기반 자동등록기를 **텔레그램 봇**으로 옮긴 버전. 봇이 켜져 있는 한 어디서든 텔레그램으로 명령어를 보내면 등록된 ID들에 쿠폰을 자동 등록합니다.

## 설치

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 설정

1. [@BotFather](https://t.me/BotFather)에서 봇을 만들고 **토큰**을 받습니다.
2. 자신의 **chat_id**를 알아냅니다. ([@userinfobot](https://t.me/userinfobot)에 메시지 보내면 알려줌)
3. `.env.example` 을 복사해 `.env` 로 저장하고 채웁니다.

```env
TELEGRAM_BOT_TOKEN=123456:ABC...
ALLOWED_CHAT_IDS=11111111,22222222
DEFAULT_SERVER=KR/JP/GLB
HEADLESS=1
```

- `ALLOWED_CHAT_IDS` 에 들어있지 않은 사용자가 명령어를 보내도 봇은 무시합니다.
- 콤마로 구분하면 여러 명 허용 가능.

## 실행

```powershell
python coupon_bot.py
```

봇이 시작되면 텔레그램에서 명령어를 보낼 수 있습니다.

## 명령어

| 명령어 | 설명 |
|--------|------|
| `/coupon <코드> [코드 ...]` | 등록된 **모든 ID**에 쿠폰을 적용. 진행 상황을 메시지로 회신 |
| `/add <ID> [ID ...]` | ID 추가 |
| `/del <ID> [ID ...]` | ID 삭제 |
| `/list` | 등록된 ID 목록 |
| `/server <KR/JP/GLB>` | 기본 서버 변경 (현재 값 조회는 인자 없이) |
| `/help` | 도움말 |

ID 목록은 `ids.json`, 서버 설정은 `bot_state.json` 에 저장됩니다 (둘 다 gitignore).

## 파일 구성

- `coupon_bot.py` — 텔레그램 봇 진입점
- `coupon_engine.py` — Selenium 자동화 로직 (헤드리스 Chrome)
- `id_store.py` — ID JSON 저장소
- `soulstrike_coupon_auto.py` — 기존 Tkinter/Excel 버전 (보존)

## exe로 빌드 (다른 PC 배포용)

Python 없는 PC에서도 돌리고 싶을 때:

```powershell
# 봇 개발 환경에서 (PyInstaller 추가 설치)
pip install pyinstaller
pyinstaller coupon_bot.spec
```

빌드가 끝나면 `dist/coupon_bot.exe` 가 생깁니다.

**배포 시 같은 폴더에 둘 파일:**

```
coupon_bot.exe
.env              ← 토큰/chat_id 채워서 함께 배포
```

`ids.json` 과 `bot_state.json` 은 봇이 실행되면서 같은 폴더에 자동 생성됩니다.

**다른 PC에서 실행 전 확인:**

- **Chrome 브라우저 설치 필요** (Selenium이 사용). ChromeDriver는 webdriver-manager가 첫 실행 시 자동으로 받습니다 (인터넷 필요).
- exe 실행 시 콘솔 창이 뜨고, **그 창을 닫으면 봇도 꺼집니다**. 닫지 마세요.
- 봇이 24시간 응답하길 원하면 PC도 24시간 켜둬야 합니다. 작업 스케줄러로 부팅 시 자동 실행을 등록할 수도 있습니다.

## 클라우드 배포 (선택)

PC가 꺼져있어도 봇이 돌아가게 하려면 클라우드로 옮기면 됩니다. 추천:

- **Oracle Cloud Free Tier (ARM Ampere, 4vCPU/24GB, 영구 무료)** — 가장 여유로움
- **GCP e2-micro (1vCPU/1GB, 영구 무료)** — 작은 봇에 충분

배포 시 추가로 필요한 것:

```bash
# Ubuntu 기준
sudo apt update
sudo apt install -y chromium-browser  # 또는 google-chrome
```

`HEADLESS=1` 이면 디스플레이 없이 동작합니다. 서비스로 등록하려면 `systemd` unit 파일을 만들어 `python coupon_bot.py` 를 실행.
