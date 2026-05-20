"""Telegram bot entry point: receive commands, run the coupon engine, reply."""
from __future__ import annotations

import asyncio
import json
import logging
import os
from collections import Counter
from functools import wraps
from pathlib import Path

from dotenv import load_dotenv
from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
)

import coupon_engine
import id_store

load_dotenv()

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("coupon_bot")

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
ALLOWED_CHAT_IDS: set[int] = {
    int(x) for x in os.environ.get("ALLOWED_CHAT_IDS", "").split(",") if x.strip()
}
DEFAULT_SERVER = os.environ.get("DEFAULT_SERVER", "KR/JP/GLB").strip()
HEADLESS = os.environ.get("HEADLESS", "1").strip() == "1"

STATE_PATH = Path(__file__).resolve().parent / "bot_state.json"
_run_lock = asyncio.Lock()


def _load_state() -> dict:
    if not STATE_PATH.exists():
        return {}
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def _save_state(state: dict) -> None:
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def current_server() -> str:
    return _load_state().get("server", DEFAULT_SERVER)


def authorized(handler):
    @wraps(handler)
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE):
        chat_id = update.effective_chat.id if update.effective_chat else None
        if chat_id not in ALLOWED_CHAT_IDS:
            logger.warning("Unauthorized chat_id=%s blocked", chat_id)
            if update.message:
                await update.message.reply_text(
                    f"권한이 없습니다. chat_id={chat_id} 를 관리자에게 전달하세요."
                )
            return
        return await handler(update, context)

    return wrapper


@authorized
async def cmd_start(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "SoulStrike 쿠폰 봇입니다.\n/help 로 사용법을 확인하세요."
    )


@authorized
async def cmd_help(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    text = (
        "*사용법*\n"
        "`/coupon <코드> [코드 ...]` — 등록된 모든 ID에 쿠폰 적용\n"
        "`/add <ID> [ID ...]` — ID 추가\n"
        "`/del <ID> [ID ...]` — ID 삭제\n"
        "`/list` — 등록된 ID 목록\n"
        "`/server <KR/JP/GLB>` — 기본 서버 변경 (현재: "
        f"{current_server()})\n"
        "`/help` — 도움말"
    )
    await update.message.reply_text(text, parse_mode=ParseMode.MARKDOWN)


@authorized
async def cmd_list(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    ids = id_store.list_ids()
    if not ids:
        await update.message.reply_text("등록된 ID가 없습니다.")
        return
    body = "\n".join(f"{i + 1}. {v}" for i, v in enumerate(ids))
    await update.message.reply_text(f"등록된 ID ({len(ids)}개):\n{body}")


@authorized
async def cmd_add(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.args:
        await update.message.reply_text("사용법: /add <ID> [ID ...]")
        return
    added, dupes = id_store.add_ids(list(context.args))
    lines = []
    if added:
        lines.append(f"추가됨 ({len(added)}): {', '.join(added)}")
    if dupes:
        lines.append(f"이미 있음 ({len(dupes)}): {', '.join(dupes)}")
    if not lines:
        lines.append("처리된 항목이 없습니다.")
    await update.message.reply_text("\n".join(lines))


@authorized
async def cmd_del(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.args:
        await update.message.reply_text("사용법: /del <ID> [ID ...]")
        return
    removed, missing = id_store.remove_ids(list(context.args))
    lines = []
    if removed:
        lines.append(f"삭제됨 ({len(removed)}): {', '.join(removed)}")
    if missing:
        lines.append(f"없음 ({len(missing)}): {', '.join(missing)}")
    if not lines:
        lines.append("처리된 항목이 없습니다.")
    await update.message.reply_text("\n".join(lines))


@authorized
async def cmd_server(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.args:
        await update.message.reply_text(
            f"현재 서버: {current_server()}\n변경: /server <KR/JP/GLB | CN | ...>"
        )
        return
    new_server = " ".join(context.args).strip()
    state = _load_state()
    state["server"] = new_server
    _save_state(state)
    await update.message.reply_text(f"서버를 '{new_server}'(으)로 변경했습니다.")


@authorized
async def cmd_coupon(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.args:
        await update.message.reply_text("사용법: /coupon <코드> [코드 ...]")
        return
    coupons = [c.strip() for c in context.args if c.strip()]
    ids = id_store.list_ids()
    if not ids:
        await update.message.reply_text("등록된 ID가 없습니다. /add 로 먼저 추가하세요.")
        return

    if _run_lock.locked():
        await update.message.reply_text("이미 다른 등록 작업이 진행 중입니다. 잠시 후 다시 시도하세요.")
        return

    server = current_server()
    progress_msg = await update.message.reply_text(
        f"시작합니다.\nID {len(ids)}개 × 쿠폰 {len(coupons)}개 = {len(ids) * len(coupons)}건\n서버: {server}"
    )

    async with _run_lock:
        loop = asyncio.get_running_loop()
        queue: asyncio.Queue[str] = asyncio.Queue()

        def on_progress(line: str) -> None:
            # Called from the Selenium worker thread; hand off to the event loop.
            asyncio.run_coroutine_threadsafe(queue.put(line), loop)

        async def drain_progress() -> None:
            while True:
                line = await queue.get()
                if line == "__DONE__":
                    return
                try:
                    await update.message.reply_text(line)
                except Exception:
                    logger.exception("progress reply failed")

        drainer = asyncio.create_task(drain_progress())

        try:
            results = await asyncio.to_thread(
                coupon_engine.register_coupons,
                ids,
                coupons,
                server,
                HEADLESS,
                on_progress,
            )
        except Exception as e:
            await queue.put("__DONE__")
            await drainer
            logger.exception("coupon run failed")
            await update.message.reply_text(f"실패: {e}")
            return

        await queue.put("__DONE__")
        await drainer

    summary = _format_summary(results)
    await update.message.reply_text(summary, parse_mode=ParseMode.MARKDOWN)


def _format_summary(results: list[coupon_engine.CouponResult]) -> str:
    total = len(results)
    ok = sum(1 for r in results if r.ok)
    fail = total - ok

    per_id_status: dict[str, Counter] = {}
    for r in results:
        per_id_status.setdefault(r.user_id, Counter())[r.message] += 1

    lines = [f"*완료* — 성공 {ok} / 실패 {fail} / 총 {total}"]
    for uid, counter in per_id_status.items():
        parts = ", ".join(f"{msg}×{n}" for msg, n in counter.items())
        lines.append(f"• `{uid}` — {parts}")
    return "\n".join(lines)


def main() -> None:
    if not TOKEN:
        raise SystemExit("TELEGRAM_BOT_TOKEN 이 비어있습니다. .env 를 확인하세요.")
    if not ALLOWED_CHAT_IDS:
        raise SystemExit("ALLOWED_CHAT_IDS 가 비어있습니다. .env 를 확인하세요.")

    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("list", cmd_list))
    app.add_handler(CommandHandler("add", cmd_add))
    app.add_handler(CommandHandler("del", cmd_del))
    app.add_handler(CommandHandler("server", cmd_server))
    app.add_handler(CommandHandler("coupon", cmd_coupon))

    logger.info("Bot starting. Allowed chats: %s", ALLOWED_CHAT_IDS)
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
