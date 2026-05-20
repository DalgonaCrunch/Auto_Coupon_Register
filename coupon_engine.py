"""Headless Selenium engine for the SoulStrike coupon page.

Reuses the same DOM selectors as soulstrike_coupon_auto.py but is callable
from the Telegram bot and returns structured results instead of writing Excel.
"""
from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Callable, Iterable

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

URL = "https://coupon.withhive.com/soulstrike?t=1737278324092"

CS_CODE_INPUT_SELECTOR = "input#cs_code"
COUPON_INPUT_SELECTOR = "input#coupon_code"
SUBMIT_BUTTON_SELECTOR = "button.btn_use"
SERVER_DROPDOWN_BTN = "div.select_wrap > button.btn_select"
SERVER_OPTION_BTN = "div.select_wrap ul.list_select li button"

SLEEP_SEC = 1.5

ProgressFn = Callable[[str], None]


@dataclass
class CouponResult:
    user_id: str
    coupon: str
    message: str
    ok: bool


def _noop(_: str) -> None:
    pass


def _build_driver(headless: bool) -> webdriver.Chrome:
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1280,900")
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )


def _select_server(driver: webdriver.Chrome, wait: WebDriverWait, server_text: str) -> None:
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SERVER_DROPDOWN_BTN))).click()
    options = wait.until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, SERVER_OPTION_BTN))
    )
    for opt in options:
        if opt.text.strip() == server_text:
            opt.click()
            return
    raise RuntimeError(f"서버 옵션 '{server_text}'을(를) 찾지 못했습니다.")


def register_coupons(
    ids: Iterable[str],
    coupons: Iterable[str],
    server: str,
    headless: bool = True,
    on_progress: ProgressFn | None = None,
) -> list[CouponResult]:
    """Run the full registration flow and return per-(id, coupon) results."""
    progress = on_progress or _noop
    ids_list = [str(x) for x in ids if str(x).strip()]
    coupons_list = [c.strip() for c in coupons if c and c.strip()]
    if not ids_list:
        raise ValueError("등록할 ID가 없습니다.")
    if not coupons_list:
        raise ValueError("쿠폰 코드가 없습니다.")

    results: list[CouponResult] = []
    driver = _build_driver(headless)
    try:
        driver.get(URL)
        wait = WebDriverWait(driver, 10)
        progress("페이지 접속 완료")

        _select_server(driver, wait, server)
        progress(f"서버 선택 완료: {server}")
        time.sleep(0.5)

        for cs_code in ids_list:
            for coupon in coupons_list:
                msg, ok = _submit_one(driver, wait, cs_code, coupon)
                results.append(CouponResult(cs_code, coupon, msg, ok))
                progress(f"{cs_code} / {coupon} → {msg}")
                time.sleep(SLEEP_SEC)
    finally:
        driver.quit()

    return results


def _submit_one(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    cs_code: str,
    coupon: str,
) -> tuple[str, bool]:
    try:
        cs_input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, CS_CODE_INPUT_SELECTOR))
        )
        cs_input.clear()
        cs_input.send_keys(cs_code)

        coupon_input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, COUPON_INPUT_SELECTOR))
        )
        coupon_input.clear()
        coupon_input.send_keys(coupon)

        submit_btn = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, SUBMIT_BUTTON_SELECTOR))
        )
        submit_btn.click()
    except Exception as e:
        return f"입력 오류: {e}", False

    try:
        popup = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "div.pop_wrap.coupon_lyr"))
        )
        msg = popup.find_element(By.ID, "layer_msg").text.strip()
        try:
            popup.find_element(By.ID, "layer_close_btn").click()
        except Exception:
            pass
        # The site returns Korean success/failure text. Treat anything mentioning
        # 사용완료/지급 as success; anything else (만료, 중복, 오류 등) as failure.
        ok = any(k in msg for k in ("사용완료", "지급", "성공"))
        return msg, ok
    except Exception:
        return "팝업 없음", False
