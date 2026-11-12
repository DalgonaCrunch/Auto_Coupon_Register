"""Vercel serverless API for the SoulStrike coupon page.

Routes (all require the X-Password header to match PAGE_PASSWORD).
Changing or deleting what already exists also needs X-Admin-Password to match
ADMIN_PASSWORD; with ADMIN_PASSWORD unset those actions stay locked.
  GET  /api/admin            -> 200 if the admin password is right, else 403
  GET  /api/state            -> {"ids": [...], "names": {cs: name}, "server": "..."}
  POST /api/ids    {"add": [...]} | {"entries": [{"cs": .., "name": ..}]} | {"remove": [...]}
  POST /api/server {"server": "KR/JP/GLB"}
  POST /api/coupon {"coupon": "CODE"}  -> registers CODE for every saved ID

Coupon registration talks to the coupon site's own endpoint
(POST /tp/coupon/use) instead of driving Chrome. The page-scoped keys that
endpoint needs are read from the coupon page HTML on every run.
"""
from __future__ import annotations

import hmac
import http.cookiejar
import json
import os
import re
import time
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler

COUPON_BASE = "https://coupon.withhive.com"
COUPON_PAGE = COUPON_BASE + "/soulstrike"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/130.0 Safari/537.36"
)
SUCCESS_CODES = {"100", "101"}  # 101 = "쿠폰 사용이 완료 되었습니다"
DEFAULT_SERVER = os.environ.get("DEFAULT_SERVER", "KR/JP/GLB").strip()
SLEEP_SEC = 0.3


# ---------------------------------------------------------------- storage ---
def _redis_conf() -> tuple[str, str]:
    url = os.environ.get("KV_REST_API_URL") or os.environ.get("UPSTASH_REDIS_REST_URL") or ""
    token = os.environ.get("KV_REST_API_TOKEN") or os.environ.get("UPSTASH_REDIS_REST_TOKEN") or ""
    if not url or not token:
        raise RuntimeError("Redis 환경변수가 없습니다 (KV_REST_API_URL / KV_REST_API_TOKEN).")
    return url.rstrip("/"), token


def _redis(*cmd: str):
    url, token = _redis_conf()
    req = urllib.request.Request(
        url,
        data=json.dumps(list(cmd)).encode(),
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=10) as r:
        body = json.loads(r.read().decode())
    if "error" in body:
        raise RuntimeError(f"Redis 오류: {body['error']}")
    return body.get("result")


def load_ids() -> list[str]:
    raw = _redis("GET", "coupon:ids")
    if not raw:
        return []
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return []
    return [str(x) for x in data] if isinstance(data, list) else []


def save_ids(ids: list[str]) -> None:
    _redis("SET", "coupon:ids", json.dumps(ids, ensure_ascii=False))


def load_names() -> dict[str, str]:
    """CS code -> 게임 아이디(표시용 이름). Stored apart from coupon:ids so the
    old list-of-strings format keeps working untouched."""
    raw = _redis("GET", "coupon:names")
    if not raw:
        return {}
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return {}
    return {str(k): str(v) for k, v in data.items()} if isinstance(data, dict) else {}


def save_names(names: dict[str, str]) -> None:
    _redis("SET", "coupon:names", json.dumps(names, ensure_ascii=False))


def clean_entries(entries) -> list[tuple[str, str]]:
    """[{"cs": "...", "name": "..."}] -> [(cs, name)], deduped by CS code.
    CS code keeps digits only; a later row for the same code wins."""
    out: dict[str, str] = {}
    for e in entries or []:
        if not isinstance(e, dict):
            continue
        cs = re.sub(r"[^0-9]", "", str(e.get("cs") or ""))
        if not cs:
            continue
        name = re.sub(r"\s+", " ", str(e.get("name") or "")).strip()[:40]
        if name or cs not in out:
            out[cs] = name
    return list(out.items())


def merge_entries(ids: list[str], names: dict[str, str], entries, admin: bool = False) -> dict:
    """Add new CS codes. Setting a name on a code that is already saved counts
    as an edit: an admin may do it, anyone else gets it back in "locked".
    (A saved code with no name yet may get its first name without admin.)"""
    ids, names = list(ids), dict(names)
    added, updated, dupes, locked = [], [], [], []
    for cs, name in clean_entries(entries):
        if cs not in ids:
            ids.append(cs)
            added.append(cs)
            if name:
                names[cs] = name
        elif name and names.get(cs) != name:
            if admin or not names.get(cs):
                names[cs] = name
                updated.append(cs)
            else:
                locked.append(cs)
        else:
            dupes.append(cs)
    return {"ids": ids, "names": names, "added": added, "updated": updated,
            "dupes": dupes, "locked": locked}


def load_server() -> str:
    return _redis("GET", "coupon:server") or DEFAULT_SERVER


def save_server(server: str) -> None:
    _redis("SET", "coupon:server", server)


def clean_items(items) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for raw in items or []:
        for t in re.split(r"[,\s]+", str(raw)):
            t = t.strip()
            if t and t not in seen:
                out.append(t)
                seen.add(t)
    return out


# ----------------------------------------------------------------- engine ---
class CouponSite:
    """One browsing session against the coupon page: cookies + page keys."""

    def __init__(self) -> None:
        self.opener = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar())
        )
        html = self._get(COUPON_PAGE)
        self.page_key = self._grab(html, r"'Page-Key':\s*'([^']*)'")
        self.ct_key = self._grab(html, r"'CT-Key':\s*'([^']*)'", required=False)
        self.page_token = self._grab(html, r"'Coupon-Page-Token':\s*'([^']*)'")
        m = re.search(r"var gParam = (\{.*?\});\s*\n", html, re.S)
        if not m:
            raise RuntimeError("쿠폰 페이지 구조가 바뀌었습니다 (gParam 없음).")
        gparam = json.loads(m.group(1))
        self.game_index = gparam["gameInfo"]["gameindex"]
        self.worlds = gparam.get("gameWorldList") or []

    def _get(self, url: str) -> str:
        req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "ko"})
        with self.opener.open(req, timeout=15) as r:
            return r.read().decode("utf-8", "replace")

    @staticmethod
    def _grab(html: str, pattern: str, required: bool = True) -> str:
        m = re.search(pattern, html)
        if not m:
            if required:
                raise RuntimeError("쿠폰 페이지 구조가 바뀌었습니다 (키 없음).")
            return ""
        return m.group(1)

    def server_names(self) -> list[str]:
        return [w.get("display_world_title_ko") or w.get("name") for w in self.worlds]

    def server_value(self, server: str) -> str:
        want = server.strip().lower()
        for w in self.worlds:
            names = {
                str(w.get("name") or "").lower(),
                str(w.get("display_world_title_ko") or "").lower(),
                str(w.get("display_world_title_en") or "").lower(),
            }
            if want in names:
                title = w.get("display_world_title_ko") or w.get("name")
                return f"{self.game_index}|{w.get('name')}|{title}"
        raise RuntimeError(f"서버 '{server}'을(를) 찾지 못했습니다. 가능: {', '.join(self.server_names())}")

    def use(self, server_value: str, cs_code: str, coupon: str) -> tuple[bool, str, str]:
        body = {
            "language": "ko",
            "server": server_value,
            "cs_code": re.sub(r"[^0-9]", "", cs_code),
            "coupon": coupon.strip(),
        }
        req = urllib.request.Request(
            COUPON_BASE + "/tp/coupon/use",
            data=json.dumps(body).encode(),
            headers={
                "User-Agent": UA,
                "Content-Type": "application/json",
                "Origin": COUPON_BASE,
                "Referer": COUPON_PAGE,
                "Page-Key": self.page_key,
                "CT-Key": self.ct_key,
                "Coupon-Page-Token": self.page_token,
            },
            method="POST",
        )
        try:
            with self.opener.open(req, timeout=15) as r:
                res = json.loads(r.read().decode("utf-8", "replace"))
        except Exception as e:  # network / HTTP error
            return False, "", f"요청 실패: {e}"
        code = str(res.get("code", ""))
        msg = re.sub(r"<br\s*/?>", " ", str(res.get("msg") or res.get("message") or ""))
        ok = code in SUCCESS_CODES or "사용이 완료" in msg
        return ok, code, msg


def register_all(coupon: str, ids: list[str], server: str, names: dict[str, str] | None = None) -> list[dict]:
    site = CouponSite()
    server_value = site.server_value(server)
    results = []
    for i, cs in enumerate(ids):
        if i:
            time.sleep(SLEEP_SEC)
        ok, code, msg = site.use(server_value, cs, coupon)
        results.append({"id": cs, "name": (names or {}).get(cs, ""), "coupon": coupon, "ok": ok, "code": code, "msg": msg})
    return results


# ---------------------------------------------------------------- handler ---
class handler(BaseHTTPRequestHandler):
    def _send(self, status: int, payload: dict) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _check(self, env: str, header: str) -> bool:
        expected = os.environ.get(env, "")
        given = urllib.parse.unquote(self.headers.get(header, ""))
        return bool(expected) and hmac.compare_digest(given.encode(), expected.encode())

    def _authorized(self) -> bool:
        return self._check("PAGE_PASSWORD", "X-Password")

    def _admin(self) -> bool:
        return self._check("ADMIN_PASSWORD", "X-Admin-Password")

    def _deny_admin(self) -> None:
        msg = ("관리자 비밀번호가 설정되지 않았습니다." if not os.environ.get("ADMIN_PASSWORD")
               else "관리자 비밀번호가 필요합니다.")
        return self._send(403, {"error": msg, "admin": False})

    def _route(self) -> str:
        return urllib.parse.urlparse(self.path).path.rstrip("/").rsplit("/", 1)[-1]

    def _body(self) -> dict:
        n = int(self.headers.get("Content-Length") or 0)
        if not n:
            return {}
        try:
            return json.loads(self.rfile.read(n).decode())
        except Exception:
            return {}

    def do_GET(self) -> None:
        if not self._authorized():
            return self._send(401, {"error": "비밀번호가 틀렸습니다."})
        try:
            if self._route() == "state":
                return self._send(200, {"ids": load_ids(), "names": load_names(), "server": load_server()})
            if self._route() == "admin":
                return self._send(200, {"admin": True}) if self._admin() else self._deny_admin()
            if self._route() == "servers":
                return self._send(200, {"servers": CouponSite().server_names()})
            return self._send(404, {"error": "not found"})
        except Exception as e:
            return self._send(500, {"error": str(e)})

    def do_POST(self) -> None:
        if not self._authorized():
            return self._send(401, {"error": "비밀번호가 틀렸습니다."})
        route, body = self._route(), self._body()
        try:
            if route == "ids":
                ids = load_ids()
                if "entries" in body:
                    names = load_names()
                    r = merge_entries(ids, names, body["entries"], admin=self._admin())
                    if r["added"]:
                        save_ids(r["ids"])
                    if r["names"] != names:
                        save_names(r["names"])
                    return self._send(200, r)
                if "add" in body:
                    new = [x for x in clean_items(body["add"]) if x not in ids]
                    dupes = [x for x in clean_items(body["add"]) if x in ids]
                    save_ids(ids + new)
                    return self._send(200, {"added": new, "dupes": dupes, "ids": ids + new, "names": load_names()})
                if "remove" in body:
                    if not self._admin():
                        return self._deny_admin()
                    rm = set(clean_items(body["remove"]))
                    kept = [x for x in ids if x not in rm]
                    save_ids(kept)
                    names = load_names()
                    if rm & set(names):
                        names = {k: v for k, v in names.items() if k not in rm}
                        save_names(names)
                    return self._send(200, {"removed": sorted(rm & set(ids)), "ids": kept, "names": names})
                return self._send(400, {"error": "add 또는 remove 가 필요합니다."})
            if route == "server":
                if not self._admin():
                    return self._deny_admin()
                server = str(body.get("server") or "").strip()
                if not server:
                    return self._send(400, {"error": "server 가 필요합니다."})
                CouponSite().server_value(server)  # validate against live list
                save_server(server)
                return self._send(200, {"server": server})
            if route == "coupon":
                coupon = str(body.get("coupon") or "").strip()
                if not coupon:
                    return self._send(400, {"error": "쿠폰 코드가 필요합니다."})
                ids = load_ids()
                if not ids:
                    return self._send(400, {"error": "등록된 CS 코드가 없습니다."})
                return self._send(200, {"results": register_all(coupon, ids, load_server(), load_names())})
            return self._send(404, {"error": "not found"})
        except Exception as e:
            return self._send(500, {"error": str(e)})
