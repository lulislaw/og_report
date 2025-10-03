import json, os, time, hashlib, platform, uuid
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError
from bs4 import BeautifulSoup
import html as htmlmod

CONFIG_URL = os.getenv("AVD_KILLSWITCH_URL", "https://github.com/lulislaw/grd/blob/main/config.yaml").strip()
CACHE_TTL_SEC = 300
ENV_BYPASS = "OG_KILLSWITCH_BYPASS"
DEBUG = os.getenv("OG_KILLSWITCH_DEBUG", "1").strip() not in ("", "0", "false", "False")
UA = "OG/1.0"

_last_fetch_t = 0.0
_last_config = None


def _dbg(msg: str):
    if DEBUG:
        print(f"[remote] {msg}")


def get_device_id() -> str:
    node = platform.node()
    mac = uuid.getnode()
    raw = f"{node}|{mac}".encode("utf-8")
    return hashlib.sha256(raw).hexdigest()[:16]


def _extract_json_from_github_blob_html(html: str) -> str | None:
    try:
        soup = BeautifulSoup(html, "lxml")
    except Exception:
        soup = BeautifulSoup(html, "html.parser")

    script = soup.find("script", attrs={"type": "application/json", "data-target": "react-app.embeddedData"})
    if script:
        try:
            data = script.string or script.get_text() or ""
            obj = json.loads(data)

            def find_raw_text(x):
                if isinstance(x, dict):
                    for k, v in x.items():
                        kl = k.lower()
                        if kl in ("rawliness", "raw_lines"):  # на всякий случай
                            pass
                        if kl in ("rawlines", "lines") and isinstance(v, list):
                            return "\n".join(map(str, v))
                        if kl in ("raw", "text", "content") and isinstance(v, str):
                            return v
                        got = find_raw_text(v)
                        if got:
                            return got
                elif isinstance(x, list):
                    for it in x:
                        got = find_raw_text(it)
                        if got:
                            return got
                return None

            text = find_raw_text(obj)
            if text and text.strip():
                return text
        except Exception:
            pass

    selectors = [
        "td.blob-code.blob-code-inner",
        "td.blob-code",
        "div.Box div.Box-body table td.blob-code",
        "div.Box div.Box-body pre",
        "div[data-target='react-app.embeddedContent'] pre",
    ]
    for sel in selectors:
        cells = soup.select(sel)
        if cells:
            text = "\n".join(htmlmod.unescape(c.get_text("", strip=False)) for c in cells)
            if text.strip():
                return text

    return None


def _read_local_json(path: str) -> dict | None:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        _dbg(f"local json error: {e}")
        return None


def _fetch_http_json(url: str, timeout=5) -> dict | None:
    try:
        req = Request(url, headers={"User-Agent": UA})
        with urlopen(req, timeout=timeout) as resp:
            data = resp.read().decode("utf-8", errors="replace")
            final_url = resp.geturl()
            ctype = (resp.headers.get("Content-Type") or "").lower()

        if "application/json" in ctype or data.strip().startswith(("{", "[")):
            try:
                return json.loads(data)
            except Exception as e:
                _dbg(f"direct json decode error: {e}")

        if "github.com" in final_url and "<html" in data.lower():
            text = _extract_json_from_github_blob_html(data)
            if text:
                text = text.replace("\r\n", "\n").replace("\r", "\n")
                try:
                    return json.loads(text)
                except Exception as e:
                    _dbg(f"html->json decode error: {e}")
        return None
    except (HTTPError, URLError, TimeoutError) as e:
        _dbg(f"http error: {e}")
        return None
    except Exception as e:
        _dbg(f"other http error: {e}")
        return None


def _fetch_config(timeout=5) -> dict | None:
    global _last_fetch_t, _last_config
    now = time.time()
    if _last_config is not None and now - _last_fetch_t < CACHE_TTL_SEC:
        return _last_config

    url = CONFIG_URL
    if not url:
        _dbg("CONFIG_URL пуст")
        return None

    if url.lower().startswith("file://"):
        path = url[7:]
        cfg = _read_local_json(path)
    elif os.path.exists(url):
        cfg = _read_local_json(url)
    else:
        cfg = _fetch_http_json(url, timeout=timeout)

    if cfg is not None:
        _last_config, _last_fetch_t = cfg, now
    else:
        _dbg("не удалось получить конфиг")
    return cfg


def _norm(s) -> str:
    return str(s or "").strip().lower()


def _parse_version(v: str) -> tuple:
    parts = []
    for p in str(v).split("."):
        try:
            parts.append(int("".join(ch for ch in p if ch.isdigit()) or 0))
        except Exception:
            parts.append(0)
    return tuple(parts or [0])


def is_allowed(app_version: str, selected_okrug: str | None, device_id: str) -> tuple[bool, str]:
    if os.getenv(ENV_BYPASS):
        return True, "Bypass переменной окружения активен"

    cfg = _fetch_config(timeout=5)
    if not cfg:
        return False, "Не удалось подключиться."

    if not cfg.get("enabled", True):
        return False, cfg.get("message", "Доступ к приложению временно ограничен")

    min_v = cfg.get("min_version")
    if min_v and _parse_version(app_version) < _parse_version(min_v):
        return False, cfg.get("message", f"Требуется версия не ниже {min_v}")

    bl = {_norm(x) for x in cfg.get("blocklist", [])}
    if _norm(device_id) in bl:
        return False, cfg.get("message", "Доступ для этого устройства запрещён")

    allow_okrugs = {_norm(x) for x in cfg.get("allow_okrugs", [])}
    if allow_okrugs and selected_okrug and _norm(selected_okrug) != _norm("АВД"):
        if _norm(selected_okrug) not in allow_okrugs:
            return False, cfg.get("message", f"Округ '{selected_okrug}' сейчас не допускается")

    return True, "OK"


def guard_or_raise(app_version: str, selected_okrug: str | None) -> None:
    did = get_device_id()
    ok, msg = is_allowed(app_version, selected_okrug, device_id=did)
    if not ok:
        raise PermissionError(msg)
