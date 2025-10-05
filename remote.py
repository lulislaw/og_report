import json, os, time, hashlib, platform, uuid
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError
from bs4 import BeautifulSoup
import html as htmlmod
import yaml

# --- настройки ---
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


# ---------- извлечение текста из HTML GitHub ----------
def _extract_text_from_github_blob_html(html: str) -> str | None:
    """
    Пытаемся достать «сырой» текст файла со страницы GitHub-блоба (HTML),
    не используя raw.*. Идем по нескольким стратегиям:
      1) embeddedData (React) — иногда содержит raw/lines
      2) таблица/пре с кодом — собираем строки и unescape
    """
    try:
        soup = BeautifulSoup(html, "lxml")
    except Exception:
        soup = BeautifulSoup(html, "html.parser")

    # 1) React embedded JSON (встречается у GitHub)
    script = soup.find("script", attrs={"type": "application/json", "data-target": "react-app.embeddedData"})
    if script:
        try:
            data = script.string or script.get_text() or ""
            obj = json.loads(data)

            def find_raw_text(x):
                if isinstance(x, dict):
                    for k, v in x.items():
                        kl = k.lower()
                        # иногда встречаются поля вроде rawLines / lines / content / text
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
        except Exception as e:
            _dbg(f"embeddedData parse error: {e}")

    # 2) Падаем к селекторам с кодом
    selectors = [
        "td.blob-code.blob-code-inner",
        "td.blob-code",
        "div.Box div.Box-body table td.blob-code",
        "div.Box div.Box-body pre",
        "div[data-target='react-app.embeddedContent'] pre",
        "table.js-file-line-container td.js-file-line",
        "div.highlight pre",  # на всякий случай
    ]
    for sel in selectors:
        cells = soup.select(sel)
        if cells:
            text = "\n".join(htmlmod.unescape(c.get_text("", strip=False)) for c in cells)
            if text.strip():
                return text

    # 3) если вдруг контент целиком в <pre> без селекторов выше
    pre = soup.find("pre")
    if pre:
        text = htmlmod.unescape(pre.get_text("", strip=False))
        if text.strip():
            return text

    return None


# ---------- локальные файлы ----------
def _read_local_text(path: str) -> str | None:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        _dbg(f"local read error: {e}")
        return None


def _fetch_http_text(url: str, timeout=5) -> tuple[str | None, str | None]:
    """
    Универсальная HTTP-качалка текста. Для github blob зовет спец-обработчик.
    Возвращает (text, content_type).
    """
    if "github.com" in url and "/blob/" in url:
        return _fetch_http_github_blob(url, timeout=timeout)

    try:
        req = Request(url, headers={"User-Agent": UA})
        with urlopen(req, timeout=timeout) as resp:
            data = resp.read().decode("utf-8", errors="replace")
            ctype = (resp.headers.get("Content-Type") or "").lower()
        return data, ctype
    except (HTTPError, URLError, TimeoutError) as e:
        _dbg(f"http error: {e}")
        return None, None
    except Exception as e:
        _dbg(f"other http error: {e}")
        return None, None


# ---------- загрузка и парсинг конфига ----------
def _fetch_config(timeout=5) -> dict | None:
    global _last_fetch_t, _last_config
    now = time.time()
    if _last_config is not None and now - _last_fetch_t < CACHE_TTL_SEC:
        return _last_config

    url = CONFIG_URL
    if not url:
        _dbg("CONFIG_URL пуст")
        return None

    # Bypass через переменную окружения (если нужно)
    if os.getenv(ENV_BYPASS, "").strip() not in ("", "0", "false", "False"):
        _dbg(f"{ENV_BYPASS}=1 — принудительный пропуск")
        _last_config, _last_fetch_t = {"allow": True, "reason": "bypass"}, now
        return _last_config

    # file:// или локальный путь
    if url.lower().startswith("file://"):
        path = url[7:]
        text = _read_local_text(path)
        cfg = _decode_config_text(text or "", None, file_hint=path) if text else None
    elif os.path.exists(url):
        text = _read_local_text(url)
        cfg = _decode_config_text(text or "", None, file_hint=url) if text else None
    else:
        text, ctype = _fetch_http_text(url, timeout=timeout)
        cfg = _decode_config_text(text or "", ctype, file_hint=url) if text else None

    if cfg is not None:
        _last_config, _last_fetch_t = cfg, now
    else:
        _dbg("не удалось получить/распарсить конфиг")
    return cfg


def _extract_text_from_github_blob_html(html: str) -> str | None:
    # Возвращаем ТОЛЬКО реальный текст файла. Никаких стилей/шапок.
    try:
        soup = BeautifulSoup(html, "lxml")
    except Exception:
        soup = BeautifulSoup(html, "html.parser")

    def looks_like_file_text(t: str) -> bool:
        s = t.strip()
        if not s:
            return False
        bad_markers = ("<!DOCTYPE html", "<html", "<head", "<body", ":root{", "--tab-size-")
        return not any(m.lower() in s.lower() for m in bad_markers)

    # 0) Самый надёжный способ: GitHub кладёт сырой текст в data-snippet-clipboard-copy-content
    node = soup.find(attrs={"data-snippet-clipboard-copy-content": True})
    if node:
        raw = node.get("data-snippet-clipboard-copy-content", "")
        if looks_like_file_text(raw):
            return raw

    # 1) React embedded JSON
    script = soup.find("script", attrs={"type": "application/json", "data-target": "react-app.embeddedData"})
    if script:
        try:
            data = script.string or script.get_text() or ""
            obj = json.loads(data)

            def find_raw_text(x):
                if isinstance(x, dict):
                    for k, v in x.items():
                        kl = k.lower()
                        if kl in ("rawlines", "lines") and isinstance(v, list):
                            txt = "\n".join(map(str, v))
                            if looks_like_file_text(txt):
                                return txt
                        if kl in ("raw", "text", "content") and isinstance(v, str):
                            if looks_like_file_text(v):
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
            if text:
                return text
        except Exception as e:
            _dbg(f"embeddedData parse error: {e}")

    # 2) Табличные ячейки кода
    selectors = [
        "table.js-file-line-container td.js-file-line",
        "td.blob-code.blob-code-inner",
        "td.blob-code",
        "div[data-target='react-app.embeddedContent'] pre",
        "div.Box div.Box-body pre",
        "div.highlight pre",
        "pre"
    ]
    for sel in selectors:
        cells = soup.select(sel)
        if not cells:
            continue
        # Собираем строки, игнорируя пустяки
        text = "\n".join(htmlmod.unescape(c.get_text("", strip=False)) for c in cells)
        if looks_like_file_text(text):
            return text

    return None


def _fetch_http_github_blob(url: str, timeout=5) -> tuple[str | None, str | None]:
    # Порядок: ?plain=1 → HTML-парсинг → (если всё равно HTML/CSS) → None
    candidates = []
    if "github.com" in url and "/blob/" in url:
        plain = url + ("&plain=1" if "?" in url else "?plain=1")
        candidates.append(plain)
    candidates.append(url)

    def is_html_like(s: str, ctype: str) -> bool:
        s_l = s.lower()
        return "text/html" in (ctype or "") or "<html" in s_l or "<!doctype html" in s_l

    for u in candidates:
        try:
            req = Request(u, headers={"User-Agent": UA})
            with urlopen(req, timeout=timeout) as resp:
                data = resp.read().decode("utf-8", errors="replace")
                ctype = (resp.headers.get("Content-Type") or "").lower()
                final_url = resp.geturl()

            # Если это «plain», но пришёл HTML — вытаскиваем из DOM
            if is_html_like(data, ctype):
                extracted = _extract_text_from_github_blob_html(data)
                if extracted:
                    return extracted, "text/plain"
                # Иногда редирект на логин/прочее — ещё одна попытка с plain на final_url
                if "github.com" in final_url and "/blob/" in final_url and "plain=1" not in final_url:
                    plain2 = final_url + ("&plain=1" if "?" in final_url else "?plain=1")
                    try:
                        req2 = Request(plain2, headers={"User-Agent": UA})
                        with urlopen(req2, timeout=timeout) as resp2:
                            data2 = resp2.read().decode("utf-8", errors="replace")
                            ctype2 = (resp2.headers.get("Content-Type") or "").lower()
                        if is_html_like(data2, ctype2):
                            extracted2 = _extract_text_from_github_blob_html(data2)
                            if extracted2:
                                return extracted2, "text/plain"
                        else:
                            return data2, ctype2
                    except Exception as e2:
                        _dbg(f"second plain fetch error: {e2}")
                # Ничего не достали
                return None, None

            # Не HTML — считаем текстом файла
            return data, ctype

        except (HTTPError, URLError, TimeoutError) as e:
            _dbg(f"http error ({u}): {e}")
        except Exception as e:
            _dbg(f"other http error ({u}): {e}")

    return None, None


def _decode_config_text(data: str, content_type: str | None, file_hint: str | None = None) -> dict | None:
    text = (data or "").strip()
    if not text:
        return None

    # Защита: если это HTML/CSS, не пытаемся парсить как YAML/JSON
    lower = text.lower()
    if "<html" in lower or "<!doctype html" in lower or ":root{" in lower or "--tab-size-" in lower:
        _dbg("получен HTML/CSS вместо файла — пропускаем декодирование")
        return None

    hint = (file_hint or "").lower()
    ctype = (content_type or "").lower()

    # JSON?
    if hint.endswith(".json") or "application/json" in ctype or text.startswith(("{", "[")):
        try:
            return json.loads(text)
        except Exception as e:
            _dbg(f"json decode error: {e}")

    # YAML?
    looks_like_yaml = any([
        hint.endswith(".yaml"),
        hint.endswith(".yml"),
        "yaml" in ctype or "x-yaml" in ctype,
        (":" in text and "\n" in text and not text.strip().startswith("<")),  # простая эвристика
    ])
    if looks_like_yaml:
        try:
            obj = yaml.safe_load(text)
            if isinstance(obj, dict):
                return obj
            if isinstance(obj, list):
                try:
                    return dict(obj)
                except Exception:
                    return {"_list": obj}
        except Exception as e:
            _dbg(f"yaml decode error: {e}")

    # Последняя попытка — JSON ещё раз
    try:
        return json.loads(text)
    except Exception:
        return None


def _ensure_parent_dir(path: str) -> None:
    d = os.path.dirname(os.path.abspath(path))
    if d:
        os.makedirs(d, exist_ok=True)


def _bytes_sha1(b: bytes) -> str:
    h = hashlib.sha1()
    h.update(b)
    return h.hexdigest()


def _persist_config(cfg: dict, dst_path: str = "resource/config.yaml") -> str:
    """
    Сохраняет cfg в YAML по пути dst_path:
    - атомарная запись через *.tmp и os.replace
    - создаёт каталоги при необходимости
    - не перезаписывает файл, если содержимое не изменилось (по SHA1)
    Возвращает фактический путь сохранения.
    """
    # сериализуем в YAML (fallback в JSON, если что-то пойдёт не так)
    try:
        text = yaml.safe_dump(cfg, allow_unicode=True, sort_keys=False)
    except Exception:
        text = json.dumps(cfg, ensure_ascii=False, indent=2)

    data = text.encode("utf-8")

    _ensure_parent_dir(dst_path)

    # если файл уже существует с тем же содержимым — пропускаем запись
    try:
        with open(dst_path, "rb") as f:
            old = f.read()
        if _bytes_sha1(old) == _bytes_sha1(data):
            _dbg(f"config unchanged, skip write: {dst_path}")
            return dst_path
    except FileNotFoundError:
        pass

    tmp_path = dst_path + ".tmp"
    with open(tmp_path, "wb") as f:
        f.write(data)
        f.flush()
        os.fsync(f.fileno())

    os.replace(tmp_path, dst_path)  # атомарная замена
    _dbg(f"config saved to {dst_path}")
    return dst_path


def _clear_config(dst_path: str = "resource/config.yaml") -> None:
    # удаляем сам файл и возможный временный
    for p in (dst_path, dst_path + ".tmp"):
        try:
            if os.path.exists(p):
                os.remove(p)
                _dbg(f"config removed: {p}")
        except Exception as e:
            _dbg(f"config remove error ({p}): {e}")


def _invalidate_cache():
    global _last_config, _last_fetch_t
    _last_config = None
    _last_fetch_t = 0.0


def is_allowed():
    cfg = _fetch_config(timeout=5)
    if not cfg:
        return False, "Не удалось подключиться."
    return True, cfg


def dwnl_cfg(save_path: str = "resource/config.yaml"):
    ok, msg = is_allowed()
    if ok:
        # msg здесь — это сам cfg (dict)
        try:
            _persist_config(msg, save_path)
        except Exception as e:
            _dbg(f"save error: {e}")
        return msg  # возвращаем cfg
    else:
        _clear_config(save_path)
        _invalidate_cache()
        return None
