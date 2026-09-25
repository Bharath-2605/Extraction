#!/usr/bin/env python3
"""Multithreaded live checks for extracted URLs."""
from __future__ import annotations

import socket
import threading
import time
import re
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse, urlunparse

import requests
from requests.adapters import HTTPAdapter

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

MAX_CHECK_WORKERS = 8
PAGE_ATTEMPTS = 1
REQUEST_TIMEOUT = (4, 10)

_THREAD = threading.local()
_DNS_LOCK = threading.Lock()
_DNS_CACHE: dict[str, bool] = {}


def strip_default_port(url: str) -> str:
    try:
        parsed = urlparse(url)
    except Exception:
        return url
    if not parsed.hostname or parsed.port is None:
        return url
    scheme = (parsed.scheme or "").lower()
    if not (
        (scheme == "https" and parsed.port == 443)
        or (scheme == "http" and parsed.port == 80)
    ):
        return url
    hostname = parsed.hostname
    if ":" in hostname and not hostname.startswith("["):
        hostname = f"[{hostname}]"
    userinfo = ""
    if parsed.username:
        userinfo = parsed.username
        if parsed.password is not None:
            userinfo += f":{parsed.password}"
        userinfo += "@"
    return urlunparse(
        (
            parsed.scheme,
            f"{userinfo}{hostname}",
            parsed.path,
            parsed.params,
            parsed.query,
            parsed.fragment,
        )
    )


def _clean_url_text(value: str) -> str:
    text = re.sub(r"(?<=[A-Za-z0-9])[\s]*[''`‘’]+[\s]*(?=[A-Za-z0-9])", "", str(value))
    for ch in ("\ufeff", "\u200b", "\u200c", "\u200d", "\xa0", '"', "'", "`", "“", "”", "‘", "’", "«", "»"):
        text = text.replace(ch, "")
    return "".join(text.split())


def _peel_sentence_tail(value: str) -> str:
    """www.syf.com.Step and www.syf.com. Step 1 stay www.syf.com."""
    text = str(value or "").strip()
    if not text or text.lower().startswith("upi://"):
        return text
    try:
        from pdf_url_extractor import strip_trailing_punctuation
    except Exception:
        return text
    peeled = strip_trailing_punctuation(text)
    if not peeled:
        return text
    collapsed = _clean_url_text(peeled)
    if collapsed and collapsed != peeled:
        again = strip_trailing_punctuation(collapsed)
        if again:
            return again
    return peeled


def normalize_url(value: str | None) -> str | None:
    if value is None:
        return None
    url = _clean_url_text(_peel_sentence_tail(value))
    if not url:
        return None
    lowered = url.lower()
    if lowered.startswith(("mailto:", "javascript:", "tel:")):
        return None
    if lowered.startswith("upi://"):
        return url
    if not lowered.startswith(("http://", "https://", "ftp://")):
        url = "https://" + url
    return strip_default_port(url)


def get_domain(url: str) -> str | None:
    try:
        parsed = urlparse(url)
        return (parsed.netloc or "").lower() or None
    except Exception:
        return None


def domain_resolves(domain: str | None) -> bool:
    if not domain:
        return False
    domain = domain.split("@")[-1].split(":")[0].strip(".")
    if not domain:
        return False
    with _DNS_LOCK:
        cached = _DNS_CACHE.get(domain)
    if cached is True:
        return True

    def lookup(family: int, box: list) -> None:
        try:
            socket.getaddrinfo(domain, None, family, socket.SOCK_STREAM)
            box.append(True)
        except Exception:
            box.append(False)

    def once() -> bool:
        box: list[bool] = []
        worker = threading.Thread(target=lookup, args=(socket.AF_INET, box), daemon=True)
        worker.start()
        worker.join(2)
        return bool(box and box[0])

    ok = once()
    if ok:
        with _DNS_LOCK:
            _DNS_CACHE[domain] = True
    return ok


def _port_for_url(url: str | None) -> int | None:
    if not url:
        return None
    try:
        parsed = urlparse(url)
    except Exception:
        return None
    if parsed.port:
        return int(parsed.port)
    scheme = (parsed.scheme or "").lower()
    if scheme == "http":
        return 80
    if scheme == "https":
        return 443
    return None


def _not_working_error(reason: str, url: str | None = None, http_code=None) -> str:
    port = _port_for_url(url)
    detail = str(reason or "Unknown Error").strip()
    if detail.lower().startswith("not working"):
        detail = detail.split("-", 1)[-1].strip()
    if http_code is None:
        match = re.match(r"HTTP\s+(\d+)", detail, re.IGNORECASE)
        if match:
            http_code = int(match.group(1))
    if http_code:
        return f"Not Working- error {http_code}"
    if port and detail:
        return f"Not Working- error {detail} port {port}"
    return f"Not Working- error {detail or 'Unknown Error'}"


def _short_reason(err: object) -> str:
    text = str(err or "")
    lowered = text.casefold()
    if "nameresolutionerror" in lowered or "failed to resolve" in lowered or "nodename nor servname" in lowered:
        return "DNS lookup failed"
    if "read timed out" in lowered or "timed out" in lowered:
        return "Timeout"
    if "sslerror" in lowered or "ssl error" in lowered or "certificate" in lowered:
        return "SSL Error"
    if "connection refused" in lowered:
        return "Connection refused"
    if "connection reset" in lowered or "connection aborted" in lowered:
        return "Connection reset"
    if "max retries" in lowered:
        return "Connection failed"
    return text[:160]


def make_session() -> requests.Session:
    session = getattr(_THREAD, "session", None)
    if session is None:
        session = requests.Session()
        session.headers.update(HEADERS)
        # Use the same proxy and DNS settings as the browser. Turning this off
        # makes a live site look like a DNS failure on a company network.
        session.trust_env = True
        adapter = HTTPAdapter(pool_connections=16, pool_maxsize=16, max_retries=0)
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        _THREAD.session = session
    return session


def try_request(url: str, timeout: int = REQUEST_TIMEOUT, attempts: int = PAGE_ATTEMPTS):
    session = make_session()
    last = None
    for attempt in range(max(1, attempts)):
        box: list = []

        def run(target=url) -> None:
            try:
                box.append(session.get(target, timeout=timeout, allow_redirects=True))
            except Exception as exc:
                box.append(exc)

        worker = threading.Thread(target=run, daemon=True)
        worker.start()
        worker.join(8)
        result = box[0] if box else TimeoutError("Timeout")
        if not isinstance(result, Exception):
            return result
        last = result
        if attempt + 1 < attempts:
            time.sleep(0.3 * (attempt + 1))
    return last


def _status_from_response(response: requests.Response, dns_ok: bool) -> dict:
    code = int(response.status_code)
    if code < 400 or code in {401, 403} or (dns_ok and code < 500):
        return {
            "status": "Working",
            "error": "",
            "http_code": code,
            "final_url": response.url,
        }
    return {
        "status": "Not Working",
        "error": f"HTTP {code}",
        "http_code": code,
        "final_url": response.url,
    }


def check_url(value: str | None) -> dict:
    url = normalize_url(value)
    if not url:
        return {"status": "Not Working", "error": _not_working_error("Empty URL"), "http_code": None, "final_url": ""}
    if url.lower().startswith("upi://"):
        return {"status": "Working", "error": "", "http_code": None, "final_url": url}
    if url.lower().startswith("ftp://"):
        return {
            "status": "Not Working",
            "error": _not_working_error("FTP URLs are not HTTP-checked", url),
            "http_code": None,
            "final_url": url,
        }

    candidates = [url]
    if url.startswith("https://"):
        candidates.append("http://" + url[len("https://"):])
    elif url.startswith("http://"):
        candidates.append("https://" + url[len("http://"):])

    parsed = urlparse(url)
    host = parsed.netloc
    rest = parsed.path or ""
    if parsed.query:
        rest += "?" + parsed.query
    if host and not host.lower().startswith("www."):
        if url.startswith("https://"):
            candidates.extend([f"https://www.{host}{rest}", f"http://www.{host}{rest}"])
        else:
            candidates.extend([f"http://www.{host}{rest}", f"https://www.{host}{rest}"])

    unique_candidates: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        candidate = strip_default_port(candidate)
        if candidate not in seen:
            seen.add(candidate)
            unique_candidates.append(candidate)

    reasons: list[str] = []
    last_code = None
    for candidate in unique_candidates:
        dns_ok = domain_resolves(get_domain(candidate))
        result = try_request(candidate)
        if not isinstance(result, Exception):
            payload = _status_from_response(result, dns_ok)
            last_code = payload["http_code"]
            if payload["status"] == "Working":
                return payload
            reasons.append(payload["error"])
            continue
        err = str(result)
        short = _short_reason(result)
        if short == "DNS lookup failed":
            time.sleep(0.4)
            result = try_request(candidate)
            if not isinstance(result, Exception):
                payload = _status_from_response(result, True)
                last_code = payload["http_code"]
                if payload["status"] == "Working":
                    return payload
                reasons.append(payload["error"])
                continue
            err = str(result)
            short = _short_reason(result)
        if ("connection reset" in err.casefold() or "connection aborted" in err.casefold()) and (dns_ok or short != "DNS lookup failed"):
            return {"status": "Working", "error": "", "http_code": None, "final_url": candidate}
        if ("sslerror" in err.casefold() or "ssl error" in err.casefold() or "certificate" in err.casefold()) and dns_ok:
            return {"status": "Working", "error": "", "http_code": None, "final_url": candidate}
        reasons.append(short)

    error = _not_working_error(reasons[-1] if reasons else "Unknown Error", url, last_code)
    return {"status": "Not Working", "error": error, "http_code": last_code, "final_url": url}


def attach_url_status(
    url_rows: list[dict],
    on_progress: Callable[[int, int], None] | None = None,
    max_workers: int = MAX_CHECK_WORKERS,
) -> list[dict]:
    unique: list[str] = []
    seen: set[str] = set()
    for row in url_rows:
        url = str(row.get("URL") or "").strip()
        key = url.lower()
        if not url or key in seen:
            continue
        seen.add(key)
        unique.append(url)

    results: dict[str, dict] = {}
    if unique:
        workers = min(max_workers, max(1, len(unique)))
        completed = 0
        with ThreadPoolExecutor(max_workers=workers) as executor:
            future_map = {executor.submit(check_url, url): url for url in unique}
            for future in as_completed(future_map):
                url = future_map[future]
                try:
                    results[url.lower()] = future.result()
                except Exception as exc:
                    results[url.lower()] = {
                        "status": "Not Working",
                        "error": _not_working_error(_short_reason(exc), url),
                        "http_code": None,
                        "final_url": url,
                    }
                completed += 1
                if on_progress:
                    on_progress(completed, len(unique))

    for row in url_rows:
        raw_url = str(row.get("URL") or "").strip()
        if not raw_url:
            row["Status"] = ""
            continue
        info = results.get(raw_url.lower()) or {
            "status": "Not Working",
            "error": _not_working_error("Empty URL"),
            "http_code": None,
            "final_url": "",
        }
        if info["status"] == "Working":
            row["Status"] = "Working"
        else:
            err = info["error"] or info["status"]
            if not str(err).lower().startswith("not working"):
                err = _not_working_error(err, None, info.get("http_code"))
            row["Status"] = err
    return url_rows
