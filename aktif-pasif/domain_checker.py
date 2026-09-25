#!/usr/bin/env python3
"""XLSX A sutunundaki domain/URL'leri kontrol eden masaustu GUI."""

from __future__ import annotations

import concurrent.futures
import datetime as dt
import os
import queue
import re
import shutil
import socket
import ssl
import subprocess
import sys
import tempfile
import threading
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
import webbrowser
from html.parser import HTMLParser
from pathlib import Path
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    TK_AVAILABLE = True
except ImportError:
    class _TkUnavailable:
        Tk = object

    tk = _TkUnavailable()
    filedialog = messagebox = ttk = None
    TK_AVAILABLE = False

APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
APP_NAME = "Domain Aktif / Pasif Kontrol"
NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkg": "http://schemas.openxmlformats.org/package/2006/relationships",
}
TIMEOUT_SECONDS = 12
MAX_WORKERS = 8
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36"
REQUEST_PROFILES = (
    (
        "Chrome",
        {
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Sec-Fetch-User": "?1",
        },
    ),
    (
        "Firefox",
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:155.0) Gecko/20100101 Firefox/155.0",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Sec-Fetch-User": "?1",
        },
    ),
    (
        "Edge",
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36 Edg/154.0.0.0",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Sec-Fetch-User": "?1",
        },
    ),
)
try:
    import httpx as _httpx
except ImportError:
    _httpx = None
MAX_PAGE_BYTES = 512 * 1024
SUMMARY_CHARS = 280


def _column_number(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref.upper())
    value = 0
    if letters:
        for char in letters.group():
            value = value * 26 + ord(char) - 64
    return value


def read_xlsx_column_a(path: Path) -> list[str]:
    """Read the first worksheet, column A, from a normal .xlsx file."""
    with zipfile.ZipFile(path) as book:
        workbook = ET.fromstring(book.read("xl/workbook.xml"))
        sheet = workbook.find("main:sheets/main:sheet", NS)
        if sheet is None:
            return []
        rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
        rel_id = sheet.attrib.get(f"{{{NS['rel']}}}id")
        target = next((r.attrib["Target"] for r in rels.findall("pkg:Relationship", NS)
                       if r.attrib.get("Id") == rel_id), None)
        if not target:
            raise ValueError("Çalışma sayfası bulunamadı.")
        target = target.replace("\\", "/")
        if target.startswith("/"):
            sheet_path = target.lstrip("/")
        elif target.startswith("xl/"):
            sheet_path = target
        else:
            sheet_path = "xl/" + target
        shared: list[str] = []
        if "xl/sharedStrings.xml" in book.namelist():
            root = ET.fromstring(book.read("xl/sharedStrings.xml"))
            for item in root.findall("main:si", NS):
                shared.append("".join(t.text or "" for t in item.findall(".//main:t", NS)))
        root = ET.fromstring(book.read(sheet_path))
        values: list[tuple[int, str]] = []
        for cell in root.findall(".//main:sheetData/main:row/main:c", NS):
            ref = cell.attrib.get("r", "")
            if _column_number(ref) != 1:
                continue
            kind = cell.attrib.get("t")
            if kind == "inlineStr":
                val = "".join(t.text or "" for t in cell.findall(".//main:t", NS))
            else:
                node = cell.find("main:v", NS)
                val = node.text or "" if node is not None else ""
                if kind == "s" and val:
                    val = shared[int(val)]
            row_num = int(re.search(r"\d+", ref).group()) if re.search(r"\d+", ref) else 0
            values.append((row_num, val.strip()))
        values.sort(key=lambda pair: pair[0])
        # Skip a column heading if it is clearly a heading; retain domain-like first cells.
        if values and values[0][1].casefold() in {"url", "urls", "domain", "domains", "site", "website", "link", "address", "adres", "web address", "asset", "assets", "host", "hostname", "subdomain", "fqdn"}:
            values = values[1:]
        return [value for _, value in values if value]


def normalize_url(raw: str) -> str:
    value = raw.strip()
    if not value:
        raise ValueError("Boş URL")
    if not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", value):
        value = "https://" + value
    parts = urllib.parse.urlsplit(value)
    if parts.scheme.lower() not in ("http", "https") or not parts.hostname:
        raise ValueError("HTTP/HTTPS adresi değil")
    if parts.username or parts.password:
        raise ValueError("Giriş bilgisi içeren URL kontrol için kabul edilmiyor")
    # Accessing .port detects malformed port values.
    _ = parts.port
    return urllib.parse.urlunsplit((parts.scheme.lower(), parts.netloc, parts.path or "/", parts.query, ""))


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


class PageTextParser(HTMLParser):
    """Extract a document title, meta description, and readable text snippet."""
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.in_title = False
        self.hidden_depth = 0
        self.title_parts: list[str] = []
        self.text_parts: list[str] = []
        self.description = ""

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        attrs = {key.lower(): value or "" for key, value in attrs}
        if tag == "title":
            self.in_title = True
        elif tag in ("script", "style", "noscript", "svg", "nav", "footer"):
            self.hidden_depth += 1
        elif tag == "meta" and (attrs.get("name", "").casefold() == "description" or
                                 attrs.get("property", "").casefold() == "og:description"):
            if not self.description:
                self.description = attrs.get("content", "").strip()

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == "title":
            self.in_title = False
        elif tag in ("script", "style", "noscript", "svg", "nav", "footer") and self.hidden_depth:
            self.hidden_depth -= 1

    def handle_data(self, data):
        text = " ".join(data.split())
        if not text:
            return
        if self.in_title:
            self.title_parts.append(text)
        elif not self.hidden_depth:
            self.text_parts.append(text)


def extract_page_info(body: bytes, headers) -> tuple[str, str]:
    content_type = headers.get("Content-Type", "") if headers else ""
    if content_type and "html" not in content_type.casefold() and "xhtml" not in content_type.casefold():
        return "", "HTML sayfası değil"
    charset = "utf-8"
    match = re.search(r"charset\s*=\s*[\"']?([^;\s\"']+)", content_type, re.I)
    if match:
        charset = match.group(1)
    try:
        html = body.decode(charset, errors="replace")
    except LookupError:
        html = body.decode("utf-8", errors="replace")
    parser = PageTextParser()
    try:
        parser.feed(html)
    except Exception:
        pass
    title = " ".join(parser.title_parts).strip()
    content = parser.description or " ".join(parser.text_parts)
    content = re.sub(r"\s+", " ", content).strip()
    if len(content) > SUMMARY_CHARS:
        content = content[:SUMMARY_CHARS].rsplit(" ", 1)[0].rstrip(" .,;:") + "…"
    return title, content or "Özet alınamadı"


def _urllib_request(url: str, headers=None, ssl_context=None):
    handlers = [NoRedirect()]
    if ssl_context is not None:
        handlers.append(urllib.request.HTTPSHandler(context=ssl_context))
    opener = urllib.request.build_opener(*handlers)
    req = urllib.request.Request(url, headers=headers or REQUEST_PROFILES[0][1], method="GET")
    try:
        response = opener.open(req, timeout=TIMEOUT_SECONDS)
        try:
            body = response.read(MAX_PAGE_BYTES)
            return response.status, response.headers.get("Location"), response.geturl(), body, response.headers
        finally:
            response.close()
    except urllib.error.HTTPError as exc:
        try:
            body = exc.read(MAX_PAGE_BYTES)
        except Exception:
            body = b""
        return exc.code, exc.headers.get("Location") if exc.headers else None, exc.geturl(), body, exc.headers


def _httpx_request(url: str, headers=None, ssl_context=None):
    """Use HTTPX when installed; keep its response body bounded like the other clients."""
    if _httpx is None:
        raise RuntimeError("HTTPX yüklü değil")
    try:
        with _httpx.Client(
            verify=ssl_context if ssl_context is not None else True,
            follow_redirects=False,
            timeout=TIMEOUT_SECONDS,
            trust_env=True,
            headers=headers or REQUEST_PROFILES[0][1],
        ) as client:
            with client.stream("GET", url) as response:
                body = bytearray()
                for chunk in response.iter_bytes():
                    remaining = MAX_PAGE_BYTES - len(body)
                    if remaining <= 0:
                        break
                    body.extend(chunk[:remaining])
                    if len(body) >= MAX_PAGE_BYTES:
                        break
                response_headers = {name.title(): value for name, value in response.headers.items()}
                return (
                    response.status_code,
                    response.headers.get("Location"),
                    str(response.url),
                    bytes(body),
                    response_headers,
                )
    except _httpx.RequestError as exc:
        raise OSError(str(exc)) from exc


def _curl_request(url: str, headers=None):
    """Use the OS curl binary as a second HTTP/TLS implementation when available."""
    executable = shutil.which("curl.exe") or shutil.which("curl")
    if not executable:
        raise RuntimeError("Sistemde curl bulunamadı; alternatif HTTP istemcisi kullanılamıyor")
    headers = headers or REQUEST_PROFILES[0][1]
    temp_path = None
    header_path = None
    try:
        with tempfile.NamedTemporaryFile(prefix="domain-check-", suffix=".html", delete=False) as output:
            temp_path = output.name
        with tempfile.NamedTemporaryFile(prefix="domain-check-", suffix=".headers", delete=False) as output:
            header_path = output.name
        command = [
            # Follow redirects in check_url, not curl: the app must validate and
            # record every hop (including its response code) before requesting it.
            executable, "--silent", "--show-error",
            "--connect-timeout", "10", "--max-time", "22",
            "--dump-header", header_path,
            "--user-agent", headers.get("User-Agent", USER_AGENT),
            "--output", temp_path,
        ]
        for name, value in headers.items():
            if name.casefold() != "user-agent":
                command.extend(["--header", f"{name}: {value}"])
        command.extend(["--write-out", "%{http_code}\\n%{url_effective}\\n%{content_type}", url])
        process_options = {}
        if os.name == "nt":
            # curl is a console program. Keep each fallback request from opening
            # a visible console window when this app was started without one.
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            process_options = {
                "startupinfo": startupinfo,
                "creationflags": subprocess.CREATE_NO_WINDOW,
            }
        completed = subprocess.run(
            command, capture_output=True, timeout=25, check=False, **process_options
        )
        metadata = completed.stdout.decode("utf-8", errors="replace").splitlines()
        code_text = metadata[0].strip() if metadata else "000"
        final_url = metadata[1].strip() if len(metadata) > 1 else url
        content_type = metadata[2].strip() if len(metadata) > 2 else ""
        if not code_text.isdigit() or code_text == "000":
            error = completed.stderr.decode("utf-8", errors="replace").strip()
            raise OSError(error or "curl yanıt kodu alamadı")
        with open(temp_path, "rb") as page:
            body = page.read(MAX_PAGE_BYTES)
        response_headers = {"Content-Type": content_type}
        try:
            with open(header_path, "rb") as header_file:
                header_blocks = [block for block in header_file.read().decode("iso-8859-1", errors="replace").split("\r\n\r\n") if block.strip()]
            if header_blocks:
                for line in header_blocks[-1].splitlines()[1:]:
                    if ":" in line:
                        name, value = line.split(":", 1)
                        response_headers[name.strip().title()] = value.strip()
        except OSError:
            pass
        return int(code_text), response_headers.get("Location"), final_url, body, response_headers
    except subprocess.TimeoutExpired as exc:
        raise TimeoutError("curl isteği 25 saniyede tamamlanmadı") from exc
    finally:
        if temp_path:
            try:
                os.unlink(temp_path)
            except OSError:
                pass
        if header_path:
            try:
                os.unlink(header_path)
            except OSError:
                pass


def _is_certificate_error(message: object) -> bool:
    text = str(message).casefold()
    cert_words = ("certificate", "cert verify", "cert_verify", "cert-verification", "x509", "issuer")
    cause_words = ("verify", "invalid", "expired", "issuer", "self signed", "self-signed", "trust", "authority", "problem", "unable", "unknown", "revoked")
    return any(word in text for word in cert_words) and any(word in text for word in cause_words)


def _response_note(attempts, selected_index: int) -> str:
    if not attempts:
        return ""
    codes = [str(response[0]) for _, response in attempts]
    if len(set(codes)) > 1:
        observed = ", ".join(f"{label}: {response[0]}" for label, response in attempts)
        selected_code = attempts[selected_index][1][0]
        return (
            f"İstekler farklı yanıt verdi ({observed}); erişim sağlayan HTTP {selected_code} yanıtı kullanıldı. "
            "Bu fark site güvenlik filtresi, IP/VPN veya tarayıcı oturumundan kaynaklanabilir."
        )
    if codes and all(code == "403" for code in codes):
        observed = ", ".join(label for label, _ in attempts)
        return (
            f"403 yanıtı {observed} isteklerinde tekrarlandı. Sunucuya ulaşıldı ancak bu denetleyiciye erişim reddedildi; "
            "tarayıcıdaki oturum/çerezler, IP/VPN veya site güvenlik filtresi farklı yanıt oluşturabilir."
        )
    return ""


def _request(url: str):
    """Try independent clients/profiles; diagnose certificate failures without hiding them."""
    chrome_headers = REQUEST_PROFILES[0][1]
    firefox_headers = REQUEST_PROFILES[1][1]
    edge_headers = REQUEST_PROFILES[2][1]
    if _httpx is not None:
        plan = [
            ("HTTPX/Chrome", "httpx", chrome_headers),
            ("urllib/Firefox", "urllib", firefox_headers),
            ("curl/Edge", "curl", edge_headers),
        ]
    else:
        plan = [
            ("urllib/Chrome", "urllib", chrome_headers),
            ("urllib/Firefox", "urllib", firefox_headers),
            ("curl/Edge", "curl", edge_headers),
        ]
    failures = []
    attempts = []
    for index, (label, client_kind, headers) in enumerate(plan):
        try:
            if client_kind == "httpx":
                response = _httpx_request(url, headers=headers)
            elif client_kind == "curl":
                response = _curl_request(url, headers=headers)
            else:
                response = _urllib_request(url, headers=headers)
        except (urllib.error.URLError, TimeoutError, socket.timeout, OSError) as exc:
            reason = getattr(exc, "reason", exc)
            failures.append((label, str(reason)))
            continue
        except (RuntimeError, subprocess.SubprocessError) as exc:
            failures.append((label, str(exc)))
            continue
        attempts.append((label, response))
        code = response[0]
        if 200 <= code < 400:
            note = _response_note(attempts, len(attempts) - 1)
            if failures and not note:
                failed_label, failed_reason = failures[0]
                note = f"{failed_label} isteği yanıt alamadı; başka istemci HTTP {code} yanıtı aldı ({failed_reason[:180]})."
            tls_note = ""
            if any(_is_certificate_error(reason) for _, reason in failures) and client_kind == "curl":
                tls_note = "Python sertifika kontrolü başarısız oldu; curl kendi doğrulamasıyla yanıt aldı"
            return (*response, tls_note, note)
        # A 403 is often specific to a request profile or HTTP client. Check the remaining profiles.
        if index == 0 and code != 403:
            return (*response, "", "")
    if attempts:
        selected_index = next((i for i, (_, response) in enumerate(attempts) if response[0] != 403), 0)
        note = _response_note(attempts, selected_index)
        return (*attempts[selected_index][1], "", note)
    if any(_is_certificate_error(reason) for _, reason in failures) and not urllib.parse.urlsplit(url).query:
        try:
            response = _urllib_request(url, headers=chrome_headers, ssl_context=ssl._create_unverified_context())
            return (*response, "Sertifika doğrulanamadı; HTTP yanıtı doğrulanmamış TLS üzerinden alındı", "")
        except (urllib.error.URLError, TimeoutError, socket.timeout, OSError) as exc:
            failures.append(("Sertifika tanılama", str(getattr(exc, "reason", exc))))
    raise OSError("; ".join(f"{label}: {reason}" for label, reason in failures))


def _availability_label(code: object, state: str, was_redirected: bool = False) -> str:
    """Explain reachability separately from whether the requested page opened."""
    if state != "Aktif":
        return {
            "Ulaşılamıyor": "Bu ağdan yanıt alınamadı",
            "Engellendi": "Atlandı · DNS/özel ağ koruması",
            "Geçersiz": "Geçersiz adres",
            "Belirsiz": "Belirsiz sonuç",
        }.get(state, state)
    try:
        status = int(code)
    except (TypeError, ValueError):
        return "Belirsiz sonuç"
    if 200 <= status < 300:
        return "Aktif · yönlendirildi" if was_redirected else "Aktif"
    if status == 401:
        return "Aktif · oturum gerekli"
    if status == 403:
        return "Yanıt veriyor · erişim kısıtlı"
    if status in (404, 410):
        return "Yanıt veriyor · sayfa bulunamadı"
    if 300 <= status < 400:
        return "Yanıt veriyor · yönlendirme tamamlanmadı"
    if 400 <= status < 500:
        return f"Yanıt veriyor · HTTP {status}"
    if 500 <= status < 600:
        return f"Yanıt veriyor · sunucu hatası ({status})"
    return "Yanıt veriyor"


def _is_application_error_page(title: str, summary: str) -> bool:
    """Catch common soft-error pages that incorrectly return HTTP 200."""
    page_title = title.casefold()
    page_summary = summary.casefold()
    title_markers = ("hata sayfası", "404 not found", "page not found", "error page")
    return (
        any(marker in page_title for marker in title_markers)
        or "lütfen adresi kontrol ederek tekrar deneyiniz" in page_summary
    )


def check_url(raw: str, url_guard=None) -> dict[str, object]:
    result = {"url": raw, "final_url": "", "code": "Kod yok", "state": "Belirsiz", "availability": "Belirsiz sonuç", "response_chain": [], "redirect_chain": [], "detail": "", "title": "", "summary": "", "tls": "—"}
    try:
        url = normalize_url(raw)
        result["url"] = url
    except Exception as exc:
        result["state"], result["detail"] = "Geçersiz", str(exc)
        result["availability"] = _availability_label(result["code"], str(result["state"]))
        return result
    current = url
    seen = set()
    request_notes = []
    scheme_fallback_used = False
    scheme_fallback_tried = False
    for _ in range(6):
        if current in seen:
            result["state"], result["detail"] = "Belirsiz", "Yönlendirme döngüsü"
            result["final_url"] = current
            return result
        seen.add(current)
        if url_guard is not None:
            try:
                url_guard(current)
            except Exception as exc:
                result["state"] = "Engellendi"
                result["code"] = "Kontrol edilmedi"
                result["detail"] = f"Güvenli tarama nedeniyle istek gönderilmedi: {exc}"
                result["final_url"] = current
                result["availability"] = _availability_label(result["code"], str(result["state"]))
                return result
        try:
            code, location, final, body, headers, tls_note, request_note = _request(current)
            if tls_note:
                result["tls"] = "Sertifika uyarısı" if "doğrulanmamış TLS" in tls_note else "Sertifika farkı"
                result["detail"] = tls_note
            if request_note:
                request_notes.append(request_note)
        except (urllib.error.URLError, TimeoutError, socket.timeout, OSError) as exc:
            reason = getattr(exc, "reason", exc)
            parts = urllib.parse.urlsplit(current)
            can_try_other_scheme = (
                current == url and not scheme_fallback_tried and not parts.query
                and parts.path in ("", "/") and parts.hostname
            )
            if can_try_other_scheme:
                scheme_fallback_tried = True
                other_scheme = "http" if parts.scheme == "https" else "https"
                alternate = urllib.parse.urlunsplit((other_scheme, parts.netloc, parts.path or "/", "", ""))
                if url_guard is not None:
                    try:
                        url_guard(alternate)
                    except Exception as guard_exc:
                        result["state"] = "Engellendi"
                        result["code"] = "Kontrol edilmedi"
                        result["detail"] = f"Güvenli tarama nedeniyle istek gönderilmedi: {guard_exc}"
                        result["final_url"] = alternate
                        result["availability"] = _availability_label(result["code"], str(result["state"]))
                        return result
                try:
                    code, location, final, body, headers, tls_note, request_note = _request(alternate)
                    current = alternate
                    scheme_fallback_used = True
                    if tls_note:
                        result["tls"] = "Sertifika uyarısı" if "doğrulanmamış TLS" in tls_note else "Sertifika farkı"
                        result["detail"] = tls_note
                    if request_note:
                        request_notes.append(request_note)
                except (urllib.error.URLError, TimeoutError, socket.timeout, OSError) as alternate_exc:
                    alternate_reason = getattr(alternate_exc, "reason", alternate_exc)
                    result["state"] = "Ulaşılamıyor"
                    result["code"] = "Kod yok"
                    result["detail"] = f"HTTP yanıtı alınamadı. İlk deneme: {reason}. {other_scheme.upper()} denemesi: {alternate_reason}"
                    result["tls"] = "Sertifika sorunu" if _is_certificate_error(reason) or _is_certificate_error(alternate_reason) else "Bağlantı sorunu"
                    result["final_url"] = alternate
                    result["availability"] = _availability_label(result["code"], str(result["state"]))
                    return result
            else:
                result["state"] = "Ulaşılamıyor"
                result["code"] = "Kod yok"
                result["detail"] = f"HTTP yanıtı alınamadı — {reason}"
                result["tls"] = "Sertifika sorunu" if _is_certificate_error(reason) else "Bağlantı sorunu"
                result["final_url"] = current
                result["availability"] = _availability_label(result["code"], str(result["state"]))
                return result
        result["code"] = str(code)
        response_chain = result["response_chain"]
        if isinstance(response_chain, list):
            response_chain.append(str(code))
        result["final_url"] = final or current
        if 300 <= code < 400 and location:
            next_url = urllib.parse.urljoin(current, location)
            try:
                next_url = normalize_url(next_url)
                redirect_chain = result["redirect_chain"]
                if isinstance(redirect_chain, list):
                    redirect_chain.append({"code": str(code), "from": current, "to": next_url})
                current = next_url
                continue
            except ValueError:
                result["state"], result["detail"] = "Belirsiz", "Geçersiz yönlendirme adresi"
                return result
        result["title"], result["summary"] = extract_page_info(body, headers)
        application_error_page = (
            200 <= code < 300
            and _is_application_error_page(str(result["title"]), str(result["summary"]))
        )
        if result["tls"] not in ("Sertifika uyarısı", "Sertifika farkı"):
            result["tls"] = "Güvenli HTTPS" if urllib.parse.urlsplit(current).scheme == "https" else "Şifrelenmemiş HTTP"
        if 200 <= code < 400:
            result["state"] = "Aktif"
            result["detail"] = "Site yanıt veriyor"
        elif code == 401:
            result["state"] = "Aktif"
            result["detail"] = "Site yanıt veriyor; oturum açma gerekiyor (HTTP 401)"
        elif code == 403:
            result["state"] = "Aktif"
            result["detail"] = "Sunucu HTTP 403 ile erişimi reddetti; bu, domainin kapalı olduğu anlamına gelmez. Tarayıcı oturumu, IP/VPN veya site güvenlik filtresi farklı yanıt verebilir."
        elif code in (404, 410):
            result["state"] = "Aktif"
            result["detail"] = "Sunucu yanıt veriyor; bu URL sayfası bulunamadı veya kaldırıldı"
        elif 400 <= code < 500:
            result["state"] = "Aktif"
            result["detail"] = "Sunucu yanıt veriyor; HTTP istemci hata kodu döndü"
        else:
            result["state"] = "Aktif"
            result["detail"] = "Sunucu yanıt veriyor; HTTP sunucu hatası döndü"
        if scheme_fallback_used:
            result["detail"] = f"İlk protokol yanıt vermedi; {urllib.parse.urlsplit(current).scheme.upper()} ile erişildi. {result['detail']}"
        elif result["tls"] == "Sertifika uyarısı":
            result["detail"] = f"{result['detail']} — Sertifika doğrulamasını tarayıcı gibi geçersiz kılıp yalnızca yanıtı teşhis etti."
        elif result["tls"] == "Sertifika farkı":
            result["detail"] = f"{result['detail']} — Curl sertifikayı doğrulayarak HTTP yanıtı alabildi."
        if application_error_page:
            result["detail"] = (
                f"HTTP {code} yanıtı alındı, ancak sayfanın içeriği uygulama hata sayfası gösteriyor."
            )
        if request_notes:
            result["detail"] = f"{result['detail']} İstek karşılaştırması: {' '.join(dict.fromkeys(request_notes))}"
        result["availability"] = (
            "Yanıt veriyor · hata sayfası"
            if application_error_page
            else _availability_label(
                result["code"], str(result["state"]), len(result["response_chain"]) > 1
            )
        )
        return result
    result["state"], result["detail"] = "Belirsiz", "Çok fazla yönlendirme"
    result["availability"] = _availability_label(result["code"], str(result["state"]))
    return result


def _xml_escape(value: object) -> str:
    return (str(value).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def write_results_xlsx(path: Path, rows: list[dict[str, object]]) -> None:
    headers = ["Kaynak", "Site / URL", "İlk Yanıt", "Yanıt Akışı", "Son Yanıt", "Son URL", "Durum Özeti", "TLS Güvenliği", "Title", "İçerik Özeti", "Açıklama"]
    all_rows = [headers]
    for row in rows:
        chain = row.get("response_chain")
        codes = [str(code) for code in chain] if isinstance(chain, list) else []
        all_rows.append([
            row.get("source", ""), row.get("url", ""), codes[0] if codes else row.get("code", ""),
            " → ".join(codes) if codes else row.get("code", ""), row.get("code", ""), row.get("final_url", ""),
            row.get("availability", row.get("state", "")), row.get("tls", "—"), row.get("title", ""),
            row.get("summary", ""), row.get("detail", ""),
        ])
    def col(n: int) -> str:
        out = ""
        while n:
            n, rem = divmod(n - 1, 26)
            out = chr(65 + rem) + out
        return out
    sheet_rows = []
    for ri, row in enumerate(all_rows, 1):
        cells = []
        for ci, value in enumerate(row, 1):
            ref = f"{col(ci)}{ri}"
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{_xml_escape(value)}</t></is></c>')
        sheet_rows.append(f'<row r="{ri}">{"".join(cells)}</row>')
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             f'<sheetData>{"".join(sheet_rows)}</sheetData></worksheet>')
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Sonuçlar" sheetId="1" r:id="rId1"/></sheets></workbook>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '</Relationships>')
    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>')
    types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
             '<Default Extension="xml" ContentType="application/xml"/>'
             '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
             '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
             '</Types>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as book:
        book.writestr("[Content_Types].xml", types)
        book.writestr("_rels/.rels", root_rels)
        book.writestr("xl/workbook.xml", workbook)
        book.writestr("xl/_rels/workbook.xml.rels", rels)
        book.writestr("xl/worksheets/sheet1.xml", sheet)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1040x690")
        self.minsize(780, 520)
        self.files: list[Path] = []
        self.results: list[dict[str, str]] = []
        self.row_data: dict[str, dict[str, str]] = {}
        self.filter_source = tk.StringVar(value="Tümü")
        self.filter_code = tk.StringVar(value="Tümü")
        self.filter_state = tk.StringVar(value="Tümü")
        self.filter_tls = tk.StringVar(value="Tümü")
        self.filter_text = tk.StringVar(value="")
        self.events: queue.Queue = queue.Queue()
        self.stop_event = threading.Event()
        self.running = False
        self._build()
        self._load_folder_files()
        self.after(100, self._poll)

    def _build(self):
        root = ttk.Frame(self, padding=14)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text=APP_NAME, font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(root, text="XLSX dosyalarının A sütunundaki adresleri kontrol eder. Sonuçlar yeni bir dosyaya kaydedilir.").pack(anchor="w", pady=(2, 10))
        bar = ttk.Frame(root)
        bar.pack(fill="x", pady=(0, 8))
        self.add_btn = ttk.Button(bar, text="XLSX Ekle…", command=self._add_files)
        self.add_btn.pack(side="left")
        self.scan_btn = ttk.Button(bar, text="Kontrolü Başlat", command=self._start)
        self.scan_btn.pack(side="left", padx=6)
        self.stop_btn = ttk.Button(bar, text="Durdur", command=self.stop_event.set, state="disabled")
        self.stop_btn.pack(side="left")
        self.open_btn = ttk.Button(bar, text="Siteyi Aç", command=self._open_selected)
        self.open_btn.pack(side="left", padx=6)
        self.details_btn = ttk.Button(bar, text="Açıklamayı Gör", command=self._show_details)
        self.details_btn.pack(side="left")
        self.export_btn = ttk.Button(bar, text="Görünenleri XLSX Kaydet…", command=self._save, state="disabled")
        self.export_btn.pack(side="right")
        panes = ttk.Panedwindow(root, orient="horizontal")
        panes.pack(fill="both", expand=True)
        left = ttk.LabelFrame(panes, text="Taranacak dosyalar", padding=8)
        right = ttk.LabelFrame(panes, text="Kontrol sonuçları", padding=8)
        panes.add(left, weight=1); panes.add(right, weight=4)
        self.file_list = tk.Listbox(left, selectmode="extended", height=15, width=36)
        self.file_list.pack(fill="both", expand=True)
        self.file_list.bind("<<ListboxSelect>>", self._update_selected_file_count)
        self.file_count_label = ttk.Label(left, text="Seçili: 0 / 0")
        self.file_count_label.pack(anchor="w", pady=(5, 0))
        ttk.Label(left, text="Bir veya daha fazla dosya seçin (Ctrl/Shift ile çoklu seçim).", wraplength=230).pack(anchor="w", pady=(6, 0))
        file_actions = ttk.Frame(left)
        file_actions.pack(fill="x", pady=(8, 0))
        ttk.Button(file_actions, text="Tümünü Seç", command=self._select_all_files).pack(side="left")
        ttk.Button(file_actions, text="Seçimi Temizle", command=self._clear_file_selection).pack(side="left", padx=(5, 0))
        ttk.Button(left, text="Seçileni Listeden Kaldır", command=self._remove_files).pack(anchor="w", pady=(6, 0))
        filters = ttk.Frame(right)
        filters.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 7))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        ttk.Label(filters, text="Kaynak:").grid(row=0, column=0, sticky="w")
        self.source_filter_box = ttk.Combobox(filters, textvariable=self.filter_source, state="readonly", width=18, values=("Tümü",))
        self.source_filter_box.grid(row=0, column=1, sticky="w", padx=(3, 10))
        self.source_filter_box.bind("<<ComboboxSelected>>", self._apply_filters)
        ttk.Label(filters, text="Response:").grid(row=0, column=2, sticky="w")
        self.code_filter_box = ttk.Combobox(filters, textvariable=self.filter_code, state="readonly", width=11, values=("Tümü",))
        self.code_filter_box.grid(row=0, column=3, sticky="w", padx=(3, 10))
        self.code_filter_box.bind("<<ComboboxSelected>>", self._apply_filters)
        ttk.Label(filters, text="Durum:").grid(row=0, column=4, sticky="w")
        self.state_filter_box = ttk.Combobox(filters, textvariable=self.filter_state, state="readonly", width=15, values=("Tümü",))
        self.state_filter_box.grid(row=0, column=5, sticky="w", padx=(3, 10))
        self.state_filter_box.bind("<<ComboboxSelected>>", self._apply_filters)
        ttk.Label(filters, text="TLS:").grid(row=0, column=6, sticky="w")
        self.tls_filter_box = ttk.Combobox(filters, textvariable=self.filter_tls, state="readonly", width=19, values=("Tümü",))
        self.tls_filter_box.grid(row=0, column=7, sticky="w", padx=(3, 0))
        self.tls_filter_box.bind("<<ComboboxSelected>>", self._apply_filters)
        ttk.Label(filters, text="Ara:").grid(row=1, column=0, sticky="w", pady=(5, 0))
        search = ttk.Entry(filters, textvariable=self.filter_text, width=46)
        search.grid(row=1, column=1, columnspan=5, sticky="ew", padx=(3, 10), pady=(5, 0))
        search.bind("<KeyRelease>", self._apply_filters)
        ttk.Button(filters, text="Filtreleri Temizle", command=self._reset_filters).grid(row=1, column=6, columnspan=2, sticky="e", pady=(5, 0))
        columns = ("source", "url", "code", "state", "tls", "title", "summary", "detail")
        self.table = ttk.Treeview(right, columns=columns, show="headings")
        labels = {"source": "Kaynak", "url": "Site / URL", "code": "Response", "state": "Durum", "tls": "TLS Güvenliği", "title": "Title", "summary": "İçerik Özeti", "detail": "Açıklama"}
        widths = {"source": 120, "url": 240, "code": 75, "state": 125, "tls": 145, "title": 190, "summary": 260, "detail": 340}
        for key in columns:
            self.table.heading(key, text=labels[key])
            self.table.column(key, width=widths[key], minwidth=65, stretch=(key in ("url", "detail")))
        yscroll = ttk.Scrollbar(right, orient="vertical", command=self.table.yview)
        xscroll = ttk.Scrollbar(right, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.table.bind("<Double-1>", self._open_selected)
        self.table.bind("<Return>", self._open_selected)
        self.table.bind("<Shift-MouseWheel>", self._horizontal_wheel)
        self.table.grid(row=1, column=0, sticky="nsew")
        yscroll.grid(row=1, column=1, sticky="ns")
        xscroll.grid(row=2, column=0, sticky="ew")
        bottom = ttk.Frame(root)
        bottom.pack(fill="x", pady=(10, 0))
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill="x")
        self.status = ttk.Label(bottom, text="Hazır")
        self.status.pack(anchor="w", pady=(5, 0))
        ttk.Label(root, text="Yalnızca sol listede seçili XLSX dosyaları taranır. Sonuçlarda yatay kaydırma çubuğunu veya Shift + fare tekerini kullanın. 🟢 Aktif · 🔴 Pasif/ulaşılamıyor · 🟡 Belirsiz. Herhangi bir HTTP yanıtı (404 ve 5xx dahil) sunucunun ulaşıldığını gösterir; sayfanın hatasız çalıştığını garanti etmez. TLS uyarısı ayrı sütundadır. Siteyi açmak için satıra çift tıklayın veya satırı seçip Siteyi Aç'a basın.", wraplength=980).pack(anchor="w", pady=(8, 0))

    def _select_all_files(self):
        if self.files:
            self.file_list.selection_set(0, "end")
        self._update_selected_file_count()

    def _clear_file_selection(self):
        self.file_list.selection_clear(0, "end")
        self._update_selected_file_count()

    def _update_selected_file_count(self, _event=None):
        if hasattr(self, "file_count_label"):
            self.file_count_label.configure(text=f"Seçili: {len(self.file_list.curselection())} / {len(self.files)}")

    def _horizontal_wheel(self, event):
        self.table.xview_scroll(int(-event.delta / 120), "units")
        return "break"

    def _filtered_results(self):
        source = self.filter_source.get()
        code = self.filter_code.get()
        state = self.filter_state.get()
        tls = self.filter_tls.get()
        query = self.filter_text.get().strip().casefold()
        visible = []
        for result in self.results:
            if source != "Tümü" and result.get("source", "") != source:
                continue
            if code != "Tümü" and result.get("code", "") != code:
                continue
            if state != "Tümü" and result.get("state", "") != state:
                continue
            if tls != "Tümü" and result.get("tls", "") != tls:
                continue
            searchable = " ".join(result.get(key, "") for key in ("url", "final_url", "title", "summary", "detail", "source")).casefold()
            if query and query not in searchable:
                continue
            visible.append(result)
        return visible

    def _refresh_filter_choices(self):
        sources = sorted({r.get("source", "") for r in self.results if r.get("source")})
        codes = {r.get("code", "") for r in self.results if r.get("code")}
        codes = sorted(codes, key=lambda item: (not item.isdigit(), int(item) if item.isdigit() else item.casefold()))
        states = sorted({r.get("state", "") for r in self.results if r.get("state")})
        tls_states = sorted({r.get("tls", "") for r in self.results if r.get("tls")})
        self.source_filter_box.configure(values=("Tümü", *sources))
        self.code_filter_box.configure(values=("Tümü", *codes))
        self.state_filter_box.configure(values=("Tümü", *states))
        self.tls_filter_box.configure(values=("Tümü", *tls_states))
        if self.filter_source.get() not in ("Tümü", *sources):
            self.filter_source.set("Tümü")
        if self.filter_code.get() not in ("Tümü", *codes):
            self.filter_code.set("Tümü")
        if self.filter_state.get() not in ("Tümü", *states):
            self.filter_state.set("Tümü")
        if self.filter_tls.get() not in ("Tümü", *tls_states):
            self.filter_tls.set("Tümü")

    def _apply_filters(self, _event=None):
        for item in self.table.get_children():
            self.table.delete(item)
        self.row_data.clear()
        for result in self._filtered_results():
            status_text = {"Aktif": "🟢 Aktif", "Pasif": "🔴 Pasif", "Belirsiz": "🟡 Belirsiz", "Ulaşılamıyor": "🔴 Ulaşılamıyor", "Geçersiz": "⚫ Geçersiz"}.get(result["state"], result["state"])
            tls_text = {"Güvenli HTTPS": "🔒 Güvenli", "Sertifika uyarısı": "⚠ Sertifika uyarısı", "Sertifika farkı": "⚠ İstemci farkı", "Sertifika sorunu": "⚠ Sertifika sorunu", "Şifrelenmemiş HTTP": "⚠ HTTP (şifresiz)"}.get(result.get("tls", "—"), result.get("tls", "—"))
            item = self.table.insert("", "end", values=(result.get("source", ""), result.get("url", ""), result.get("code", ""), status_text, tls_text, result.get("title", ""), result.get("summary", ""), result.get("detail", "")))
            self.row_data[item] = result

    def _reset_filters(self):
        self.filter_source.set("Tümü")
        self.filter_code.set("Tümü")
        self.filter_state.set("Tümü")
        self.filter_tls.set("Tümü")
        self.filter_text.set("")
        self._apply_filters()

    def _open_selected(self, _event=None):
        selected = self.table.selection()
        if not selected:
            messagebox.showinfo(APP_NAME, "Önce sonuçlardan bir satır seçin.")
            return
        # URL is the second displayed column; prefer its resolved redirect target when available.
        result = self.row_data.get(selected[0])
        if not result:
            return
        address = result.get("final_url") or result.get("url", "")
        if address.startswith(("http://", "https://")):
            webbrowser.open_new_tab(address)
        else:
            messagebox.showwarning(APP_NAME, "Bu satırda açılabilir bir HTTP/HTTPS adresi yok.")

    def _show_details(self):
        selected = self.table.selection()
        if not selected:
            messagebox.showinfo(APP_NAME, "Önce sonuçlardan bir satır seçin.")
            return
        result = self.row_data.get(selected[0])
        if not result:
            return
        detail_window = tk.Toplevel(self)
        detail_window.title(f"Kontrol açıklaması — {result.get('url', '')}")
        detail_window.geometry("760x520")
        detail_window.minsize(520, 320)
        content = ttk.Frame(detail_window, padding=12)
        content.pack(fill="both", expand=True)
        text = tk.Text(content, wrap="word", height=20, width=80)
        scroll = ttk.Scrollbar(content, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        fields = (
            ("Kaynak", result.get("source", "")),
            ("Site / URL", result.get("url", "")),
            ("Son URL", result.get("final_url", "")),
            ("Response", result.get("code", "")),
            ("Durum", result.get("state", "")),
            ("TLS Güvenliği", result.get("tls", "")),
            ("Title", result.get("title", "")),
            ("İçerik Özeti", result.get("summary", "")),
            ("Açıklama", result.get("detail", "")),
        )
        text.insert("1.0", "\n\n".join(f"{label}:\n{value or '—'}" for label, value in fields))
        text.configure(state="disabled")

    def _load_folder_files(self):
        for path in sorted(APP_DIR.glob("*.xlsx")):
            if path.name.lower().startswith("domain-kontrol-"):
                continue
            self._add_path(path)

    def _add_path(self, path: Path):
        path = path.resolve()
        if path not in self.files:
            self.files.append(path)
            self.file_list.insert("end", path.name)
            self._update_selected_file_count()

    def _add_files(self):
        chosen = filedialog.askopenfilenames(title="XLSX dosyalarını seçin", filetypes=[("Excel çalışma kitabı", "*.xlsx")])
        for name in chosen:
            self._add_path(Path(name))

    def _remove_files(self):
        for i in reversed(self.file_list.curselection()):
            self.file_list.delete(i)
            del self.files[i]
        self._update_selected_file_count()

    def _start(self):
        if self.running:
            return
        if not self.files:
            messagebox.showinfo(APP_NAME, "Önce en az bir XLSX dosyası ekleyin.")
            return
        selected_indices = self.file_list.curselection()
        if not selected_indices:
            messagebox.showinfo(APP_NAME, "Taranacak XLSX dosyalarını soldaki listeden seçin.")
            return
        selected_files = [self.files[i] for i in selected_indices]
        tasks = []
        errors = []
        for path in selected_files:
            try:
                urls = read_xlsx_column_a(path)
                tasks.extend((path.stem, url) for url in urls)
            except Exception as exc:
                errors.append(f"{path.name}: {exc}")
        if not tasks:
            messagebox.showerror(APP_NAME, "A sütununda kontrol edilebilir adres bulunamadı." + ("\n" + "\n".join(errors) if errors else ""))
            return
        self.results.clear()
        self.row_data.clear()
        for item in self.table.get_children():
            self.table.delete(item)
        self._reset_filters()
        self.stop_event.clear()
        self.running = True
        self.scan_btn.configure(state="disabled"); self.add_btn.configure(state="disabled"); self.stop_btn.configure(state="normal")
        self.export_btn.configure(state="disabled")
        self.progress.configure(maximum=len(tasks), value=0)
        self.status.configure(text=f"0 / {len(tasks)} kontrol ediliyor…")
        if errors:
            messagebox.showwarning(APP_NAME, "Bazı dosyalar okunamadı:\n" + "\n".join(errors))
        threading.Thread(target=self._worker, args=(tasks,), daemon=True).start()

    def _worker(self, tasks):
        completed = 0
        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            futures = {pool.submit(check_url, url): (source, url) for source, url in tasks}
            for future in concurrent.futures.as_completed(futures):
                if self.stop_event.is_set():
                    for pending in futures:
                        pending.cancel()
                    break
                source, original = futures[future]
                try:
                    result = future.result()
                except Exception as exc:
                    result = {"url": original, "final_url": "", "code": "Kod yok", "state": "Belirsiz", "detail": f"HTTP yanıtı alınamadı — {exc}", "title": "", "summary": "", "tls": "Bağlantı sorunu"}
                result["source"] = source
                completed += 1
                self.events.put(("row", result, completed, len(tasks)))
        self.events.put(("done", self.stop_event.is_set(), completed, len(tasks)))

    def _poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "row":
                    _, result, done, total = event
                    self.results.append(result)
                    self._refresh_filter_choices()
                    self._apply_filters()
                    self.progress.configure(value=done)
                    self.status.configure(text=f"{done} / {total} tamamlandı — Aktif: {sum(r['state']=='Aktif' for r in self.results)} | Pasif: {sum(r['state']=='Pasif' for r in self.results)} | Belirsiz: {sum(r['state'] not in ('Aktif','Pasif') for r in self.results)}")
                else:
                    _, stopped, done, total = event
                    self.running = False
                    self.scan_btn.configure(state="normal"); self.add_btn.configure(state="normal"); self.stop_btn.configure(state="disabled")
                    self.export_btn.configure(state="normal" if self.results else "disabled")
                    self.status.configure(text=f"{'Durduruldu' if stopped else 'Tamamlandı'} — {done} / {total} adres işlendi.")
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _save(self):
        if not self.results:
            return
        suggested = f"domain-kontrol-{dt.datetime.now():%Y%m%d-%H%M}.xlsx"
        name = filedialog.asksaveasfilename(title="Sonuçları kaydet", initialdir=str(APP_DIR), initialfile=suggested, defaultextension=".xlsx", filetypes=[("Excel çalışma kitabı", "*.xlsx")])
        if not name:
            return
        try:
            visible_results = self._filtered_results()
            write_results_xlsx(Path(name), visible_results)
            self.status.configure(text=f"Kaydedildi: {name}")
            messagebox.showinfo(APP_NAME, f"Filtreye uyan {len(visible_results)} sonuç XLSX olarak kaydedildi.")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Dosya kaydedilemedi:\n{exc}")


if __name__ == "__main__":
    if not TK_AVAILABLE:
        raise SystemExit("Masaüstü arayüzü için Tkinter kurulu Python gerekir. Web arayüzü için python run.py kullanın.")
    App().mainloop()
