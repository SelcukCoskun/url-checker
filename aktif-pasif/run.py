#!/usr/bin/env python3
"""Local-only URL checker web app. No cloud services or external UI assets."""

from __future__ import annotations

import argparse
import concurrent.futures
import csv
import ipaddress
import io
import json
import re
import secrets
import socket
import threading
import time
import urllib.parse
import webbrowser
import zipfile
import xml.etree.ElementTree as ET
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from domain_checker import check_url, read_xlsx_column_a, write_results_xlsx

APP_DIR = Path(__file__).resolve().parent
WEB_DIR = APP_DIR / "web"
MAX_UPLOAD_BYTES = 10 * 1024 * 1024
MAX_UNPACKED_BYTES = 32 * 1024 * 1024
MAX_ARCHIVE_ENTRIES = 3000
MAX_URLS = 100
SCAN_WORKERS = 3
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".txt", ".ods"}
HEADER_VALUES = {"url", "urls", "domain", "domains", "site", "website", "link", "address", "adres", "web address", "asset", "assets", "host", "hostname", "subdomain", "fqdn"}
ODS_NS = {
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}
HTTP_URL_RE = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)


def _decode_text(raw: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-16", "cp1254", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _urls_from_text_cell(value: str) -> list[str]:
    value = value.strip()
    if not value or value.casefold() in HEADER_VALUES:
        return []
    found = HTTP_URL_RE.findall(value)
    if found:
        return [url.rstrip(".,;:!?)\"'}]") for url in found]
    # TXT files often contain bare domains, one per line. Ignore prose and headings.
    candidate = re.sub(r"^\s*(?:[-*•]|\d+[.)])\s*", "", value).strip(" \t,;\"'<>[]()")
    return [candidate] if candidate and not any(char.isspace() for char in candidate) else []


def _read_delimited_column_a(raw: bytes, extension: str) -> list[str]:
    text = _decode_text(raw)
    if extension == ".txt":
        values = []
        for line in text.splitlines():
            if line.lstrip().startswith("#"):
                continue
            values.extend(_urls_from_text_cell(line))
        return values

    sample = text[:8192]
    if extension == ".tsv":
        delimiter = "\t"
    else:
        try:
            delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
        except csv.Error:
            delimiter = ";" if sample.count(";") > sample.count(",") else ","
    values = []
    for row in csv.reader(io.StringIO(text), delimiter=delimiter):
        if row:
            values.extend(_urls_from_text_cell(row[0]))
    return values


def _read_xls_column_a(raw: bytes) -> list[str]:
    try:
        import xlrd
    except ImportError as exc:
        raise ValueError(".xls desteği için Kurulum.bat dosyasını çalıştırıp bağımlılıkları kurun.") from exc
    try:
        book = xlrd.open_workbook(file_contents=raw, on_demand=True)
    except Exception as exc:
        raise ValueError(f"XLS dosyası açılamadı: {exc}") from exc
    try:
        if not book.nsheets:
            return []
        sheet = book.sheet_by_index(0)
        values = [str(sheet.cell_value(row, 0)).strip() for row in range(sheet.nrows)]
        if values and values[0].casefold() in HEADER_VALUES:
            values = values[1:]
        return [value for value in values if value]
    finally:
        book.release_resources()


def _read_ods_column_a(raw: bytes) -> list[str]:
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as book:
            root = ET.fromstring(book.read("content.xml"))
    except Exception as exc:
        raise ValueError(f"ODS dosyası açılamadı: {exc}") from exc
    sheet = root.find(".//table:table", ODS_NS)
    if sheet is None:
        return []
    values = []
    rows = sheet.findall("table:table-row", ODS_NS)
    for row in rows[:10000]:
        cells = row.findall("table:table-cell", ODS_NS)
        if not cells:
            cells = row.findall("table:covered-table-cell", ODS_NS)
        if cells:
            value = " ".join("".join(node.itertext()).strip() for node in cells[0].findall("text:p", ODS_NS)).strip()
            if value:
                values.append(value)
    if values and values[0].casefold() in HEADER_VALUES:
        values = values[1:]
    return values


def read_uploaded_addresses(filename: str, raw: bytes) -> list[str]:
    extension = Path(filename).suffix.casefold()
    if extension in {".xlsx", ".xlsm"}:
        with zipfile.ZipFile(io.BytesIO(raw)) as book:
            entries = book.infolist()
            if len(entries) > MAX_ARCHIVE_ENTRIES:
                raise ValueError("Çalışma kitabında çok fazla dosya girdisi var.")
            if sum(item.file_size for item in entries) > MAX_UNPACKED_BYTES:
                raise ValueError("Açılmış çalışma kitabı izin verilen boyutu aşıyor.")
            if any(item.flag_bits & 0x1 for item in entries):
                raise ValueError("Şifreli Excel dosyaları desteklenmiyor.")
        return read_xlsx_column_a(io.BytesIO(raw))
    if extension == ".ods":
        with zipfile.ZipFile(io.BytesIO(raw)) as book:
            entries = book.infolist()
            if len(entries) > MAX_ARCHIVE_ENTRIES:
                raise ValueError("Çalışma kitabında çok fazla dosya girdisi var.")
            if sum(item.file_size for item in entries) > MAX_UNPACKED_BYTES:
                raise ValueError("Açılmış çalışma kitabı izin verilen boyutu aşıyor.")
            if any(item.flag_bits & 0x1 for item in entries):
                raise ValueError("Şifreli ODS dosyaları desteklenmiyor.")
        return _read_ods_column_a(raw)
    if extension == ".xls":
        return _read_xls_column_a(raw)
    if extension in {".csv", ".tsv", ".txt"}:
        return _read_delimited_column_a(raw, extension)
    raise ValueError("Desteklenen biçimler: XLSX, XLSM, XLS, ODS, CSV, TSV ve TXT.")

SECURITY_HEADERS = {
    "X-Content-Type-Options": "nosniff",
    "Referrer-Policy": "no-referrer",
    "X-Frame-Options": "DENY",
    "Cross-Origin-Resource-Policy": "same-origin",
    "Permissions-Policy": "camera=(), microphone=(), geolocation=()",
    "Content-Security-Policy": (
        "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; "
        "connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'"
    ),
}


class LocalState:
    def __init__(self):
        self.lock = threading.RLock()
        self.datasets: dict[str, dict] = {}
        self.jobs: dict[str, dict] = {}
        self.active_job_id: str | None = None


STATE = LocalState()


def validate_public_web_target(url: str) -> None:
    """Reject private targets and unapproved ports before each request/redirect."""
    parts = urllib.parse.urlsplit(url)
    scheme = parts.scheme.lower()
    if scheme not in {"http", "https"} or not parts.hostname:
        raise ValueError("Yalnızca HTTP/HTTPS adresleri kontrol edilir.")
    host = parts.hostname.rstrip(".").casefold()
    allowed_ports = {443} if scheme == "https" else {80, 82}
    default_port = 443 if scheme == "https" else 80
    try:
        port = parts.port or default_port
    except ValueError as exc:
        raise ValueError("Adresin port bilgisi geçersiz.") from exc
    if port not in allowed_ports:
        raise ValueError("Yalnızca herkese açık web adreslerinde HTTP 80/82 ve HTTPS 443 portlarına izin verilir.")
    if host in {"localhost", "localhost.localdomain"} or host.endswith((".localhost", ".local", ".internal", ".test")):
        raise ValueError("Yerel veya kurum içi adrese istek gönderilmez.")

    try:
        literal_ip = ipaddress.ip_address(host)
    except ValueError:
        if "." not in host:
            raise ValueError("Tek parçalı kurum içi bilgisayar adları kontrol edilmez.")
        try:
            answers = socket.getaddrinfo(host, port, type=socket.SOCK_STREAM)
        except OSError as exc:
            raise ValueError("Alan adı genel DNS üzerinden çözümlenemedi.") from exc
        addresses = {entry[4][0].split("%", 1)[0] for entry in answers}
        if not addresses:
            raise ValueError("Alan adı için IP adresi bulunamadı.")
        if any(not ipaddress.ip_address(address).is_global for address in addresses):
            raise ValueError("Alan adı özel, yerel veya ayrılmış bir IP adresine çözülüyor; istek engellendi.")
    else:
        if not literal_ip.is_global:
            raise ValueError("Özel, yerel veya ayrılmış IP adreslerine istek gönderilmez.")


def _scan_one(index: int, raw_url: str, source: str) -> tuple[int, dict[str, object]]:
    result = check_url(raw_url, url_guard=validate_public_web_target)
    result["source"] = source
    result["position"] = str(index)
    return index, result


def _run_scan(job_id: str, urls: list[str], source: str) -> None:
    with STATE.lock:
        job = STATE.jobs.get(job_id)
        if job is None:
            return
        job["status"] = "running"

    try:
        with concurrent.futures.ThreadPoolExecutor(
            max_workers=min(SCAN_WORKERS, max(1, len(urls))),
            thread_name_prefix="url-check",
        ) as pool:
            futures = {pool.submit(_scan_one, index, url, source): index for index, url in enumerate(urls)}
            for future in concurrent.futures.as_completed(futures):
                index = futures[future]
                try:
                    _, result = future.result()
                except Exception as exc:
                    result = {
                        "url": urls[index], "final_url": "", "code": "Kod yok", "state": "Belirsiz",
                        "availability": "Belirsiz sonuç", "response_chain": [], "redirect_chain": [],
                        "detail": f"Kontrol tamamlanamadı: {exc}", "title": "", "summary": "",
                        "tls": "Bağlantı sorunu", "source": source, "position": str(index),
                    }
                with STATE.lock:
                    current = STATE.jobs.get(job_id)
                    if current is not None:
                        current["results"].append(result)
                        current["results"].sort(key=lambda row: int(row.get("position", "0")))
                        current["completed"] += 1
        with STATE.lock:
            if job_id in STATE.jobs:
                STATE.jobs[job_id]["status"] = "completed"
            if STATE.active_job_id == job_id:
                STATE.active_job_id = None
    except Exception as exc:
        with STATE.lock:
            if job_id in STATE.jobs:
                STATE.jobs[job_id]["status"] = "error"
                STATE.jobs[job_id]["error"] = str(exc)
            if STATE.active_job_id == job_id:
                STATE.active_job_id = None


def make_handler(app_token: str):
    class LocalHandler(BaseHTTPRequestHandler):
        server_version = "LocalDomainChecker/1.0"

        def log_message(self, _format, *_args):
            # Do not put uploaded file names, URLs, or request data in the terminal log.
            return

        def _is_local_host(self) -> bool:
            host = self.headers.get("Host", "").split(":", 1)[0].casefold()
            return host in {"127.0.0.1", "localhost"}

        def _authorized(self) -> bool:
            return secrets.compare_digest(self.headers.get("X-App-Token", ""), app_token)

        def _send(self, body: bytes, status: int = 200, content_type: str = "application/json; charset=utf-8", extra_headers=None):
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            for name, value in SECURITY_HEADERS.items():
                self.send_header(name, value)
            for name, value in (extra_headers or {}).items():
                self.send_header(name, value)
            self.end_headers()
            if body:
                self.wfile.write(body)

        def _json(self, payload, status: int = 200):
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self._send(body, status)

        def _error(self, message: str, status: int = 400):
            self._json({"error": message}, status)

        def _read_body(self, limit: int) -> bytes:
            if self.headers.get("Transfer-Encoding"):
                raise ValueError("Parçalı aktarım kabul edilmiyor.")
            try:
                length = int(self.headers.get("Content-Length", "0"))
            except ValueError as exc:
                raise ValueError("İstek boyutu okunamadı.") from exc
            if length <= 0:
                raise ValueError("İstek içeriği boş.")
            if length > limit:
                raise OverflowError("İstek boyutu izin verilen sınırı aşıyor.")
            body = self.rfile.read(length)
            if len(body) != length:
                raise ValueError("Dosya aktarımı tamamlanmadı.")
            return body

        def _require_local_and_token(self) -> bool:
            if not self._is_local_host():
                self._error("Bu yerel uygulamaya yalnızca bu bilgisayardan erişilebilir.", 403)
                return False
            if not self._authorized():
                self._error("Oturum anahtarı geçersiz. Sayfayı yenileyin.", 403)
                return False
            return True

        def do_GET(self):
            if not self._is_local_host():
                self._error("Bu yerel uygulamaya yalnızca bu bilgisayardan erişilebilir.", 403)
                return
            path = urllib.parse.urlsplit(self.path).path
            if path == "/":
                try:
                    page = (WEB_DIR / "index.html").read_text(encoding="utf-8").replace("__APP_TOKEN__", app_token)
                except OSError:
                    self._error("Arayüz dosyaları bulunamadı.", 500)
                    return
                self._send(page.encode("utf-8"), content_type="text/html; charset=utf-8")
                return
            if path in {"/app.js", "/styles.css"}:
                name = path.lstrip("/")
                kind = "text/javascript; charset=utf-8" if name.endswith(".js") else "text/css; charset=utf-8"
                try:
                    self._send((WEB_DIR / name).read_bytes(), content_type=kind)
                except OSError:
                    self._error("Arayüz dosyası bulunamadı.", 404)
                return
            if not self._require_local_and_token():
                return
            parts = path.strip("/").split("/")
            if len(parts) == 3 and parts[:2] == ["api", "jobs"]:
                self._get_job(parts[2])
                return
            if len(parts) == 4 and parts[:2] == ["api", "jobs"] and parts[3] == "export":
                self._export_job(parts[2])
                return
            self._error("Adres bulunamadı.", 404)

        def _get_job(self, job_id: str):
            with STATE.lock:
                job = STATE.jobs.get(job_id)
                if job is None:
                    self._error("Tarama kaydı bulunamadı.", 404)
                    return
                snapshot = {
                    "id": job["id"], "status": job["status"], "total": job["total"],
                    "completed": job["completed"], "results": [dict(row) for row in job["results"]],
                    "error": job.get("error", ""),
                }
            self._json(snapshot)

        def _export_job(self, job_id: str):
            with STATE.lock:
                job = STATE.jobs.get(job_id)
                if job is None:
                    self._error("Tarama kaydı bulunamadı.", 404)
                    return
                rows = [dict(row) for row in job["results"]]
            if not rows:
                self._error("Dışa aktarılacak sonuç yok.", 409)
                return
            output = io.BytesIO()
            try:
                write_results_xlsx(output, rows)
            except Exception as exc:
                self._error(f"XLSX raporu oluşturulamadı: {exc}", 500)
                return
            self._send(
                output.getvalue(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                extra_headers={"Content-Disposition": 'attachment; filename="domain-kontrol-sonuclari.xlsx"'},
            )

        def do_POST(self):
            if not self._require_local_and_token():
                return
            path = urllib.parse.urlsplit(self.path).path
            if path == "/api/upload":
                self._upload_file()
            elif path == "/api/scan":
                self._start_scan()
            elif path == "/api/reset":
                self._reset_data()
            else:
                self._error("Adres bulunamadı.", 404)

        def _reset_data(self):
            with STATE.lock:
                if STATE.active_job_id is not None:
                    active = STATE.jobs.get(STATE.active_job_id)
                    if active and active["status"] in {"queued", "running"}:
                        self._error("Tarama tamamlanmadan dosya verileri temizlenemez.", 409)
                        return
                STATE.datasets.clear()
                STATE.jobs.clear()
                STATE.active_job_id = None
            self._json({"cleared": True})

        def _upload_file(self):
            raw_name = urllib.parse.unquote(self.headers.get("X-File-Name", ""))
            filename = Path(raw_name.replace("\\", "/")).name
            extension = Path(filename).suffix.casefold()
            if not filename or extension not in SUPPORTED_EXTENSIONS:
                self._error("Desteklenen biçimler: XLSX, XLSM, XLS, ODS, CSV, TSV ve TXT.", 415)
                return
            try:
                raw = self._read_body(MAX_UPLOAD_BYTES)
                urls = read_uploaded_addresses(filename, raw)
            except OverflowError as exc:
                self._error(str(exc), 413)
                return
            except Exception as exc:
                self._error(f"Dosya okunamadı: {exc}", 400)
                return
            if not urls:
                self._error("Dosyanın ilk sütununda veya TXT satırlarında adres bulunamadı.", 422)
                return
            if len(urls) > MAX_URLS:
                self._error(f"Tek dosyada en fazla {MAX_URLS} URL kontrol edilebilir.", 413)
                return
            dataset_id = secrets.token_urlsafe(18)
            with STATE.lock:
                STATE.datasets.clear()
                STATE.datasets[dataset_id] = {"filename": filename, "urls": urls, "created": time.monotonic()}
            self._json({"dataset_id": dataset_id, "filename": filename, "count": len(urls), "preview": urls[:8]})

        def _start_scan(self):
            try:
                payload = json.loads(self._read_body(4096).decode("utf-8"))
            except (ValueError, UnicodeDecodeError, json.JSONDecodeError) as exc:
                self._error(f"Tarama isteği okunamadı: {exc}", 400)
                return
            dataset_id = payload.get("dataset_id") if isinstance(payload, dict) else None
            if not isinstance(dataset_id, str):
                self._error("Dosya oturumu geçersiz. Dosyayı yeniden seçin.", 409)
                return
            with STATE.lock:
                dataset = STATE.datasets.get(dataset_id)
                if dataset is None:
                    self._error("Önce bir dosya yükleyip URL listesini hazırlayın.", 409)
                    return
                if STATE.active_job_id is not None:
                    active = STATE.jobs.get(STATE.active_job_id)
                    if active and active["status"] in {"queued", "running"}:
                        self._error("Bir tarama halen çalışıyor.", 409)
                        return
                job_id = secrets.token_urlsafe(18)
                urls = list(dataset["urls"])
                filename = dataset["filename"]
                STATE.jobs[job_id] = {
                    "id": job_id, "status": "queued", "total": len(urls), "completed": 0,
                    "results": [], "created": time.monotonic(),
                }
                STATE.active_job_id = job_id
            threading.Thread(target=_run_scan, args=(job_id, urls, filename), daemon=True, name="local-url-scan").start()
            self._json({"job_id": job_id, "total": len(urls)})

    return LocalHandler


class LocalOnlyHTTPServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = False


def main() -> None:
    parser = argparse.ArgumentParser(description="Bu bilgisayarda çalışan yerel URL kontrol arayüzü")
    parser.add_argument("--port", type=int, default=0, help="Yerel port (varsayılan: otomatik seçilir)")
    parser.add_argument("--no-browser", action="store_true", help="Tarayıcıyı otomatik açma")
    args = parser.parse_args()
    if not 0 <= args.port <= 65535:
        parser.error("Port 0 ile 65535 arasında olmalıdır.")

    app_token = secrets.token_urlsafe(32)
    server = LocalOnlyHTTPServer(("127.0.0.1", args.port), make_handler(app_token))
    address = f"http://127.0.0.1:{server.server_port}/"
    print("Yerel URL kontrol arayüzü hazır.")
    print(f"Adres: {address}")
    print("Sunucu yalnızca bu bilgisayara açıktır. Durdurmak için Ctrl+C kullanın.")
    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open_new_tab(address)).start()
    try:
        server.serve_forever(poll_interval=0.5)
    except KeyboardInterrupt:
        print("\nYerel sunucu durduruluyor.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
