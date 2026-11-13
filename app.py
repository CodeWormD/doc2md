# -*- coding: utf-8 -*-
"""DOC2MD — локальный веб-сервис конвертации документов в Markdown."""
import os
import secrets
import socket
import sys
import tempfile
import time

from flask import Flask, jsonify, render_template, request

from converter import ConvertError, convert_file

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 МБ

ALLOWED = {".xlsx", ".xlsm", ".xls", ".csv", ".docx", ".doc", ".pdf"}

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

TMP_DIR = tempfile.mkdtemp(prefix="doc2md_")

# результаты последней конвертации — для скачивания кнопкой «Сохранить как»
RESULTS: dict[str, dict] = {}
RESULT_TTL = 3600  # секунд
RESULT_MAX = 50


def _store_result(markdown: str, download_name: str) -> str:
    now = time.time()
    for k in [k for k, v in RESULTS.items() if now - v["ts"] > RESULT_TTL]:
        RESULTS.pop(k, None)
    rid = secrets.token_urlsafe(8)
    RESULTS[rid] = {"markdown": markdown, "download_name": download_name, "ts": now}
    while len(RESULTS) > RESULT_MAX:
        RESULTS.pop(next(iter(RESULTS)))
    return rid


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/convert")
def convert():
    f = request.files.get("file")
    if f is None or not f.filename:
        return jsonify(error="Файл не передан."), 400

    name = os.path.basename(f.filename.replace("\\", "/"))
    ext = os.path.splitext(name)[1].lower()
    if ext not in ALLOWED:
        return jsonify(error=(
            f"Формат «{ext or '???'}» не поддерживается. "
            "Доступны: .xlsx .xlsm .xls .csv .docx .doc .pdf"
        )), 415

    strip_headers = request.form.get("strip_headers", "true") == "true"
    page_markers = request.form.get("page_markers", "false") == "true"

    fd, tmp_path = tempfile.mkstemp(suffix=ext, dir=TMP_DIR)
    os.close(fd)
    try:
        f.save(tmp_path)
        started = time.perf_counter()
        md, warnings = convert_file(tmp_path, ext, strip_headers=strip_headers,
                                    page_markers=page_markers, tmp_dir=TMP_DIR,
                                    title=os.path.splitext(name)[0])
        elapsed = time.perf_counter() - started
        stats = {
            "chars": len(md),
            "words": len(md.split()),
            "lines": md.count("\n") + 1,
            "tokens_est": round(len(md) / 3.5),  # грубая оценка для LLM
            "seconds": round(elapsed, 2),
        }
        return jsonify(
            markdown=md,
            warnings=warnings,
            stats=stats,
            download_name=os.path.splitext(name)[0] + ".md",
            source_name=name,
            download_url=f"/download/{_store_result(md, os.path.splitext(name)[0] + '.md')}",
        )
    except ConvertError as exc:
        return jsonify(error=str(exc)), 422
    except Exception as exc:  # noqa: BLE001
        app.logger.exception("conversion failed")
        return jsonify(error=f"Внутренняя ошибка: {exc}"), 500
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


@app.get("/download/<rid>")
def download(rid: str):
    item = RESULTS.get(rid)
    if item is None:
        return jsonify(error="Результат не найден или устарел — сконвертируйте файл заново."), 404
    resp = app.response_class(item["markdown"].encode("utf-8"), mimetype="text/markdown")
    # RFC 6266: ASCII-fallback + filename* с процентным кодированием (кириллица и т.п.)
    from urllib.parse import quote
    safe = "".join(ch if ord(ch) < 128 and ch not in '"\\' else "_" for ch in item["download_name"]) or "doc2md.md"
    resp.headers.set(
        "Content-Disposition",
        "attachment",
        **{"filename": safe, "filename*": f"UTF-8''{quote(item['download_name'])}"},
    )
    return resp


@app.errorhandler(413)
def too_large(_):
    return jsonify(error="Файл больше 200 МБ."), 413


def _free_port(start: int) -> int:
    for port in range(start, start + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return start


def main() -> None:
    port = _free_port(int(os.environ.get("PORT", "8765")))
    url = f"http://127.0.0.1:{port}"
    print(f"\n  DOC2MD работает: {url}")
    print("  Остановить: Ctrl+C\n")
    if os.environ.get("NO_BROWSER") != "1":
        import threading
        import webbrowser

        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)


if __name__ == "__main__":
    main()
