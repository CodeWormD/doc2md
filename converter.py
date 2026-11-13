# -*- coding: utf-8 -*-
"""
Конвертация офисных документов в «плоский» Markdown, оптимизированный
для чтения языковыми моделями: минимум мусора, максимум смысла.

Поддержка: .xlsx/.xlsm, .xls, .csv, .docx, .doc (через MS Word), .pdf.
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import os
import re


class ConvertError(Exception):
    """Ошибка конвертации с понятным пользователю сообщением."""


# ---------------------------------------------------------------------------
# Общие утилиты
# ---------------------------------------------------------------------------

def tidy(md: str) -> str:
    """Компактный вид: без пробелов в конце строк и слипшихся пустых строк."""
    md = md.replace("\r\n", "\n").replace("\r", "\n")
    md = re.sub(r"[ \t]+\n", "\n", md)
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md.strip() + "\n"


def _esc_cell(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ").strip()


def _fmt_value(v) -> str:
    """Значение ячейки -> компактная строка (даты ISO, числа без хвостового .0)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, dt.datetime):
        if v.hour or v.minute or v.second:
            return v.strftime("%Y-%m-%d %H:%M:%S")
        return v.strftime("%Y-%m-%d")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, dt.time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, float):
        if v.is_integer() and abs(v) < 1e15:
            return str(int(v))
        return str(v)
    return str(v).strip()


def _grid_to_md(rows) -> str:
    """Сетка значений -> markdown-таблица. Пустые строки/колонки обрезаются."""
    rows = [[("" if c is None else str(c)) for c in row] for row in rows]
    width = max((len(r) for r in rows), default=0)
    if width == 0:
        return ""
    rows = [r + [""] * (width - len(r)) for r in rows]

    last = -1  # индекс последней непустой колонки
    for r in rows:
        for i in range(width - 1, -1, -1):
            if r[i].strip():
                last = i
                break
        if last == width - 1:
            break
    if last < 0:
        return ""
    rows = [r[: last + 1] for r in rows]
    # пустые строки убираем ДО выбора «шапки»
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        return ""

    head = rows[0]
    out = [
        "| " + " | ".join(_esc_cell(c) for c in head) + " |",
        "|" + "|".join([" --- "] * len(head)) + "|",
    ]
    out.extend("| " + " | ".join(_esc_cell(c) for c in r) + " |" for r in rows[1:])
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def _xlsx(path: str, title: str | None = None):
    from openpyxl import load_workbook

    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception:
        wb = load_workbook(path, data_only=True)

    parts, warnings = [], []
    filled = 0
    try:
        for ws in wb.worksheets:
            rows = ([_fmt_value(v) for v in row] for row in ws.iter_rows(values_only=True))
            table = _grid_to_md(rows)
            if table:
                filled += 1
                parts.append(f"## Лист: {ws.title}\n\n{table}")
            else:
                warnings.append(f"Лист «{ws.title}» пуст — пропущен.")
    finally:
        wb.close()
    if not parts:
        raise ConvertError("В книге Excel не найдено данных (все листы пусты).")

    title = title or os.path.splitext(os.path.basename(path))[0]
    head = f"# {title}\n\n" if filled > 1 else ""
    return head + "\n\n".join(parts) + "\n", warnings


def _xls(path: str, title: str | None = None):
    import xlrd

    book = xlrd.open_workbook(path)
    parts, warnings = [], []
    filled = 0
    for ws in book.sheets():
        rows = []
        for r in range(ws.nrows):
            row = []
            for c in range(ws.ncols):
                cell = ws.cell(r, c)
                v = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    v = xlrd.xldate.xldate_as_datetime(v, book.datemode)
                row.append(_fmt_value(v))
            rows.append(row)
        table = _grid_to_md(rows)
        if table:
            filled += 1
            parts.append(f"## Лист: {ws.name}\n\n{table}")
        else:
            warnings.append(f"Лист «{ws.name}» пуст — пропущен.")
    if not parts:
        raise ConvertError("В книге Excel не найдено данных (все листы пусты).")

    title = title or os.path.splitext(os.path.basename(path))[0]
    head = f"# {title}\n\n" if filled > 1 else ""
    return head + "\n\n".join(parts) + "\n", warnings


# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------

_CSV_DELIMS = ",;\t|"


def _csv(path: str):
    raw = open(path, "rb").read()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        text = raw.decode("utf-16")
    else:
        text = None
        for enc in ("utf-8-sig", "cp1251"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            text = raw.decode("utf-8", errors="replace")

    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=_CSV_DELIMS)
    except csv.Error:
        dialect = csv.excel
    rows = list(csv.reader(io.StringIO(text), dialect))
    table = _grid_to_md(rows)
    if not table:
        raise ConvertError("CSV-файл пуст.")
    return table + "\n", []


# ---------------------------------------------------------------------------
# Word
# ---------------------------------------------------------------------------

def _fix_md_tables(md: str) -> str:
    """Если «шапка» markdown-таблицы полностью пуста — делаем заголовком
    первую содержательную строку (mammoth иногда добавляет пустую строку)."""
    blocks = md.split("\n\n")
    fixed = []
    for block in blocks:
        lines = [ln for ln in block.split("\n") if ln.strip()]
        if len(lines) >= 3 and all(ln.lstrip().startswith("|") for ln in lines):
            cells = [c.strip() for c in lines[0].strip().strip("|").split("|")]
            if cells and not any(cells):
                lines = [lines[2], lines[1]] + lines[3:]
        fixed.append("\n".join(lines))
    return "\n\n".join(fixed)


def _docx(path: str):
    import mammoth
    from markdownify import markdownify

    with open(path, "rb") as fh:
        result = mammoth.convert_to_html(fh)
    md = _fix_md_tables(tidy(markdownify(result.value, heading_style="ATX", bullets="-", strip=["img"])))
    if not md.strip():
        raise ConvertError("Документ не содержит текста (только картинки?).")
    return md, []


def _doc(path: str, tmp_dir: str):
    docx_path = _doc_to_docx(path, tmp_dir)
    md, warnings = _docx(docx_path)
    warnings.insert(0, "Формат .doc: промежуточная конвертация выполнена через Microsoft Word (COM).")
    try:
        os.remove(docx_path)
    except OSError:
        pass
    return md, warnings


def _doc_to_docx(path: str, tmp_dir: str) -> str:
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise ConvertError(
            "Для файлов .doc нужен установленный Microsoft Word (пакет pywin32). "
            "Пересохраните документ в .docx либо установите Word."
        )
    src = os.path.abspath(path)
    dst = os.path.join(tmp_dir, f"doc_{os.getpid()}_{os.path.basename(src)}.docx")
    word = None
    try:
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
        doc.SaveAs2(dst, FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
        doc.Close(SaveChanges=False)
    except Exception as exc:  # noqa: BLE001
        raise ConvertError(
            f"Не удалось открыть .doc через Microsoft Word: {exc}. "
            "Пересохраните файл в .docx и повторите."
        )
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:  # noqa: BLE001
                pass
        pythoncom.CoUninitialize()
    if not os.path.exists(dst):
        raise ConvertError("Microsoft Word не сохранил промежуточный .docx.")
    return dst


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

# «Хвостовые» номера страниц: 3, - 5 -, стр. 7, Page 12 of 40, [200] и т.п.
_PAGENO_RE = re.compile(
    r"^\s*[-–—·•\[\(|]*(?:стр(?:\.|аница)?|page|p\.?)?\s*\d{1,4}\s*"
    r"(?:(?:/|из|of)\s*\d{1,4})?\s*[-–—·•\]\)|]*\.?\s*$",
    re.IGNORECASE,
)


def _is_pageno(line: str) -> bool:
    return bool(_PAGENO_RE.match(line)) and any(ch.isdigit() for ch in line)


def _overlap(ax0, ay0, ax1, ay1, rect) -> float:
    """Доля пересечения прямоугольника (ax..) с rect (bbox таблицы)."""
    ix = max(0.0, min(ax1, rect[2]) - max(ax0, rect[0]))
    iy = max(0.0, min(ay1, rect[3]) - max(ay0, rect[1]))
    area = max((ax1 - ax0) * (ay1 - ay0), 1e-6)
    return (ix * iy) / area


def _pdf(path: str, strip_headers: bool = True, page_markers: bool = False):
    try:
        import pymupdf as fitz
    except ImportError:  # старые версии пакета
        import fitz  # type: ignore

    doc = fitz.open(path)
    page_segments, page_lines = [], []
    with doc:
        for page in doc:
            segments, table_rects = [], []
            try:
                finder = page.find_tables()
                for t in finder.tables:
                    data = t.extract()
                    if data:
                        table_rects.append(tuple(t.bbox))
                        segments.append((t.bbox[1], t.bbox[0], "table", data))
            except Exception:  # noqa: BLE001 — страницы без таблиц/нестандартные
                pass

            for b in page.get_text("blocks"):
                if len(b) < 7 or b[6] != 0:  # только текстовые блоки
                    continue
                text = (b[4] or "").strip()
                if not text:
                    continue
                x0, y0, x1, y1 = b[:4]
                # блок внутри таблицы уже попал в markdown-таблицу
                if any(_overlap(x0, y0, x1, y1, tr) > 0.5 for tr in table_rects):
                    continue
                segments.append((y0, x0, "text", text))

            segments.sort(key=lambda s: (round(s[0], 1), s[1]))
            page_segments.append(segments)
            lines = []
            for s in segments:
                if s[2] == "text":
                    lines.extend(ln.strip() for ln in s[3].splitlines() if ln.strip())
            page_lines.append(lines)

    # Колонтитулы: короткие строки, повторяющиеся на большинстве страниц
    junk = set()
    if strip_headers and len(page_lines) >= 3:
        import math
        from collections import Counter

        pages_with = Counter()
        for lines in page_lines:
            pages_with.update(set(lines))
        threshold = math.ceil(0.7 * len(page_lines))
        junk = {ln for ln, cnt in pages_with.items() if cnt >= threshold and len(ln) <= 150}

    parts, warnings = [], []
    empty = []
    for i, segments in enumerate(page_segments, 1):
        chunks = []
        for _, _, kind, payload in segments:
            if kind == "table":
                rows = [[("" if c is None else str(c)) for c in row] for row in payload]
                tmd = _grid_to_md(rows)
                if tmd:
                    chunks.append(tmd)
            else:
                lines = payload.splitlines()
                if junk:
                    lines = [ln for ln in lines if ln.strip() not in junk]
                if strip_headers:
                    lines = [ln for ln in lines if not _is_pageno(ln.strip())]
                text = "\n".join(lines).strip()
                if text:
                    chunks.append(text)
        body = "\n\n".join(chunks).strip()
        if body:
            prefix = f"<!-- стр. {i} -->\n" if page_markers else ""
            parts.append(prefix + body)
        else:
            empty.append(i)

    if not parts:
        raise ConvertError(
            "В PDF не найден текстовый слой — вероятно, это скан или набор картинок. "
            "Нужен OCR (например, Tesseract), чтобы извлечь текст."
        )
    if empty:
        shown = ", ".join(map(str, empty[:10]))
        more = f" …(всего {len(empty)})" if len(empty) > 10 else ""
        warnings.append(f"Страницы без текста: {shown}{more} — возможно, сканы или картинки.")
    return "\n\n".join(parts) + "\n", warnings


# ---------------------------------------------------------------------------
# Точка входа
# ---------------------------------------------------------------------------

def convert_file(path: str, ext: str, *, strip_headers: bool = True,
                 page_markers: bool = False, tmp_dir: str | None = None,
                 title: str | None = None):
    """Возвращает (markdown, warnings). Бросает ConvertError с понятным текстом.

    title — исходное имя файла (без расширения) для заголовка в Excel-книгах."""
    ext = ext.lower()
    if ext in (".xlsx", ".xlsm"):
        return _xlsx(path, title)
    if ext == ".xls":
        return _xls(path, title)
    if ext == ".csv":
        return _csv(path)
    if ext == ".docx":
        return _docx(path)
    if ext == ".doc":
        return _doc(path, tmp_dir or os.path.dirname(os.path.abspath(path)))
    if ext == ".pdf":
        return _pdf(path, strip_headers=strip_headers, page_markers=page_markers)
    raise ConvertError(
        f"Формат «{ext}» не поддерживается. Доступны: .xlsx .xlsm .xls .csv .docx .doc .pdf"
    )
