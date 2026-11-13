# -*- coding: utf-8 -*-
"""E2E-проверка интерфейса DOC2MD в headless Chromium (dev-скрипт, нужен `pip install playwright`).

Проверяет: drag&drop на зону, drag&drop мимо зоны, выбор файла кликом,
сообщение об ошибке для неподдерживаемого формата, визуальный отклик при переносе.
"""
import io
import pathlib
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from playwright.sync_api import sync_playwright

BASE = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8766"
ROOT = pathlib.Path(__file__).resolve().parents[1]

# dispatches a full drag sequence (dragenter/dragover/drop) with a real DataTransfer
DROP_JS = """async ([selector, name, content]) => {
    const dt = new DataTransfer();
    dt.items.add(new File([content], name, {type: "text/csv"}));
    const target = document.querySelector(selector);
    for (const type of ["dragenter", "dragover", "drop"]) {
        target.dispatchEvent(new DragEvent(type, {bubbles: true, cancelable: true, dataTransfer: dt}));
    }
}"""


def drop(page, selector, name, content):
    page.evaluate(DROP_JS, [selector, name, content])


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        js_errors = []
        page.on("pageerror", lambda e: js_errors.append(str(e)))
        page.goto(BASE)
        assert page.title().startswith("DOC2MD"), page.title()

        # --- A: drag&drop прямо на зону
        drop(page, "#drop", "zone.csv", "col1;col2\n1;2\n")
        page.wait_for_selector("#result:not([hidden])", timeout=20000)
        md = page.input_value("#out")
        assert "col1" in md and "---" in md, "неожиданный markdown: " + md[:120]
        assert page.text_content("#dropTitle") == "Перетащите файл сюда", "заголовок зоны не вернулся"
        print("A) DnD на зону            OK")

        # --- B: drag&drop мимо зоны (на футер) — должен тоже принять
        page.evaluate("document.getElementById('result').hidden = true")
        drop(page, "footer", "miss.csv", "alpha;beta\n5;6\n")
        page.wait_for_selector("#result:not([hidden])", timeout=20000)
        md = page.input_value("#out")
        assert "alpha" in md, "drop мимо зоны не сработал: " + md[:120]
        print("B) DnD вне зоны (футер)   OK")

        # --- C: выбор файла кликом (input)
        page.evaluate("document.getElementById('result').hidden = true")
        page.set_input_files("#file", str(ROOT / "tests/samples/sample.csv"))
        page.wait_for_selector("#result:not([hidden])", timeout=20000)
        md = page.input_value("#out")
        assert "| id | Товар |" in md, "клик-путь сломан: " + md[:120]
        print("C) Выбор файла кликом     OK")

        # --- D: неподдерживаемый формат через drop -> понятная ошибка
        page.evaluate("document.getElementById('result').hidden = true")
        drop(page, "#drop", "bad.py", "print(1)\n")
        page.wait_for_selector("#error:not([hidden])", timeout=20000)
        err = page.text_content("#error")
        assert "не поддерживается" in err, err
        print("D) Ошибка формата         OK")

        # --- E: визуальный отклик при переносе
        page.evaluate("""() => {
            const dt = new DataTransfer();
            dt.items.add(new File(["x"], "h.csv"));
            document.dispatchEvent(new DragEvent("dragenter", {bubbles: true, dataTransfer: dt}));
        }""")
        assert page.text_content("#dropTitle") == "Отпустите файл здесь", "нет подсказки при переносе"
        assert "hover" in page.get_attribute("#drop", "class"), "зона не подсветилась"
        page.evaluate("""() => {
            const dt = new DataTransfer();
            dt.items.add(new File(["x"], "h.csv"));
            document.dispatchEvent(new DragEvent("dragleave", {bubbles: true, dataTransfer: dt}));
        }""")
        assert page.text_content("#dropTitle") == "Перетащите файл сюда", "подсказка не сбросилась"
        print("E) Отклик при переносе    OK")

        browser.close()
        if js_errors:
            print("JS-ОШИБКИ НА СТРАНИЦЕ:", js_errors)
            return 1
        print("\nВсе проверки UI пройдены.")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
