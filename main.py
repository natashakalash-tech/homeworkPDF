#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
CLI-скрипт для генерации PDF-счетов из CSV.
Читает data/products.csv, подставляет данные в шаблон, сохраняет PDF в output/.
PDF генерируется через ReportLab с шрифтом Arial — кириллица отображается корректно.

Установка: pip install -r requirements.txt
Запуск: python main.py
"""

import csv
import os
import sys

# Корректный вывод кириллицы в консоли Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle

# Корневая папка проекта
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_CSV = os.path.join(BASE_DIR, "data", "products.csv")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# Имя зарегистрированного шрифта с кириллицей
FONT_NAME = "ArialCyr"
_font_registered = False


def register_arial() -> bool:
    """Регистрирует Arial из Windows Fonts для кириллицы в PDF. Вызывать до построения PDF."""
    global _font_registered
    if _font_registered:
        return True
    if sys.platform != "win32":
        return False
    path = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts", "arial.ttf")
    if not os.path.isfile(path):
        return False
    try:
        pdfmetrics.registerFont(TTFont(FONT_NAME, path))
        _font_registered = True
        return True
    except Exception:
        return False


def safe_filename(product: str) -> str:
    """Убирает символы, недопустимые в имени файла Windows."""
    invalid = '<>:"/\\|?*'
    for c in invalid:
        product = product.replace(c, "_")
    return product.strip() or "invoice"


def build_invoice_pdf(output_path: str, rows: list[tuple[str, str, str, str]]) -> bool:
    """Создаёт один PDF-счёт с таблицей: заголовок + все строки из rows (product, price, qty, total)."""
    if not register_arial():
        return False
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=20 * mm,
        leftMargin=20 * mm,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
    )
    style_heading = ParagraphStyle(
        name="InvoiceHeading",
        fontName=FONT_NAME,
        fontSize=16,
        spaceAfter=12,
    )
    style_cell = ParagraphStyle(
        name="Cell",
        fontName=FONT_NAME,
        fontSize=10,
    )

    title = Paragraph("Счёт на товары", style_heading)
    # Заголовок таблицы
    data = [
        [Paragraph("Товар", style_cell), Paragraph("Цена", style_cell), Paragraph("Количество", style_cell), Paragraph("Сумма", style_cell)],
    ]
    # Строка данных для каждой записи из CSV
    for product, price, qty, total in rows:
        data.append([
            Paragraph(product, style_cell),
            Paragraph(price + " ₽", style_cell),
            Paragraph(qty + " шт.", style_cell),
            Paragraph(total + " ₽", style_cell),
        ])
    table = Table(data, colWidths=[60 * mm, 35 * mm, 40 * mm, 40 * mm])
    table.setStyle(
        TableStyle(
            [
                ("TEXTFONT", (0, 0), (-1, -1), FONT_NAME),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f2f2f2")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story = [title, table]
    try:
        doc.build(story)
        return True
    except Exception:
        return False


def main() -> None:
    if not os.path.isfile(DATA_CSV):
        print(f"Ошибка: файл не найден: {DATA_CSV}", file=sys.stderr)
        sys.exit(1)
    if not register_arial():
        print("Ошибка: не удалось загрузить шрифт Arial (C:\\Windows\\Fonts\\arial.ttf). Кириллица может отображаться некорректно.", file=sys.stderr)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(DATA_CSV, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames or "product" not in (reader.fieldnames or []):
            print("Ошибка: в CSV должны быть столбцы product, price, qty", file=sys.stderr)
            sys.exit(1)
        rows_raw = list(reader)

    # Собираем все строки для таблицы: (product, price, qty, total)
    table_rows: list[tuple[str, str, str, str]] = []
    for row in rows_raw:
        product = (row.get("product") or "").strip()
        price = (row.get("price") or "0").strip().replace(",", ".")
        qty = (row.get("qty") or "0").strip().replace(",", ".")
        if not product:
            continue
        try:
            total_val = float(price) * float(qty)
            total = f"{total_val:.2f}"
        except ValueError:
            print(f"Пропуск строки (неверные число): {row}", file=sys.stderr)
            continue
        table_rows.append((product, price, qty, total))
        print(f"Обрабатываю: {product}...")

    pdf_path = os.path.join(OUTPUT_DIR, "invoice.pdf")
    if not table_rows:
        print("Нет данных для счёта.", file=sys.stderr)
        sys.exit(1)
    if not build_invoice_pdf(pdf_path, table_rows):
        print(f"Ошибка при создании PDF: {pdf_path}", file=sys.stderr)
        sys.exit(1)
    print(f"Готово: {pdf_path}")

    if os.path.isfile(pdf_path):
        print("Открываю PDF...")
        os.startfile(pdf_path)


if __name__ == "__main__":
    main()
