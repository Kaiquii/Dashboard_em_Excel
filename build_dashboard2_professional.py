from pathlib import Path
import re
import unicodedata

import openpyxl
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


INPUT = Path(r"C:\Users\kaiqu\Downloads\CQ Produto Acabado - Procytrat 2026.xlsm")
OUTPUT = Path("CQ Produto Acabado - Procytrat 2026 - Dashboard2-profissional-sem-simulados.xlsm")

ROWS_PER_PRODUCT = 500
HELPER_FIRST_ROW = 2
EXCLUDED = {"dashboard", "dashboard2", "menu", "padrao", "planilha1", "planilha3"}

COLORS = {
    "bg": "F3F6FA",
    "nav": "17212B",
    "nav2": "213043",
    "surface": "FFFFFF",
    "surface2": "F8FAFC",
    "line": "D7DEE8",
    "line2": "E7EDF4",
    "ink": "111827",
    "muted": "667085",
    "blue": "2563EB",
    "blue2": "EAF1FF",
    "teal": "0F766E",
    "teal2": "E7F7F5",
    "green": "16A34A",
    "green2": "E8F8EF",
    "amber": "D97706",
    "amber2": "FFF4D6",
    "rose": "BE123C",
    "rose2": "FFE4E8",
    "purple": "7C3AED",
    "purple2": "F1E9FF",
}


def normalize(value):
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text)


def excel_text(value):
    return str(value).replace('"', '""')


def quote_sheet(name):
    return "'" + name.replace("'", "''") + "'"


def product_name(sheet_name):
    return re.sub(r"^procytrat\s+", "", sheet_name, flags=re.I).strip() or sheet_name


def find_header(ws):
    best_row = None
    best_score = -1
    best_values = []
    max_scan_row = min(ws.max_row, 10)
    max_scan_col = min(ws.max_column, 35)

    for row in range(1, max_scan_row + 1):
        values = [ws.cell(row=row, column=col).value for col in range(1, max_scan_col + 1)]
        text = normalize(" | ".join(str(v) for v in values if v is not None))
        score = sum(token in text for token in ("data", "recebimento", "lote", "quantidade", "status", "aprovacao", "validade"))
        if score > best_score:
            best_row = row
            best_score = score
            best_values = values

    if best_score < 2:
        return None

    header = {"row": best_row}
    for idx, raw in enumerate(best_values, start=1):
        text = normalize(raw)
        if not text:
            continue
        if ("data" in text or "recebimento" in text) and "validade" not in text:
            header.setdefault("date", idx)
        if "lote" in text:
            header.setdefault("lot", idx)
        if "quantidade" in text:
            header.setdefault("qty", idx)
        if "status" in text or "aprovacao" in text:
            header.setdefault("status", idx)
        if "validade" in text:
            header.setdefault("validity", idx)

    return header if {"date", "lot", "qty"}.issubset(header) else None
