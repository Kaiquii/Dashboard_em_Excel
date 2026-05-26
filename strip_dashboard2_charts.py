from pathlib import Path

import openpyxl


SOURCE = Path("CQ Produto Acabado - Procytrat 2026 - Dashboard2-profissional-ajustada.xlsm")
OUTPUT = Path("CQ Produto Acabado - Procytrat 2026 - Dashboard2-base-sem-graficos.xlsm")


def main():
    wb = openpyxl.load_workbook(SOURCE, keep_vba=True)
    ws = wb["Dashboard2"]
    ws._charts = []
    ws.protection.sheet = False
    ws.protection.objects = False
    ws.protection.scenarios = False
    wb.save(OUTPUT)
    print(OUTPUT.resolve())


if __name__ == "__main__":
    main()