from pathlib import Path

import openpyxl


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parents[1]
OUTPUT_DIR = PROJECT_ROOT / "outputs" / SCRIPT_DIR.name
SOURCE = OUTPUT_DIR / "CQ Produto Acabado - Procytrat 2026 - Dashboard2-profissional-sem-simulados.xlsm"
OUTPUT = OUTPUT_DIR / "CQ Produto Acabado - Procytrat 2026 - Dashboard2-base-sem-graficos-sem-simulados.xlsm"


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
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
