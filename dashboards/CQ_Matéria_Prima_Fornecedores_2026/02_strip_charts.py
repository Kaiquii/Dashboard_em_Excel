from pathlib import Path

import openpyxl


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parents[1]
WORK_DIR = PROJECT_ROOT / ".tmp" / SCRIPT_DIR.name
SOURCE = WORK_DIR / "CQ Matéria Prima - Fornecedores 2026 - Dashboard-openpyxl.xlsm"
OUTPUT = WORK_DIR / "CQ Matéria Prima - Fornecedores 2026 - Dashboard-base-sem-graficos.xlsm"


def main():
    wb = openpyxl.load_workbook(SOURCE, keep_vba=True)
    ws = wb["Dashboard2"]
    ws._charts = []
    ws.protection.sheet = False
    ws.protection.objects = False
    ws.protection.scenarios = False
    wb.save(OUTPUT)
    print("Etapa 2 concluida.")


if __name__ == "__main__":
    main()
