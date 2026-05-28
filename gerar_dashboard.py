from pathlib import Path
import shutil
import subprocess
import sys


DASHBOARD_PADRAO = "CQ_Produto_Acabado_Procytrat_2026"


def run(command):
    print(" ".join(str(item) for item in command))
    subprocess.run(command, check=True)


def main():
    root = Path(__file__).resolve().parent
    dashboard = sys.argv[1] if len(sys.argv) > 1 else DASHBOARD_PADRAO
    dashboard_dir = root / "dashboards" / dashboard
    input_file = root / "inputs" / dashboard / "CQ Produto Acabado - Procytrat 2026.xlsm"
    work_dir = root / ".tmp" / dashboard
    output_dir = root / "outputs" / dashboard
    final_file = output_dir / "CQ Produto Acabado - Procytrat 2026 - Dashboard2-sem-simulados.xlsm"

    if not dashboard_dir.exists():
        raise SystemExit(f"Dashboard nao encontrado: {dashboard}")
    if not input_file.exists():
        raise SystemExit(f"Planilha original nao encontrada: {input_file}")

    if work_dir.exists():
        shutil.rmtree(work_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    for old_name in [
        final_file.name,
        "CQ Produto Acabado - Procytrat 2026 - Dashboard2-profissional-sem-simulados.xlsm",
        "CQ Produto Acabado - Procytrat 2026 - Dashboard2-base-sem-graficos-sem-simulados.xlsm",
    ]:
        old_file = output_dir / old_name
        if old_file.exists():
            old_file.unlink()

    try:
        run([sys.executable, str(dashboard_dir / "01_build_dashboard.py")])
        run([sys.executable, str(dashboard_dir / "02_strip_charts.py")])
        run(["powershell", "-ExecutionPolicy", "Bypass", "-File", str(dashboard_dir / "03_recreate_native_charts.ps1")])
        run(["powershell", "-ExecutionPolicy", "Bypass", "-File", str(dashboard_dir / "04_cleanup_final.ps1")])
    except Exception:
        if final_file.exists():
            final_file.unlink()
        raise
    finally:
        if work_dir.exists():
            shutil.rmtree(work_dir)

    if not final_file.exists():
        raise SystemExit(f"Arquivo final nao foi gerado: {final_file}")
    print(f"\nConcluido. Arquivo final: {final_file}")


if __name__ == "__main__":
    main()
