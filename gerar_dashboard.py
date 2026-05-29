from pathlib import Path
import json
import shutil
import subprocess
import sys


DASHBOARD_PADRAO = "CQ_Produto_Acabado_Procyfloc_2026"


def run(command):
    print(" ".join(str(item) for item in command))
    subprocess.run(command, check=True)


def main():
    root = Path(__file__).resolve().parent
    dashboard = sys.argv[1] if len(sys.argv) > 1 else DASHBOARD_PADRAO
    dashboard_dir = root / "dashboards" / dashboard
    config_path = dashboard_dir / "config.json"
    work_dir = root / ".tmp" / dashboard

    if not dashboard_dir.exists():
        raise SystemExit(f"Dashboard nao encontrado: {dashboard}")
    if not config_path.exists():
        raise SystemExit(f"Config nao encontrado: {config_path}")

    config = json.loads(config_path.read_text(encoding="utf-8"))
    input_file = root / config["entrada"]
    final_file = root / config["saidaFinal"]
    output_dir = final_file.parent
    steps = config.get("etapas") or [
        "01_build_dashboard.py",
        "02_strip_charts.py",
        "03_recreate_native_charts.ps1",
        "04_cleanup_final.ps1",
    ]

    if not input_file.exists():
        raise SystemExit(f"Planilha original nao encontrada: {input_file}")

    if work_dir.exists():
        shutil.rmtree(work_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    for old_name in [final_file.name, *config.get("limparSaida", [])]:
        old_file = output_dir / old_name
        if old_file.exists():
            old_file.unlink()

    try:
        for step in steps:
            step_path = dashboard_dir / step
            if step_path.suffix.lower() == ".py":
                run([sys.executable, str(step_path)])
            elif step_path.suffix.lower() == ".ps1":
                command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(step_path)]
                if step_path.name == "03_recreate_native_charts.ps1":
                    sources = list(work_dir.glob("*base-sem-graficos.xlsm"))
                    if sources:
                        command.extend(["-Source", str(sources[0]), "-Output", str(final_file)])
                elif step_path.name == "04_cleanup_final.ps1":
                    command.extend(["-Path", str(final_file)])
                run(command)
            else:
                raise SystemExit(f"Etapa sem suporte: {step}")
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
