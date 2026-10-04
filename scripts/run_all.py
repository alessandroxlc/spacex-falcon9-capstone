"""Ejecuta el proyecto completo de principio a fin.

    python scripts/run_all.py                    # notebooks + capturas + presentación (.pptx y PDF)
    python scripts/run_all.py --skip-screenshots # sin capturas de mapas/dashboard

Orden: notebooks 01 → 07 (con sus salidas guardadas), capturas de Folium y Dash, y presentación.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def run(*cmd: str) -> None:
    print(">>", " ".join(cmd), flush=True)
    subprocess.run([sys.executable, *cmd], check=True, cwd=ROOT)


def main() -> None:
    for notebook in sorted((ROOT / "notebooks").glob("0*.ipynb")):
        run("-m", "jupyter", "nbconvert", "--to", "notebook", "--execute", "--inplace",
            "--ExecutePreprocessor.timeout=900", str(notebook.relative_to(ROOT)))
    if "--skip-screenshots" not in sys.argv:
        run("scripts/capture_screenshots.py")
    run("presentation/generate_presentation.py", "--pdf")


if __name__ == "__main__":
    main()
