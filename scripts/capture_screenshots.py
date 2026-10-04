"""Genera las capturas de los mapas Folium y del dashboard de Plotly Dash para la presentación.

Requisitos previos:
    * Haber ejecutado notebooks/06_interactive_map_folium.ipynb (crea maps/*.html).
    * pip install playwright && python -m playwright install chromium

Uso:
    python scripts/capture_screenshots.py

También calcula results/dashboard.json con las cifras que se ven en el dashboard.
"""
from __future__ import annotations

import json
import os
import socket
import subprocess
import sys
import time
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
IMAGES = ROOT / "images"
RESULTS = ROOT / "results"
MAPS = ROOT / "maps"
DASH_URL = "http://127.0.0.1:8050"
VIEWPORT = {"width": 1600, "height": 900}
DASH_VIEWPORT = {"width": 1000, "height": 800}
PAYLOAD_RANGE = (2000, 6000)  # rango reducido para la captura del slider


def dashboard_stats() -> dict:
    """Cifras del dashboard calculadas sobre el mismo CSV que usa la app."""
    sys.path.insert(0, str(ROOT / "dashboard"))
    from spacex_dash_app import spacex_df as df  # noqa: E402  (reutiliza la carga de la app)

    by_site = df.groupby("Launch Site")["class"].agg(total="count", exitos="sum")
    by_site["tasa_exito"] = by_site["exitos"] / by_site["total"]
    by_site["cuota_exitos"] = by_site["exitos"] / by_site["exitos"].sum()
    bins = [0, 2000, 4000, 6000, 8000, 10000]
    labels = ["0–2k", "2k–4k", "4k–6k", "6k–8k", "8k–10k"]
    payload_bin = pd.cut(df["Payload Mass (kg)"], bins=bins, labels=labels, include_lowest=True)
    by_payload = df.groupby(payload_bin, observed=False)["class"].agg(["count", "mean"])
    by_booster = df.groupby("Booster Version Category")["class"].agg(["count", "mean"])
    stats = {
        "by_site": {
            site: {k: round(float(v), 4) for k, v in row.items()} for site, row in by_site.iterrows()
        },
        "top_share_site": by_site["cuota_exitos"].idxmax(),
        "top_rate_site": by_site["tasa_exito"].idxmax(),
        "success_rate_by_payload_bin": {
            str(k): {"launches": int(r["count"]), "rate": round(float(r["mean"]), 4) if r["count"] else None}
            for k, r in by_payload.iterrows()
        },
        "success_rate_by_booster": {
            str(k): {"launches": int(r["count"]), "rate": round(float(r["mean"]), 4)}
            for k, r in by_booster.iterrows()
        },
    }
    RESULTS.mkdir(exist_ok=True)
    (RESULTS / "dashboard.json").write_text(json.dumps(stats, indent=2, ensure_ascii=False))
    print("escrito results/dashboard.json")
    return stats


def wait_for_port(port: int, timeout: float = 60) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", port)) == 0:
                return
        time.sleep(0.5)
    raise TimeoutError(f"El dashboard no abrió el puerto {port}")


def capture_maps(page) -> None:
    for html_file, png in [
        ("01_launch_sites.html", "map_launch_sites.png"),
        ("02_launch_outcomes.html", "map_launch_outcomes.png"),
        ("03_proximities.html", "map_proximities.png"),
    ]:
        path = MAPS / html_file
        if not path.exists():
            print(f"[aviso] falta {path.relative_to(ROOT)}: ejecuta el notebook 06")
            continue
        page.goto(path.as_uri())
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2500)  # margen para que terminen de pintarse las teselas
        page.screenshot(path=str(IMAGES / png))
        print("guardada", IMAGES / png)


def wait_for_graphs(page) -> None:
    page.wait_for_selector("#success-pie-chart .main-svg", timeout=30000)
    page.wait_for_selector("#success-payload-scatter-chart .main-svg", timeout=30000)
    page.wait_for_timeout(1500)


def clear_hover(page) -> None:
    """Aparta el ratón para que no queden tooltips en la captura."""
    page.mouse.move(2, 2, steps=5)
    page.evaluate(
        "document.querySelectorAll('.js-plotly-plot')"
        ".forEach(g => window.Plotly && window.Plotly.Fx.unhover(g))"
    )
    page.wait_for_timeout(500)


def select_site(page, label: str) -> None:
    """Elige una opción del desplegable de sitios (Dash 4)."""
    page.locator("#site-dropdown").click()
    page.locator("[role=option]", has_text=label).first.click()
    page.wait_for_timeout(2000)
    clear_hover(page)


def set_payload_range(page, payload_range: tuple[int, int]) -> None:
    """Fija el rango del RangeSlider escribiendo en sus campos numéricos (Dash 4)."""
    for selector, value in zip(
        (".dash-range-slider-min-input", ".dash-range-slider-max-input"), payload_range
    ):
        field = page.locator(f"#payload-slider {selector}")
        field.fill(str(value))
        field.press("Enter")
    page.wait_for_timeout(2000)
    clear_hover(page)


def capture_dashboard(page, top_rate_site: str) -> None:
    app = subprocess.Popen([sys.executable, str(ROOT / "dashboard" / "spacex_dash_app.py")], cwd=ROOT)
    try:
        wait_for_port(8050)
        page.goto(DASH_URL)
        wait_for_graphs(page)
        clear_hover(page)

        # 1. Pastel con todos los sitios
        page.locator("#success-pie-chart").screenshot(path=str(IMAGES / "dash_pie_all_sites.png"))
        # 2. Dispersión con todo el rango de carga
        page.locator("#success-payload-scatter-chart").screenshot(path=str(IMAGES / "dash_scatter_all.png"))

        # 3. Pastel del sitio con mayor tasa de éxito
        select_site(page, top_rate_site)
        page.locator("#success-pie-chart").screenshot(path=str(IMAGES / "dash_pie_top_site.png"))

        # 4. Dispersión de todos los sitios con un rango de carga reducido
        select_site(page, "Todos los sitios")
        set_payload_range(page, PAYLOAD_RANGE)
        page.locator("#success-payload-scatter-chart").screenshot(path=str(IMAGES / "dash_scatter_range.png"))
        print("capturas del dashboard guardadas en images/")
    finally:
        app.terminate()
        app.wait(timeout=10)


def main() -> None:
    from playwright.sync_api import sync_playwright

    IMAGES.mkdir(exist_ok=True)
    stats = dashboard_stats()
    with sync_playwright() as p:
        # PLAYWRIGHT_CHROMIUM_EXECUTABLE permite usar un Chromium ya instalado en lugar del de Playwright
        browser = p.chromium.launch(executable_path=os.environ.get("PLAYWRIGHT_CHROMIUM_EXECUTABLE") or None)
        page = browser.new_page(viewport=VIEWPORT, device_scale_factor=1.5)
        capture_maps(page)
        # Ventana más estrecha para que los gráficos del dashboard queden menos apaisados
        dash_page = browser.new_page(viewport=DASH_VIEWPORT, device_scale_factor=2)
        capture_dashboard(dash_page, stats["top_rate_site"])
        browser.close()


if __name__ == "__main__":
    main()
