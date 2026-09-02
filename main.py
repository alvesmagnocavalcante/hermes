import logging
import os
from pathlib import Path

import flet as ft

from hermes_ui.app import main
from hermes_ui.telemetry import configure_telemetry

if __name__ == "__main__":
    logging.basicConfig(
        level=os.environ.get("HERMES_LOG_LEVEL", "INFO").upper(),
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
    )
    configure_telemetry()
    project_root = Path(__file__).resolve().parent
    assets_dir = project_root / "assets"

    # Em desenvolvimento, impede o Flet de confundir o executável já
    # compilado em build/windows com o cliente usado por `python main.py`.
    if Path(__file__).suffix == ".py" and (project_root / "build" / "windows").is_dir():
        os.chdir(project_root.parent)

    ft.run(main, assets_dir=str(assets_dir))
